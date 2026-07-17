import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { cert, getApps, initializeApp, type ServiceAccount } from 'firebase-admin/app';
import { getStorage } from 'firebase-admin/storage';

type FirebaseUploadResult = {
  storagePath: string;
  downloadUrl: string;
};

type ServiceAccountWithProjectId = ServiceAccount & {
  project_id?: string;
  private_key?: string;
};

let cachedServiceAccount: ServiceAccountWithProjectId | null = null;

function normalizeServiceAccount(serviceAccount: ServiceAccountWithProjectId): ServiceAccountWithProjectId {
  if (typeof serviceAccount.private_key === 'string' && !serviceAccount.privateKey) {
    serviceAccount.privateKey = serviceAccount.private_key;
  }

  if (typeof serviceAccount.privateKey === 'string') {
    serviceAccount.privateKey = serviceAccount.privateKey.replaceAll(String.raw`\n`, '\n');
  }

  return serviceAccount;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === 'string') {
    return error;
  }

  if (error && typeof error === 'object') {
    try {
      return JSON.stringify(error);
    } catch {
      return 'Unknown object error';
    }
  }

  return 'Unknown error';
}

function getServiceAccountFromEnv(): ServiceAccountWithProjectId | null {
  const oneLineJson = process.env.FIREBASE_JSON?.trim();

  if (!oneLineJson) {
    return null;
  }

  try {
    const parsed = JSON.parse(oneLineJson) as ServiceAccountWithProjectId;
    return normalizeServiceAccount(parsed);
  } catch (error) {
    throw new Error(
      'Invalid Firebase JSON in env. '
        + 'Expected one-line JSON in FIREBASE_JSON. '
        + `Details: ${getErrorMessage(error)}`,
      { cause: error }
    );
  }
}

function shouldTryNextBucket(error: unknown): boolean {
  const message = getErrorMessage(error).toLowerCase();
  return message.includes('bucket does not exist')
    || message.includes('notfound')
    || message.includes('no such bucket');
}

function normalizeBucketName(bucket: string): string {
  let normalized = bucket.trim();

  if (normalized.startsWith('gs://')) {
    normalized = normalized.slice('gs://'.length);
  }

  normalized = normalized.replace(/^https?:\/\/[^/]+\//, '');
  normalized = normalized.replace(/^b\//, '');
  normalized = normalized.replace(/\/o\/?$/, '');

  const firstSegment = normalized.split('/')[0] ?? '';
  return firstSegment.trim();
}

function getBucketCandidates(serviceAccount: ServiceAccountWithProjectId): string[] {
  const configuredBucket = process.env.FIREBASE_STORAGE_BUCKET?.trim();
  if (configuredBucket) {
    const normalizedBucket = normalizeBucketName(configuredBucket);
    if (!normalizedBucket) {
      throw new Error('FIREBASE_STORAGE_BUCKET is set but invalid. Example valid values: overtime-gen-file.appspot.com or overtime-gen-file.firebasestorage.app');
    }

    return [normalizedBucket];
  }

  const projectId = serviceAccount.project_id ?? serviceAccount.projectId;
  if (!projectId) {
    throw new Error('Cannot resolve Firebase Storage bucket: missing project_id in service account and FIREBASE_STORAGE_BUCKET is not set.');
  }

  return [`${projectId}.firebasestorage.app`, `${projectId}.appspot.com`];
}

function getFirebaseServiceAccount(): ServiceAccountWithProjectId {
  if (cachedServiceAccount) {
    return cachedServiceAccount;
  }

  const envServiceAccount = getServiceAccountFromEnv();
  if (envServiceAccount) {
    cachedServiceAccount = envServiceAccount;
    return cachedServiceAccount;
  }

  throw new Error('Missing FIREBASE_JSON in environment. Expected one-line Firebase service account JSON.');
}

function initializeFirebaseApp(): void {
  const serviceAccount = getFirebaseServiceAccount();
  if (!getApps().length) {
    initializeApp({
      credential: cert(serviceAccount)
    });
  }
}

export async function uploadGeneratedFileToFirebase(localFilePath: string, storagePath?: string): Promise<FirebaseUploadResult> {
  initializeFirebaseApp();

  const serviceAccount = getFirebaseServiceAccount();
  const bucketCandidates = getBucketCandidates(serviceAccount);

  const targetPath = storagePath ?? path.basename(localFilePath);

  let lastError: unknown;
  for (const bucketName of bucketCandidates) {
    try {
      const bucket = getStorage().bucket(bucketName);
      const downloadToken = randomUUID();

      await bucket.upload(localFilePath, {
        destination: targetPath,
        resumable: false,
        metadata: {
          metadata: {
            firebaseStorageDownloadTokens: downloadToken
          }
        }
      });

      const encodedPath = encodeURIComponent(targetPath);
      const downloadUrl = `https://firebasestorage.googleapis.com/v0/b/${bucketName}/o/${encodedPath}?alt=media&token=${downloadToken}`;

      return {
        storagePath: targetPath,
        downloadUrl
      };
    } catch (error) {
      lastError = error;
      if (!shouldTryNextBucket(error)) {
        throw error;
      }
    }
  }

  const projectId = serviceAccount.project_id ?? serviceAccount.projectId ?? 'unknown-project';
  throw new Error(
    `Firebase upload failed: no valid bucket found for project '${projectId}'. `
      + `Tried: ${bucketCandidates.join(', ')}. `
      + `Set FIREBASE_STORAGE_BUCKET to the correct bucket name. Last error: ${getErrorMessage(lastError)}`
  );
}
