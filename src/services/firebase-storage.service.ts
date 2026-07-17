import { existsSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { cert, getApps, initializeApp, type ServiceAccount } from 'firebase-admin/app';
import { getStorage } from 'firebase-admin/storage';

const DEFAULT_SERVICE_ACCOUNT_FILE = 'firebase.json';

type FirebaseUploadResult = {
  storagePath: string;
  gsUri: string;
  downloadUrl: string;
};

type ServiceAccountWithProjectId = ServiceAccount & {
  project_id?: string;
};

let cachedServiceAccount: ServiceAccountWithProjectId | null = null;

function getServiceAccountPath(): string {
  const configuredPath = process.env.FIREBASE_SERVICE_ACCOUNT_PATH?.trim();
  if (configuredPath) {
    return path.isAbsolute(configuredPath)
      ? configuredPath
      : path.join(process.cwd(), configuredPath);
  }

  return path.join(process.cwd(), DEFAULT_SERVICE_ACCOUNT_FILE);
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
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

  const serviceAccountPath = getServiceAccountPath();
  if (!existsSync(serviceAccountPath)) {
    throw new Error(`Firebase service account file not found: ${serviceAccountPath}`);
  }

  cachedServiceAccount = require(serviceAccountPath) as ServiceAccountWithProjectId;
  return cachedServiceAccount;
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
        gsUri: "",
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
