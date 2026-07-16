import mongoose from 'mongoose';
import { logger } from '../lib/logger';

const DATABASE_SERVER_SELECTION_TIMEOUT_MS = 5000;

export async function connectDatabase(): Promise<void> {
  const mongoUri = process.env.MONGO_URI;

  if (!mongoUri) {
    logger.warn('MONGO_URI is not set. API will run without database connection.');
    return;
  }

  await mongoose.connect(mongoUri, {
    serverSelectionTimeoutMS: DATABASE_SERVER_SELECTION_TIMEOUT_MS
  });

  logger.info('Connected to MongoDB');
}
