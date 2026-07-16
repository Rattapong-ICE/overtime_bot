import 'dotenv/config';
import express, { type Request, type Response } from 'express';
import { apiRouter, REGISTERED_API_ENDPOINTS } from './api/api.module';
import { connectDatabase } from './config/database';
import { logger } from './lib/logger';

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const isVercelRuntime = process.env.VERCEL === '1';
let runtimeInitPromise: Promise<void> | null = null;

app.use(express.json());

app.get('/', (_request: Request, response: Response) => {
  response.json({
    name: 'overtime_bot',
    status: 'ok',
    message: 'API is running'
  });
});

function initRuntime(): Promise<void> {
  if (!runtimeInitPromise) {
    runtimeInitPromise = (async () => {
      await connectDatabase();
      logReadyApis();
    })();
  }

  return runtimeInitPromise;
}

app.use(async (_request: Request, response: Response, next) => {
  try {
    await initRuntime();
    next();
  } catch (error) {
    logger.error({ error }, 'Failed to initialize runtime');
    response.status(500).json({
      status: 'error',
      message: 'Runtime initialization failed'
    });
  }
});

app.use('/api', apiRouter); 

function logReadyApis(): void {
  logger.info('APIs ready to use:');
  for (const endpoint of REGISTERED_API_ENDPOINTS) {
    logger.info({ endpoint }, 'API ready');
  }
}

async function bootstrap(): Promise<void> {
  try {
    await initRuntime();
    app.listen(PORT, () => {
      logger.info({ port: PORT }, `Server is running on http://localhost:${PORT}`);
    });
  } catch (error) {
    logger.error({ error }, 'Failed to start server');
    process.exit(1);
  }
}

if (!isVercelRuntime) {
  void bootstrap();
}

export default app;
