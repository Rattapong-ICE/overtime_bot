import 'dotenv/config';
import express, { type Request, type Response } from 'express';
import { apiRouter, REGISTERED_API_ENDPOINTS } from './api/api.module';
import { connectDatabase } from './config/database';
import { logger } from './lib/logger';

const app = express();
const PORT = Number(process.env.PORT) || 3000;

app.use(express.json());

app.get('/', (_request: Request, response: Response) => {
  response.json({
    name: 'overtime_bot',
    status: 'ok',
    message: 'API is running'
  });
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
    await connectDatabase();
    app.listen(PORT, () => {
      logger.info({ port: PORT }, `Server is running on http://localhost:${PORT}`);
      logReadyApis();
    });
  } catch (error) {
    logger.error({ error }, 'Failed to start server');
    process.exit(1);
  }
}

void bootstrap();
