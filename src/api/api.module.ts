import { Router } from 'express';
import { helloRouter, HELLO_API_ENDPOINTS } from '../routes/hello.routes';
import { overtimeRouter, OVERTIME_API_ENDPOINTS } from '../routes/overtime.routes';

const apiRouter = Router();

apiRouter.use(helloRouter);
apiRouter.use(overtimeRouter);

const REGISTERED_API_ENDPOINTS = [
  'GET /',
  ...HELLO_API_ENDPOINTS,
  ...OVERTIME_API_ENDPOINTS
];

export { apiRouter, REGISTERED_API_ENDPOINTS };