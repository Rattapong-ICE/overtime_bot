import { Router, type Request, type Response } from 'express';

const helloRouter = Router();

const HELLO_API_ENDPOINTS = ['GET /api/hello'];

helloRouter.get('/hello', (request: Request, response: Response) => {
  const nameFromQuery = request.query.name;
  const safeName = typeof nameFromQuery === 'string' && nameFromQuery.trim() !== ''
    ? nameFromQuery.trim()
    : 'Guest';

  response.json({
    message: `Hello, ${safeName}`,
    timestamp: new Date().toISOString()
  });
});

export { helloRouter, HELLO_API_ENDPOINTS };