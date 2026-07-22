import { type Request, type Response } from 'express';
import { z } from 'zod';
import { logger } from '../lib/logger';
import { createUserNacc } from '../services/user-nacc.service';

const createUserNaccBodySchema = z.object({
  employee_id: z.string().trim().min(1),
  name: z.string().trim().min(1),
  position: z.string().trim().min(1)
});

export async function createUserNaccHandler(request: Request, response: Response): Promise<void> {
  try {
    const parsed = createUserNaccBodySchema.safeParse(request.body);
    if (!parsed.success) {
      response.status(400).json({
        message: 'Invalid body. Required: employee_id, name, position (string).'
      });
      return;
    }

    const userNacc = await createUserNacc(parsed.data);

    response.status(201).json({
      message: 'user_nacc created successfully',
      user: userNacc
    });
  } catch (error) {
    const maybeMongoError = error as { code?: number };
    if (maybeMongoError.code === 11000) {
      response.status(409).json({
        message: 'employee_id already exists'
      });
      return;
    }

    logger.error({ error }, 'Failed to create user_nacc');
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    response.status(400).json({
      message: 'Failed to create user_nacc',
      ...(process.env.NODE_ENV !== 'production' ? { reason: errorMessage } : {})
    });
  }
}
