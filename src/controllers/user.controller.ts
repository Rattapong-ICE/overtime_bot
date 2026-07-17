import { type Request, type Response } from 'express';
import { z } from 'zod';
import { logger } from '../lib/logger';
import { saveUser } from '../services/user.service';

const saveUserBodySchema = z.object({
  username: z.string().trim().min(1),
  name: z.string().trim().min(1),
  company: z.string().trim().min(1),
  team: z.string().trim().min(1)
});

export async function saveUserHandler(request: Request, response: Response): Promise<void> {
  try {
    const parsed = saveUserBodySchema.safeParse(request.body);
    if (!parsed.success) {
      response.status(400).json({
        message: 'Invalid body. Required: username, name, company, team (string).'
      });
      return;
    }

    const savedUser = await saveUser(parsed.data);

    response.status(201).json({
      message: 'User saved successfully',
      user: savedUser
    });
  } catch (error) {
    logger.error({ error }, 'Failed to save user');
    response.status(400).json({ message: 'Failed to save user' });
  }
}
