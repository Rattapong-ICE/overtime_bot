import { type Request, type Response } from 'express';
import { z } from 'zod';
import { logger } from '../lib/logger';
import { saveUser } from '../services/user.service';

const saveUserBodySchema = z.object({
  username: z.string().trim().min(1),
  name: z.string().trim().min(1),
  company: z.string().trim().min(1),
  team: z.string().trim().min(1),
  attemp: z.number().int().min(0).optional()
});

export async function saveUserHandler(request: Request, response: Response): Promise<void> {
  try {
    const parsed = saveUserBodySchema.safeParse(request.body);
    if (!parsed.success) {
      response.status(400).json({
        message: 'Invalid body. Required: username, name, company, team (string). Optional: attemp (number, min 0).'
      });
      return;
    }

    const result = await saveUser(parsed.data);

    response.status(result.operation === 'created' ? 201 : 200).json({
      message: result.operation === 'created' ? 'User created successfully' : 'User updated successfully',
      user: result.user,
      operation: result.operation
    });
  } catch (error) {
    logger.error({ error }, 'Failed to save user');
    response.status(400).json({ message: 'Failed to save user' });
  }
}
