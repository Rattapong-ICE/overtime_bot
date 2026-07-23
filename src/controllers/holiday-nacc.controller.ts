import { type Request, type Response } from 'express';
import { z } from 'zod';
import { logger } from '../lib/logger';
import { saveHolidayNacc } from '../services/holiday-nacc.service';

const saveHolidayNaccBodySchema = z.object({
  date: z.string().regex(/^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/),
  enabled: z.boolean()
});

export async function saveHolidayNaccHandler(request: Request, response: Response): Promise<void> {
  try {
    const parsed = saveHolidayNaccBodySchema.safeParse(request.body);
    if (!parsed.success) {
      response.status(400).json({
        message: 'Invalid body. Required: date (YYYY-MM-DD), enabled (boolean).'
      });
      return;
    }

    const savedHoliday = await saveHolidayNacc(parsed.data);

    response.status(201).json({
      message: 'Holiday NACC saved successfully',
      holiday: savedHoliday
    });
  } catch (error) {
    logger.error({ error }, 'Failed to save holiday_nacc');
    response.status(400).json({ message: 'Failed to save holiday_nacc' });
  }
}
