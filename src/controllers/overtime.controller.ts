import { type Request, type Response } from 'express';
import { z } from 'zod';
import {
  listOvertimeEntriesByEmployee,
  summarizeOvertimeByEmployee
} from '../services/overtime.service';
import { logger } from '../lib/logger';

const employeeParamsSchema = z.object({
  employeeId: z.string().trim().min(1)
});

function parseEmployeeParams(request: Request): { employeeId: string } {
  const parsed = employeeParamsSchema.safeParse(request.params);

  if (!parsed.success) {
    throw new Error('Invalid employeeId parameter');
  }

  return parsed.data;
}

export async function getOvertimeEntriesByEmployee(request: Request, response: Response): Promise<void> {
  try {
    const { employeeId } = parseEmployeeParams(request);
    const entries = await listOvertimeEntriesByEmployee(employeeId);

    response.json({
      employeeId,
      count: entries.length,
      entries
    });
  } catch (error) {
    logger.error({ error }, 'Failed to fetch overtime entries');
    response.status(400).json({ message: 'Invalid request' });
  }
}

export async function getOvertimeSummaryByEmployee(request: Request, response: Response): Promise<void> {
  try {
    const { employeeId } = parseEmployeeParams(request);
    const summary = await summarizeOvertimeByEmployee(employeeId);

    if (!summary) {
      response.status(404).json({ message: 'Overtime summary not found' });
      return;
    }

    response.json(summary);
  } catch (error) {
    logger.error({ error }, 'Failed to fetch overtime summary');
    response.status(400).json({ message: 'Invalid request' });
  }
}
