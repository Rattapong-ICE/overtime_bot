import { type Request, type Response } from 'express';
import { unlink } from 'node:fs/promises';
import { z } from 'zod';
import {
  listOvertimeEntriesByEmployee,
  summarizeOvertimeByEmployee
} from '../services/overtime.service';
import { generateOvertimeTemplateXlsx as buildOvertimeTemplateXlsx } from '../services/overtime-template.service';
import { uploadGeneratedFileToFirebase } from '../services/firebase-storage.service';
import { readUploadedOvertimeFiles } from '../services/overtime-upload-reader.service';
import { findUserByUsername } from '../services/user.service';
import { listEnabledHolidaysByMonth } from '../services/holiday.service';
import { logger } from '../lib/logger';

const employeeParamsSchema = z.object({
  employeeId: z.string().trim().min(1)
});

const templateBodySchema = z.object({
  date: z.string().regex(/^\d{4}-(0[1-9]|1[0-2])$/),
  username: z.string().trim().min(1),
  text_ot: z.string().optional()
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

export async function generateOvertimeTemplateFile(request: Request, response: Response): Promise<void> {
  try {
    const files = request.files as {
      timesheet?: Express.Multer.File[];
    } | undefined;

    const timesheetFile = files?.timesheet?.[0];

    const parsedBody = templateBodySchema.safeParse(request.body);
    if (!parsedBody.success) {
      response.status(400).json({
        message: 'Invalid body. Required: date (YYYY-MM), username. Optional: text_ot.'
      });
      return;
    }

    const username = parsedBody.data.username.trim();
    const textOt = parsedBody.data.text_ot?.trim() ?? '';

    if (!timesheetFile) {
      response.status(400).json({
        message: 'Missing input. Send timesheet (.csv/.xlsx) and body.username. body.text_ot is optional.'
      });
      return;
    }

    const userProfile = await findUserByUsername(username);
    if (!userProfile) {
      response.status(404).json({
        message: `User not found for username: ${username}`
      });
      return;
    }

    const holidays = await listEnabledHolidaysByMonth(parsedBody.data.date);
    const holidayDates = holidays.map((holiday) => holiday.date);

    const data = await readUploadedOvertimeFiles(timesheetFile, textOt);
    
    logger.info(
      {
        textOtSource: data.textOt.fileName,
        textOtRowCount: data.textOt.tableRows.length,
        textOtRows: data.textOt.tableRows
      },
      'text_ot parsed rows'
    );
    logger.info(
      {
        timesheetFileName: data.timesheet.fileName,
        timesheetCount: data.timesheet.rows.length
      },
      'Timesheet data loaded for template generation'
    );
    const output = await buildOvertimeTemplateXlsx(
      parsedBody.data.date,
      data.timesheet.rows,
      data.textOt.tableRows,
      {
        username: userProfile.username,
        name: userProfile.name,
        company: userProfile.company,
        team: userProfile.team
      },
      holidayDates
    );

    const nodeEnv = (process.env.NODE_ENV ?? '').toLowerCase();
    const isProd = nodeEnv === 'prod' || nodeEnv === 'production';

    if (isProd) {
      const firebaseUpload = await uploadGeneratedFileToFirebase(
        output.filePath,
        `overtime/${output.fileName}`
      );

      await unlink(output.filePath).catch(() => undefined);

      response.status(201).json({
        message: 'Overtime template generated successfully',
        generated: 'firebase',
        firebaseUpload
      });
      return;
    }

    response.status(201).json({
      message: 'Overtime template generated successfully',
      generated: 'local',
      output
    });
  } catch (error) {
    logger.error({ error }, 'Failed to generate overtime template');
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    response.status(400).json({
      message: 'Failed to generate overtime template.',
      ...(process.env.NODE_ENV !== 'production' ? { reason: errorMessage } : {})
    });
  }
}
