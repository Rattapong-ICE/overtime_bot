import { type Request, type Response } from 'express';
import { z } from 'zod';
import {
  listOvertimeEntriesByEmployee,
  summarizeOvertimeByEmployee
} from '../services/overtime.service';
import { generateOvertimeTemplateXlsx as buildOvertimeTemplateXlsx } from '../services/overtime-template.service';
import { uploadGeneratedFileToFirebase } from '../services/firebase-storage.service';
import { readUploadedOvertimeFiles } from '../services/overtime-upload-reader.service';
import { logger } from '../lib/logger';

const employeeParamsSchema = z.object({
  employeeId: z.string().trim().min(1)
});

const templateBodySchema = z.object({
  date: z.string().regex(/^\d{4}-(0[1-9]|1[0-2])$/)
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
      sheet_ot_pdf?: Express.Multer.File[];
    } | undefined;

    const timesheetFile = files?.timesheet?.[0];
    const sheetOtPdfFile = files?.sheet_ot_pdf?.[0];

    if (!timesheetFile || !sheetOtPdfFile) {
      response.status(400).json({
        message: 'Missing files. Send timesheet (.csv/.xlsx) and sheet_ot_pdf in multipart/form-data.'
      });
      return;
    }

    const parsedBody = templateBodySchema.safeParse(request.body);
    if (!parsedBody.success) {
      response.status(400).json({
        message: 'Invalid body.date. Use format YYYY-MM, for example 2026-06.'
      });
      return;
    }

    const data = await readUploadedOvertimeFiles(timesheetFile, sheetOtPdfFile);
    // logger.info(
    //   {
    //     timesheetFileName: data.timesheet.fileName,
    //     timesheetCount: data.timesheet.rows.length
    //   },
    //   'Timesheet data loaded for template generation'
    // );

    // logger.info(
    //   {
    //     pdfFileName: data.sheetOtPdf.fileName,
    //     pdfRowCount: data.sheetOtPdf.tableRows.length,
    //     pdfRows: data.sheetOtPdf.tableRows
    //   },
    //   'Parsed sheet_ot_pdf table rows'
    // );
    // logger.info(
    //   {
    //     timesheetSample: data.timesheet.rows[0] ?? null,
    //     timesheetCount: data.timesheet.rows.length,
    //     timesheetJson: data.timesheet.rows,
    //     timesheetFileName: data.timesheet.fileName,
    //     timesheetFileType: data.timesheet.fileType
    //   },
    //   'Timesheet parsed to JSON successfully'
    // );
    // logger.info({ sheetOtPdfText: data.sheetOtPdf.text }, 'sheet_ot_pdf text extracted successfully');

    const output = await buildOvertimeTemplateXlsx(
      parsedBody.data.date,
      data.timesheet.rows,
      data.sheetOtPdf.tableRows
    );

    const firebaseUpload = await uploadGeneratedFileToFirebase(
      output.filePath,
      `overtime/${output.fileName}`
    );

    response.status(201).json({
      message: 'Overtime template generated successfully',
      //output,
      firebaseUpload
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
