import { type Request, type Response } from 'express';
import { unlink } from 'node:fs/promises';
import { z } from 'zod';
import { logger } from '../lib/logger';
import { uploadGeneratedFileToFirebase } from '../services/firebase-storage.service';
import { generateNaccOvertimeExcel, readNaccOvertimeFile } from '../services/nacc-overtime.service';

const requestBodySchema = z.object({
  month: z.string().regex(/^\d{4}-(0[1-9]|1[0-2])$/)
});

export async function readNaccOvertimeFileHandler(request: Request, response: Response): Promise<void> {
  try {
    const parsedBody = requestBodySchema.safeParse(request.body);
    if (!parsedBody.success) {
      response.status(400).json({
        message: 'Invalid body. Required: month (YYYY-MM).'
      });
      return;
    }

    const files = request.files as {
      overtimeFile?: Express.Multer.File[];
    } | undefined;

    const overtimeFile = files?.overtimeFile?.[0];

    if (!overtimeFile) {
      response.status(400).json({
        message: 'Missing input file. Send overtimeFile (.dat, .csv, .xlsx).'
      });
      return;
    }

    const result = await readNaccOvertimeFile(overtimeFile);
    const generatedFile = await generateNaccOvertimeExcel(parsedBody.data.month, result.rows);
    const firebaseUpload = await uploadGeneratedFileToFirebase(
      generatedFile.filePath,
      `nacc-overtime/${generatedFile.fileName}`
    );

    await unlink(generatedFile.filePath).catch(() => undefined);

    const responsePayload = {
      message: 'NACC overtime file generated successfully',
      month: parsedBody.data.month,
      file: {
        sourceFileName: result.fileName,
        outputFileName: generatedFile.fileName,
        rowCount: generatedFile.rowCount,
        downloadUrl: firebaseUpload.downloadUrl
      },
      generated: 'firebase',
      firebaseUpload
    };

    response.status(200).json(responsePayload);
  } catch (error) {
    logger.error({ error }, 'Failed to parse NACC overtime file');
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    response.status(400).json({
      message: 'Failed to parse NACC overtime file',
      ...(process.env.NODE_ENV !== 'production' ? { reason: errorMessage } : {})
    });
  }
}
