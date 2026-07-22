import { Router } from 'express';
import multer from 'multer';
import { readNaccOvertimeFileHandler } from '../controllers/nacc-overtime.controller';

const naccOvertimeRouter = Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024
  }
});

const NACC_OVERTIME_API_ENDPOINTS = [
  'POST /api/nacc/overtime/read (multipart: overtimeFile, body: month YYYY-MM, supports .dat/.csv/.xlsx)'
];

naccOvertimeRouter.post(
  '/nacc/overtime/read',
  upload.fields([{ name: 'overtimeFile', maxCount: 1 }]),
  readNaccOvertimeFileHandler
);

export { naccOvertimeRouter, NACC_OVERTIME_API_ENDPOINTS };
