import { Router } from 'express';
import multer from 'multer';
import {
  generateOvertimeTemplateFile,
  getOvertimeEntriesByEmployee,
  getOvertimeSummaryByEmployee
} from '../controllers/overtime.controller';

const overtimeRouter = Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024
  }
});
const OVERTIME_API_ENDPOINTS = [
  'POST /api/overtime/template/xlsx (multipart: timesheet, sheet_ot_pdf)',
  'GET /api/employees/:employeeId/overtimes',
  'GET /api/employees/:employeeId/overtimes/summary'
];

overtimeRouter.post(
  '/overtime/template/xlsx',
  upload.fields([
    { name: 'timesheet', maxCount: 1 },
    { name: 'sheet_ot_pdf', maxCount: 1 }
  ]),
  generateOvertimeTemplateFile
);
overtimeRouter.get('/employees/:employeeId/overtimes', getOvertimeEntriesByEmployee);
overtimeRouter.get('/employees/:employeeId/overtimes/summary', getOvertimeSummaryByEmployee);

export { overtimeRouter, OVERTIME_API_ENDPOINTS };
