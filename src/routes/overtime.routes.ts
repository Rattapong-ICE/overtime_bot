import { Router } from 'express';
import {
  getOvertimeEntriesByEmployee,
  getOvertimeSummaryByEmployee
} from '../controllers/overtime.controller';

const overtimeRouter = Router();
const OVERTIME_API_ENDPOINTS = [
  'GET /api/employees/:employeeId/overtimes',
  'GET /api/employees/:employeeId/overtimes/summary'
];

overtimeRouter.get('/employees/:employeeId/overtimes', getOvertimeEntriesByEmployee);
overtimeRouter.get('/employees/:employeeId/overtimes/summary', getOvertimeSummaryByEmployee);

export { overtimeRouter, OVERTIME_API_ENDPOINTS };
