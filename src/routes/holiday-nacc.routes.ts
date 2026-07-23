import { Router } from 'express';
import { saveHolidayNaccHandler } from '../controllers/holiday-nacc.controller';

const holidayNaccRouter = Router();

const HOLIDAY_NACC_API_ENDPOINTS = ['POST /api/holidays-nacc'];

holidayNaccRouter.post('/holidays-nacc', saveHolidayNaccHandler);

export { holidayNaccRouter, HOLIDAY_NACC_API_ENDPOINTS };
