import { Router } from 'express';
import { saveHolidayHandler } from '../controllers/holiday.controller';

const holidayRouter = Router();

const HOLIDAY_API_ENDPOINTS = ['POST /api/holidays'];

holidayRouter.post('/holidays', saveHolidayHandler);

export { holidayRouter, HOLIDAY_API_ENDPOINTS };
