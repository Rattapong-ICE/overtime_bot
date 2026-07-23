import { Router } from 'express';
import { helloRouter, HELLO_API_ENDPOINTS } from '../routes/hello.routes';
import { overtimeRouter, OVERTIME_API_ENDPOINTS } from '../routes/overtime.routes';
import { userRouter, USER_API_ENDPOINTS } from '../routes/user.routes';
import { holidayRouter, HOLIDAY_API_ENDPOINTS } from '../routes/holiday.routes';
import { holidayNaccRouter, HOLIDAY_NACC_API_ENDPOINTS } from '../routes/holiday-nacc.routes';
import { naccOvertimeRouter, NACC_OVERTIME_API_ENDPOINTS } from '../routes/nacc-overtime.routes';
import { userNaccRouter, USER_NACC_API_ENDPOINTS } from '../routes/user-nacc.routes';

const apiRouter = Router();

apiRouter.use(helloRouter);
apiRouter.use(overtimeRouter);
apiRouter.use(userRouter);
apiRouter.use(holidayRouter);
apiRouter.use(holidayNaccRouter);
apiRouter.use(naccOvertimeRouter);
apiRouter.use(userNaccRouter);

const REGISTERED_API_ENDPOINTS = [
  'GET /',
  ...HELLO_API_ENDPOINTS,
  ...OVERTIME_API_ENDPOINTS,
  ...USER_API_ENDPOINTS,
  ...HOLIDAY_API_ENDPOINTS,
  ...HOLIDAY_NACC_API_ENDPOINTS,
  ...NACC_OVERTIME_API_ENDPOINTS,
  ...USER_NACC_API_ENDPOINTS
];

export { apiRouter, REGISTERED_API_ENDPOINTS };