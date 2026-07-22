import { Router } from 'express';
import { createUserNaccHandler } from '../controllers/user-nacc.controller';

const userNaccRouter = Router();

const USER_NACC_API_ENDPOINTS = ['POST /api/user-nacc'];

userNaccRouter.post('/user-nacc', createUserNaccHandler);

export { userNaccRouter, USER_NACC_API_ENDPOINTS };
