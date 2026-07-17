import { Router } from 'express';
import { saveUserHandler } from '../controllers/user.controller';

const userRouter = Router();

const USER_API_ENDPOINTS = ['POST /api/users'];

userRouter.post('/users', saveUserHandler);

export { userRouter, USER_API_ENDPOINTS };
