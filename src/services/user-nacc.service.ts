import { UserNaccModel, type UserNaccDocument } from '../models/user-nacc.model';

export type CreateUserNaccInput = {
  employee_id: string;
  name: string;
  position: string;
};

export async function createUserNacc(input: CreateUserNaccInput): Promise<UserNaccDocument> {
  const created = await UserNaccModel.create({
    employee_id: input.employee_id,
    name: input.name,
    position: input.position
  });

  return created.toObject();
}
