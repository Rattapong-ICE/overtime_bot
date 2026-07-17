import { UserModel, type UserDocument } from '../models/user.model';

export type SaveUserInput = {
  username: string;
  name: string;
  company: string;
  team: string;
};

export async function saveUser(input: SaveUserInput): Promise<UserDocument> {
  const user = await UserModel.findOneAndUpdate(
    { username: input.username },
    {
      $set: {
        name: input.name,
        company: input.company,
        team: input.team
      }
    },
    {
      upsert: true,
      new: true,
      setDefaultsOnInsert: true
    }
  ).lean();

  if (!user) {
    throw new Error('Failed to save user');
  }

  return user;
}

export async function findUserByUsername(username: string): Promise<UserDocument | null> {
  return UserModel.findOne({ username }).lean();
}
