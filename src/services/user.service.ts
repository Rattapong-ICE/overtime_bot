import { UserModel, type UserDocument } from '../models/user.model';

export type SaveUserInput = {
  username: string;
  name: string;
  company: string;
  team: string;
  attemp?: number;
};

export type SaveUserResult = {
  user: UserDocument;
  operation: 'created' | 'updated';
};

export type ConsumeUserGenFileAttemptResult =
  | {
      ok: true;
      user: UserDocument;
    }
  | {
      ok: false;
      reason: 'USER_NOT_FOUND' | 'QUOTA_EXHAUSTED';
    };

export async function saveUser(input: SaveUserInput): Promise<SaveUserResult> {
  const existingUser = await UserModel.findOne({ username: input.username }).lean();

  if (existingUser) {
    const updatePayload: {
      name: string;
      company: string;
      team: string;
      attemp?: number;
    } = {
      name: input.name,
      company: input.company,
      team: input.team
    };

    if (typeof input.attemp === 'number') {
      updatePayload.attemp = input.attemp;
    }

    const updatedUser = await UserModel.findOneAndUpdate(
      { username: input.username },
      {
        $set: updatePayload
      },
      {
        new: true,
        runValidators: true
      }
    ).lean();

    if (!updatedUser) {
      throw new Error('Failed to update user');
    }

    return {
      user: updatedUser,
      operation: 'updated'
    };
  }

  await UserModel.create({
    username: input.username,
    name: input.name,
    company: input.company,
    team: input.team,
    attemp: input.attemp ?? 3
  });

  const createdUser = await UserModel.findOne({ username: input.username }).lean();

  if (!createdUser) {
    throw new Error('Failed to create user');
  }

  return {
    user: createdUser,
    operation: 'created'
  };
}

export async function findUserByUsername(username: string): Promise<UserDocument | null> {
  return UserModel.findOne({ username }).lean();
}

export async function consumeUserGenFileAttempt(username: string): Promise<ConsumeUserGenFileAttemptResult> {
  let updatedUser = await UserModel.findOneAndUpdate(
    {
      username,
      attemp: { $gt: 0 }
    },
    {
      $inc: { attemp: -1 }
    },
    {
      new: true,
      runValidators: true
    }
  ).lean();

  // Backward compatibility for legacy documents that may not have attemp.
  if (!updatedUser) {
    updatedUser = await UserModel.findOneAndUpdate(
      {
        username,
        $or: [{ attemp: { $exists: false } }, { attemp: null }]
      },
      {
        $set: { attemp: 2 }
      },
      {
        new: true,
        runValidators: true
      }
    ).lean();
  }

  if (updatedUser) {
    return {
      ok: true,
      user: updatedUser
    };
  }

  const existingUser = await UserModel.findOne({ username }).lean();
  if (!existingUser) {
    return {
      ok: false,
      reason: 'USER_NOT_FOUND'
    };
  }

  return {
    ok: false,
    reason: 'QUOTA_EXHAUSTED'
  };
}
