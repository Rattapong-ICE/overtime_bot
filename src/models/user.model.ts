import { Schema, model, type InferSchemaType } from 'mongoose';

const userSchema = new Schema(
  {
    username: {
      type: String,
      required: true,
      trim: true,
      unique: true,
      index: true
    },
    name: {
      type: String,
      required: true,
      trim: true
    },
    company: {
      type: String,
      required: true,
      trim: true
    },
    team: {
      type: String,
      required: true,
      trim: true
    }
  },
  {
    timestamps: true,
    collection: 'users'
  }
);

export type UserDocument = InferSchemaType<typeof userSchema>;

export const UserModel = model<UserDocument>('User', userSchema);
