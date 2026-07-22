import { Schema, model, type InferSchemaType } from 'mongoose';

const userNaccSchema = new Schema(
  {
    employee_id: {
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
    position: {
      type: String,
      required: true,
      trim: true
    }
  },
  {
    timestamps: true,
    collection: 'user_nacc'
  }
);

export type UserNaccDocument = InferSchemaType<typeof userNaccSchema>;

export const UserNaccModel = model<UserNaccDocument>('UserNacc', userNaccSchema);
