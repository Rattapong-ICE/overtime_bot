import { Schema, model, type InferSchemaType } from 'mongoose';

const overtimeSchema = new Schema(
  {
    employeeId: {
      type: String,
      required: true,
      trim: true,
      index: true
    },
    employeeName: {
      type: String,
      required: true,
      trim: true
    },
    date: {
      type: Date,
      required: true,
      index: true
    },
    hours: {
      type: Number,
      required: true,
      min: 0
    },
    reason: {
      type: String,
      required: true,
      trim: true
    }
  },
  {
    timestamps: true
  }
);

export type OvertimeDocument = InferSchemaType<typeof overtimeSchema>;

export const OvertimeModel = model<OvertimeDocument>('Overtime', overtimeSchema);
