import { Schema, model, type InferSchemaType } from 'mongoose';

const holidaySchema = new Schema(
  {
    date: {
      type: String,
      required: true,
      trim: true,
      unique: true,
      index: true
    },
    enabled: {
      type: Boolean,
      required: true,
      default: true
    }
  },
  {
    timestamps: true,
    collection: 'holidays'
  }
);

export type HolidayDocument = InferSchemaType<typeof holidaySchema>;

export const HolidayModel = model<HolidayDocument>('Holiday', holidaySchema);
