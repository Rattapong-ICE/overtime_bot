import { Schema, model, type InferSchemaType } from 'mongoose';

const holidayNaccSchema = new Schema(
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
    collection: 'holiday_nacc'
  }
);

export type HolidayNaccDocument = InferSchemaType<typeof holidayNaccSchema>;

export const HolidayNaccModel = model<HolidayNaccDocument>('HolidayNacc', holidayNaccSchema);
