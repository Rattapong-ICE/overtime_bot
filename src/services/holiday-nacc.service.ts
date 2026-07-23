import { HolidayNaccModel, type HolidayNaccDocument } from '../models/holiday-nacc.model';

export type SaveHolidayNaccInput = {
  date: string;
  enabled: boolean;
};

export async function saveHolidayNacc(input: SaveHolidayNaccInput): Promise<HolidayNaccDocument> {
  const holiday = await HolidayNaccModel.findOneAndUpdate(
    { date: input.date },
    {
      $set: {
        enabled: input.enabled
      }
    },
    {
      upsert: true,
      new: true,
      setDefaultsOnInsert: true
    }
  ).lean();

  if (!holiday) {
    throw new Error('Failed to save holiday_nacc');
  }

  return holiday;
}
