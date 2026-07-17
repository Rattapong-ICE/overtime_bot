import { HolidayModel, type HolidayDocument } from '../models/holiday.model';

export type SaveHolidayInput = {
  date: string;
  enabled: boolean;
};

export async function saveHoliday(input: SaveHolidayInput): Promise<HolidayDocument> {
  const holiday = await HolidayModel.findOneAndUpdate(
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
    throw new Error('Failed to save holiday');
  }

  return holiday;
}

export async function listEnabledHolidaysByMonth(targetMonth: string): Promise<HolidayDocument[]> {
  const monthPrefix = `${targetMonth}-`;

  return HolidayModel.find({
    enabled: true,
    date: { $regex: `^${monthPrefix}` }
  }).lean();
}
