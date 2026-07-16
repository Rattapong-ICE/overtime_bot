import { OvertimeModel, type OvertimeDocument } from '../models/overtime.model';

export type OvertimeSummary = {
  employeeId: string;
  employeeName: string;
  totalHours: number;
  entries: number;
};

export async function listOvertimeEntriesByEmployee(employeeId: string): Promise<OvertimeDocument[]> {
  return OvertimeModel.find({ employeeId }).sort({ date: -1 }).lean();
}

export async function summarizeOvertimeByEmployee(employeeId: string): Promise<OvertimeSummary | null> {
  const [summary] = await OvertimeModel.aggregate<OvertimeSummary>([
    {
      $match: {
        employeeId
      }
    },
    {
      $group: {
        _id: '$employeeId',
        employeeName: { $first: '$employeeName' },
        totalHours: { $sum: '$hours' },
        entries: { $sum: 1 }
      }
    },
    {
      $project: {
        _id: 0,
        employeeId: '$_id',
        employeeName: 1,
        totalHours: 1,
        entries: 1
      }
    }
  ]);

  return summary ?? null;
}
