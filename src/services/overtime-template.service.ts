import { mkdir } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';

const THAI_TIMEZONE = 'Asia/Bangkok';
const FONT_NAME = 'TH Sarabun New';
const HEADER_BLUE = 'FF2C61D6';
const SUBHEADER_ORANGE = 'FFF0C38A';
const ROW_WEEKEND_GRAY = 'FFD6D6D6';
const SUMMARY_GRAY = 'FFE7E7E7';
const LUNCH_BREAK_MINUTES = 60;

type TemplateGenerationResult = {
  fileName: string;
  filePath: string;
  daysInMonth: number;
  month: string;
};

type TimesheetInputRow = Record<string, string>;

type PdfTableInputRow = {
  day: string;
  from: string;
  to: string;
  hours: string;
};

type DailyWorkSummary = {
  totalMinutes: number;
  descriptions: string[];
  leaveReasons: string[];
  hasLeave: boolean;
};

type DaySummaryAccumulator = {
  hasWork: boolean;
  hasLeave: boolean;
  isWeekend: boolean;
  period1ClockIn: string;
  period1ClockOut: string;
  hasAnyClockInOut: boolean;
  totalMinutes: number;
  x1Minutes: number;
  x15Minutes: number;
  x3Minutes: number;
};

function pad2(value: number): string {
  return value.toString().padStart(2, '0');
}

function getCurrentMonthMeta(): { year: number; monthIndex: number; monthLabel: string } {
  const now = new Date();
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: THAI_TIMEZONE,
    year: 'numeric',
    month: '2-digit'
  }).formatToParts(now);

  const yearText = parts.find((part) => part.type === 'year')?.value;
  const monthText = parts.find((part) => part.type === 'month')?.value;

  if (!yearText || !monthText) {
    throw new Error('Failed to resolve current month in Thai timezone.');
  }

  const year = Number(yearText);
  const monthNumber = Number(monthText);

  return {
    year,
    monthIndex: monthNumber - 1,
    monthLabel: `${year}-${pad2(monthNumber)}`
  };
}

function getMonthMetaFromInput(targetMonth?: string): { year: number; monthIndex: number; monthLabel: string } {
  if (!targetMonth) {
    return getCurrentMonthMeta();
  }

  const [yearPart, monthPart] = targetMonth.split('-');
  const year = Number(yearPart);
  const monthNumber = Number(monthPart);

  if (!Number.isInteger(year) || !Number.isInteger(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    throw new Error('Invalid target month. Use YYYY-MM.');
  }

  return {
    year,
    monthIndex: monthNumber - 1,
    monthLabel: `${year}-${pad2(monthNumber)}`
  };
}

function applyBorder(cell: ExcelJS.Cell): void {
  cell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
}

function toExcelSerialDate(year: number, monthIndex: number, day: number): number {
  const excelEpochUtcMs = Date.UTC(1899, 11, 30);
  const targetDateUtcMs = Date.UTC(year, monthIndex, day);
  return (targetDateUtcMs - excelEpochUtcMs) / 86_400_000;
}

function getRowField(row: TimesheetInputRow, candidates: string[]): string {
  for (const candidate of candidates) {
    if (candidate in row) {
      return (row[candidate] ?? '').trim();
    }
  }

  return '';
}

function parseWorkDate(value: string): { year: number; month: number; day: number } | null {
  const normalized = value.trim();
  if (!normalized) {
    return null;
  }

  const ymdMatch = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (ymdMatch) {
    const year = Number(ymdMatch[1]);
    const month = Number(ymdMatch[2]);
    const day = Number(ymdMatch[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return { year, month, day };
    }
    return null;
  }

  const dmyMatch = normalized.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmyMatch) {
    const day = Number(dmyMatch[1]);
    const month = Number(dmyMatch[2]);
    const year = Number(dmyMatch[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return { year, month, day };
    }
    return null;
  }

  return null;
}

function parseWorkTimeToMinutes(value: string): number {
  const normalized = value.trim();
  if (!normalized) {
    return 0;
  }

  const match = normalized.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) {
    return 0;
  }

  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] ?? '0');
  if (!Number.isInteger(hours) || !Number.isInteger(minutes) || !Number.isInteger(seconds)) {
    return 0;
  }

  const total = (hours * 60) + minutes + (seconds >= 30 ? 1 : 0);
  return Math.max(total, 0);
}

function formatMinutesToTime(totalMinutes: number): string {
  const safeTotal = Math.max(0, totalMinutes);
  const hours = Math.floor(safeTotal / 60);
  const minutes = safeTotal % 60;
  return `${pad2(hours)}:${pad2(minutes)}`;
}

function formatSummaryHours(totalMinutes: number): string {
  const safeTotal = Math.max(0, totalMinutes);
  const hours = Math.floor(safeTotal / 60);
  const minutes = safeTotal % 60;
  return `${hours}:${pad2(minutes)}`;
}

function parseClockToMinutes(value: string): number | null {
  const normalized = value.trim();
  const match = /^(\d{1,2}):(\d{2})$/.exec(normalized);
  if (!match) {
    return null;
  }

  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (Number.isNaN(hours) || Number.isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
    return null;
  }

  return (hours * 60) + minutes;
}

function overlapMinutes(start: number, end: number, windowStart: number, windowEnd: number): number {
  const from = Math.max(start, windowStart);
  const to = Math.min(end, windowEnd);
  return Math.max(0, to - from);
}

function classifyMinutesByDayType(startMinute: number, endMinute: number, isWeekend: boolean): {
  x1Minutes: number;
  x15Minutes: number;
  x3Minutes: number;
} {
  let start = startMinute;
  let end = endMinute;

  if (end <= start) {
    end += 24 * 60;
  }

  let x1Minutes = 0;
  let x15Minutes = 0;
  let x3Minutes = 0;

  while (start < end) {
    const segmentEnd = Math.min(end, 24 * 60);

    if (isWeekend) {
      x3Minutes += overlapMinutes(start, segmentEnd, 0, 8 * 60 + 30);
      x1Minutes += overlapMinutes(start, segmentEnd, 8 * 60 + 30, 24 * 60);
    } else {
      x15Minutes += overlapMinutes(start, segmentEnd, 0, 8 * 60 + 30);
      x15Minutes += overlapMinutes(start, segmentEnd, 18 * 60 + 30, 24 * 60);
    }

    if (segmentEnd >= end) {
      break;
    }

    // Continue from midnight for an interval that crosses to the next day.
    end -= 24 * 60;
    start = 0;
  }

  return { x1Minutes, x15Minutes, x3Minutes };
}

function buildDailyWorkSummary(
  timesheetRows: TimesheetInputRow[],
  year: number,
  monthIndex: number
): Map<number, DailyWorkSummary> {
  const daily = new Map<number, DailyWorkSummary>();
  const targetMonth = monthIndex + 1;

  for (const row of timesheetRows) {
    const workDateRaw = getRowField(row, ['WorkDate', 'workDate']);
    const parsedDate = parseWorkDate(workDateRaw);
    if (!parsedDate) {
      continue;
    }

    if (parsedDate.year !== year || parsedDate.month !== targetMonth) {
      continue;
    }

    const workType = getRowField(row, ['WorkType', 'workType']).toUpperCase();
    const stageType = getRowField(row, ['StageType', 'stageType']);
    const workTimeRaw = getRowField(row, ['WorkTime', 'workTime']);
    const description = getRowField(row, ['Description', 'description']);

    const current = daily.get(parsedDate.day) ?? {
      totalMinutes: 0,
      descriptions: [],
      leaveReasons: [],
      hasLeave: false
    };

    if (workType === 'LEAVE') {
      current.hasLeave = true;
      if (stageType) {
        current.leaveReasons.push(stageType);
      }
      daily.set(parsedDate.day, current);
      continue;
    }

    current.totalMinutes += parseWorkTimeToMinutes(workTimeRaw);

    if (description) {
      current.descriptions.push(description);
    }

    daily.set(parsedDate.day, current);
  }

  return daily;
}

function buildDailyPdfSummary(pdfRows: PdfTableInputRow[]): Map<number, PdfTableInputRow> {
  const daily = new Map<number, PdfTableInputRow>();

  for (const row of pdfRows) {
    const dayNumber = Number(row.day);
    if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > 31) {
      continue;
    }

    if (!row.from && !row.to && !row.hours) {
      continue;
    }

    daily.set(dayNumber, {
      day: row.day,
      from: row.from?.trim() ?? '',
      to: row.to?.trim() ?? '',
      hours: row.hours?.trim() ?? ''
    });
  }

  return daily;
}

export async function generateOvertimeTemplateXlsx(
  targetMonth?: string,
  timesheetRows: TimesheetInputRow[] = [],
  pdfTableRows: PdfTableInputRow[] = []
): Promise<TemplateGenerationResult> {
  const { year, monthIndex, monthLabel } = getMonthMetaFromInput(targetMonth);
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const dailyWorkSummary = buildDailyWorkSummary(timesheetRows, year, monthIndex);
  const dailyPdfSummary = buildDailyPdfSummary(pdfTableRows);

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Overtime');

  worksheet.columns = [
    { width: 13 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 10 },
    { width: 14 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 },
    { width: 8 }
  ];

  worksheet.mergeCells('A1:D1');
  worksheet.getCell('A1').value = 'รายงานเวลาทำงานล่วงเวลาของพนักงาน';
  worksheet.getCell('A1').font = { name: FONT_NAME, size: 16, bold: true, underline: true };

  worksheet.getCell('A3').value = 'Username :';
  worksheet.mergeCells('B3:C3');
  worksheet.getCell('F3').value = 'Name-Surname :';
  worksheet.mergeCells('G3:I3');

  worksheet.getCell('A4').value = 'Company :';
  worksheet.mergeCells('B4:C4');
  worksheet.getCell('D4').value = 'Team :';
  worksheet.mergeCells('E4:F4');
  worksheet.getCell('G4').value = 'Month :';
  worksheet.mergeCells('H4:I4');
  worksheet.getCell('H4').value = monthLabel;

  for (const labelCell of ['A3', 'F3', 'A4', 'D4', 'G4']) {
    worksheet.getCell(labelCell).font = { name: FONT_NAME, size: 12, bold: true };
    worksheet.getCell(labelCell).alignment = { vertical: 'middle', horizontal: 'left' };
  }

  for (const valueCell of ['B3', 'G3', 'B4', 'E4', 'H4']) {
    worksheet.getCell(valueCell).font = { name: FONT_NAME, size: 12 };
    applyBorder(worksheet.getCell(valueCell));
  }

  worksheet.mergeCells('A6:A7');
  worksheet.getCell('A6').value = 'วันที่ทำงาน';

  worksheet.mergeCells('B6:C6');
  worksheet.getCell('B6').value = 'ช่วงที่ 1';
  worksheet.mergeCells('D6:E6');
  worksheet.getCell('D6').value = 'ช่วงที่ 2';
  worksheet.mergeCells('F6:G6');
  worksheet.getCell('F6').value = 'ช่วงที่ 3';
  worksheet.mergeCells('H6:J6');
  worksheet.getCell('H6').value = 'รวมเวลา';
  worksheet.mergeCells('K6:Q6');
  worksheet.getCell('K6').value = 'เหตุผลการทำงาน(โปรดระบุประเภทของงาน)';

  worksheet.getCell('B7').value = 'เข้า';
  worksheet.getCell('C7').value = 'ออก';
  worksheet.getCell('D7').value = 'เข้า';
  worksheet.getCell('E7').value = 'ออก';
  worksheet.getCell('F7').value = 'เข้า';
  worksheet.getCell('G7').value = 'ออก';
  worksheet.getCell('H7').value = 'รวม';
  worksheet.getCell('I7').value = 'OT';
  worksheet.getCell('J7').value = 'อนุมัติ OT';
  worksheet.mergeCells('K7:Q7');

  const headerCells = ['A6', 'B6', 'D6', 'F6', 'H6', 'K6'];
  for (const cellRef of headerCells) {
    const cell = worksheet.getCell(cellRef);
    cell.font = { name: FONT_NAME, size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: HEADER_BLUE }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  }

  for (const cellRef of ['B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7', 'I7', 'J7']) {
    const cell = worksheet.getCell(cellRef);
    cell.font = { name: FONT_NAME, size: 12, bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: SUBHEADER_ORANGE }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  worksheet.getRow(6).height = 24;
  worksheet.getRow(7).height = 22;

  const startRow = 8;
  const endRow = startRow + daysInMonth - 1;
  const daySummaries = new Map<number, DaySummaryAccumulator>();

  for (let day = 1; day <= daysInMonth; day += 1) {
    const rowIndex = startRow + day - 1;
    const dayOfWeek = new Date(Date.UTC(year, monthIndex, day)).getUTCDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const daySummary: DaySummaryAccumulator = {
      hasWork: false,
      hasLeave: false,
      isWeekend,
      period1ClockIn: '',
      period1ClockOut: '',
      hasAnyClockInOut: false,
      totalMinutes: 0,
      x1Minutes: 0,
      x15Minutes: 0,
      x3Minutes: 0
    };
    const dateCell = worksheet.getCell(`A${rowIndex}`);
    dateCell.value = toExcelSerialDate(year, monthIndex, day);
    dateCell.numFmt = 'dd/mm/yyyy';
    dateCell.alignment = { vertical: 'middle', horizontal: 'left' };
    dateCell.font = { name: FONT_NAME, size: 12 };

    worksheet.mergeCells(`K${rowIndex}:Q${rowIndex}`);

    for (let column = 1; column <= 17; column += 1) {
      const cell = worksheet.getRow(rowIndex).getCell(column);
      cell.font = { name: FONT_NAME, size: 12 };
      if (isWeekend) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: ROW_WEEKEND_GRAY }
        };
      }
      if (column !== 1) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    const workSummary = dailyWorkSummary.get(day);
    if (workSummary) {
      daySummary.hasLeave = workSummary.hasLeave;
      const uniqueDescriptions = Array.from(new Set(workSummary.descriptions));
      const uniqueLeaveReasons = Array.from(new Set(workSummary.leaveReasons));
      const descriptionText = workSummary.hasLeave
        ? uniqueLeaveReasons.join(', ')
        : uniqueDescriptions.join(', ');

      if (!workSummary.hasLeave) {
        const cappedMinutes = Math.min(workSummary.totalMinutes, 8 * 60);
        const clockIn = '08:30';
        const clockOut = formatMinutesToTime((8 * 60) + 30 + cappedMinutes + LUNCH_BREAK_MINUTES);
        worksheet.getCell(`B${rowIndex}`).value = clockIn;
        worksheet.getCell(`C${rowIndex}`).value = clockOut;
        daySummary.period1ClockIn = clockIn;
        daySummary.period1ClockOut = clockOut;
        daySummary.hasAnyClockInOut = true;

        const period1Start = parseClockToMinutes(clockIn);
        const period1End = parseClockToMinutes(clockOut);
        if (period1Start !== null && period1End !== null) {
          let workedMinutes = period1End - period1Start;
          if (workedMinutes <= 0) {
            workedMinutes += 24 * 60;
          }
          daySummary.totalMinutes += workedMinutes;
          daySummary.hasWork = workedMinutes > 0;

          const period1Classified = classifyMinutesByDayType(period1Start, period1End, isWeekend);
          daySummary.x1Minutes += period1Classified.x1Minutes;
          daySummary.x15Minutes += period1Classified.x15Minutes;
          daySummary.x3Minutes += period1Classified.x3Minutes;
        }
      }

      worksheet.getCell(`K${rowIndex}`).value = descriptionText;
      worksheet.getCell(`K${rowIndex}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    }

    const pdfSummary = dailyPdfSummary.get(day);
    if (pdfSummary) {
      worksheet.getCell(`D${rowIndex}`).value = pdfSummary.from;
      worksheet.getCell(`E${rowIndex}`).value = pdfSummary.to;
      worksheet.getCell(`I${rowIndex}`).value = pdfSummary.hours;
      daySummary.hasAnyClockInOut = Boolean(pdfSummary.from || pdfSummary.to) || daySummary.hasAnyClockInOut;

      const period2Start = parseClockToMinutes(pdfSummary.from);
      const period2End = parseClockToMinutes(pdfSummary.to);
      if (period2Start !== null && period2End !== null) {
        let workedMinutes = period2End - period2Start;
        if (workedMinutes <= 0) {
          workedMinutes += 24 * 60;
        }
        daySummary.totalMinutes += workedMinutes;
        daySummary.hasWork = workedMinutes > 0 || daySummary.hasWork;

        const period2Classified = classifyMinutesByDayType(period2Start, period2End, isWeekend);
        daySummary.x1Minutes += period2Classified.x1Minutes;
        daySummary.x15Minutes += period2Classified.x15Minutes;
        daySummary.x3Minutes += period2Classified.x3Minutes;
      }
    }

    daySummaries.set(day, daySummary);
  }

  for (let row = 6; row <= endRow; row += 1) {
    for (let column = 1; column <= 17; column += 1) {
      applyBorder(worksheet.getRow(row).getCell(column));
    }
  }

  const summaryStartRow = endRow + 3;
  const summaryItems = [
    'ทำงานทั้งหมด',
    'สาย + ออกก่อน',
    'ขาดงาน',
    'ลางาน',
    'ล่วงเวลาทั้งหมด',
    'ล่วงเวลาอนุมัติ',
    'จำนวนชั่วโมงการบริการส่วนเพิ่ม x 1',
    'จำนวนชั่วโมงการบริการส่วนเพิ่ม x 1.5',
    'จำนวนชั่วโมงการบริการส่วนเพิ่ม x 3'
  ];

  worksheet.mergeCells(`A${summaryStartRow}:I${summaryStartRow}`);
  worksheet.getCell(`A${summaryStartRow}`).value = 'สรุปเวลา :';
  worksheet.getCell(`J${summaryStartRow}`).value = 'วัน';
  worksheet.getCell(`K${summaryStartRow}`).value = 'ชั่วโมง';

  for (const cellRef of [`A${summaryStartRow}`, `J${summaryStartRow}`, `K${summaryStartRow}`]) {
    const cell = worksheet.getCell(cellRef);
    cell.font = { name: FONT_NAME, size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: HEADER_BLUE }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  }

  summaryItems.forEach((label, index) => {
    const row = summaryStartRow + index + 1;
    worksheet.mergeCells(`A${row}:I${row}`);
    worksheet.getCell(`A${row}`).value = label;
    worksheet.getCell(`A${row}`).font = { name: FONT_NAME, size: 12 };
    worksheet.getCell(`A${row}`).alignment = { vertical: 'middle', horizontal: 'left' };

    for (const cellRef of [`A${row}`, `J${row}`, `K${row}`]) {
      const cell = worksheet.getCell(cellRef);
      cell.font = { name: FONT_NAME, size: 12 };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: SUMMARY_GRAY }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
    }
  });

  const summaryEndRow = summaryStartRow + summaryItems.length;

  const totalWorkedDays = Array.from(daySummaries.values()).filter(
    (daySummary) => daySummary.period1ClockIn === '08:30' && daySummary.period1ClockOut === '17:30'
  ).length;
  const totalWorkedMinutes = Array.from(daySummaries.values()).reduce((sum, daySummary) => sum + daySummary.totalMinutes, 0);
  const leaveDays = Array.from(daySummaries.values()).filter((daySummary) => daySummary.hasLeave).length;
  const absentDays = Array.from(daySummaries.values()).filter(
    (daySummary) => !daySummary.isWeekend && !daySummary.hasLeave && !daySummary.hasAnyClockInOut
  ).length;
  const lateOrEarlyDays = Array.from(daySummaries.values()).filter(
    (daySummary) => !daySummary.isWeekend && !daySummary.hasLeave && daySummary.period1ClockOut !== '' && daySummary.period1ClockOut !== '17:30'
  ).length;

  const x1Minutes = Array.from(daySummaries.values()).reduce((sum, daySummary) => sum + daySummary.x1Minutes, 0);
  const x15Minutes = Array.from(daySummaries.values()).reduce((sum, daySummary) => sum + daySummary.x15Minutes, 0);
  const x3Minutes = Array.from(daySummaries.values()).reduce((sum, daySummary) => sum + daySummary.x3Minutes, 0);
  const totalOtMinutes = x1Minutes + x15Minutes + x3Minutes;
  const x1Days = Array.from(daySummaries.values()).filter((daySummary) => daySummary.x1Minutes > 0).length;
  const x15Days = Array.from(daySummaries.values()).filter((daySummary) => daySummary.x15Minutes > 0).length;
  const x3Days = Array.from(daySummaries.values()).filter((daySummary) => daySummary.x3Minutes > 0).length;
  const totalOtDays = Array.from(daySummaries.values()).filter(
    (daySummary) => daySummary.x1Minutes + daySummary.x15Minutes + daySummary.x3Minutes > 0
  ).length;

  worksheet.getCell(`J${summaryStartRow + 1}`).value = totalWorkedDays;
  worksheet.getCell(`K${summaryStartRow + 1}`).value = '';
  worksheet.getCell(`J${summaryStartRow + 2}`).value = lateOrEarlyDays;
  worksheet.getCell(`K${summaryStartRow + 2}`).value = '0:00';
  worksheet.getCell(`J${summaryStartRow + 3}`).value = absentDays;
  worksheet.getCell(`K${summaryStartRow + 3}`).value = '0:00';
  worksheet.getCell(`J${summaryStartRow + 4}`).value = leaveDays;
  worksheet.getCell(`K${summaryStartRow + 4}`).value = '0:00';
  worksheet.getCell(`J${summaryStartRow + 5}`).value = '';
  worksheet.getCell(`K${summaryStartRow + 5}`).value = formatSummaryHours(totalOtMinutes);
  worksheet.getCell(`J${summaryStartRow + 6}`).value = '';
  worksheet.getCell(`K${summaryStartRow + 6}`).value = formatSummaryHours(totalOtMinutes);
  worksheet.getCell(`J${summaryStartRow + 7}`).value = '';
  worksheet.getCell(`K${summaryStartRow + 7}`).value = formatSummaryHours(x1Minutes);
  worksheet.getCell(`J${summaryStartRow + 8}`).value = '';
  worksheet.getCell(`K${summaryStartRow + 8}`).value = formatSummaryHours(x15Minutes);
  worksheet.getCell(`J${summaryStartRow + 9}`).value = '';
  worksheet.getCell(`K${summaryStartRow + 9}`).value = formatSummaryHours(x3Minutes);

  for (let row = summaryStartRow; row <= summaryEndRow; row += 1) {
    for (let column = 1; column <= 11; column += 1) {
      applyBorder(worksheet.getRow(row).getCell(column));
    }
  }

  const notesStartRow = summaryStartRow;
  worksheet.mergeCells(`N${notesStartRow}:Q${notesStartRow}`);
  worksheet.getCell(`N${notesStartRow}`).value = 'หมายเหตุ';
  worksheet.getCell(`N${notesStartRow}`).font = { name: FONT_NAME, size: 16, bold: true };
  worksheet.getCell(`N${notesStartRow}`).alignment = { vertical: 'middle', horizontal: 'left' };

  const notes = [
    { symbol: '1', text: 'ทำงานล่วงเวลา' },
    { symbol: '2', text: 'มาสาย' },
    { symbol: '3', text: 'ออกก่อน' }
  ];

  notes.forEach((note, index) => {
    const row = notesStartRow + index + 1;
    worksheet.getCell(`M${row}`).value = note.symbol;
    worksheet.getCell(`M${row}`).font = { name: FONT_NAME, size: 12 };
    worksheet.getCell(`M${row}`).alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells(`N${row}:Q${row}`);
    worksheet.getCell(`N${row}`).value = note.text;
    worksheet.getCell(`N${row}`).font = { name: FONT_NAME, size: 12 };
    worksheet.getCell(`N${row}`).alignment = { vertical: 'middle', horizontal: 'left' };
  });

  const receiverRow = notesStartRow + 8;
  worksheet.mergeCells(`N${receiverRow}:Q${receiverRow}`);
  worksheet.getCell(`N${receiverRow}`).value = 'Reciever';
  worksheet.getCell(`N${receiverRow}`).font = { name: FONT_NAME, size: 16, bold: true };
  worksheet.getCell(`N${receiverRow}`).alignment = { vertical: 'middle', horizontal: 'left' };

  const dateRow = receiverRow + 3;
  worksheet.getCell(`N${dateRow}`).value = 'Date:';
  worksheet.getCell(`N${dateRow}`).font = { name: FONT_NAME, size: 16 };
  worksheet.getCell(`N${dateRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
  for (let column = 14; column <= 17; column += 1) {
    worksheet.getRow(dateRow).getCell(column).border = {
      bottom: { style: 'thin' }
    };
  }

  const printAreaEndRow = Math.max(summaryEndRow, dateRow);
  worksheet.pageSetup.printArea = `A1:Q${printAreaEndRow}`;
  worksheet.pageSetup.orientation = 'landscape';
  worksheet.pageSetup.fitToPage = true;
  worksheet.pageSetup.fitToWidth = 1;
  worksheet.pageSetup.fitToHeight = 1;
  worksheet.pageSetup.paperSize = 9;

  worksheet.getCell('K6').alignment = { vertical: 'middle', horizontal: 'left' };

  // const outputDirectory = path.join(process.cwd(), 'output');
  // await mkdir(outputDirectory, { recursive: true });
  const outputDirectory = path.join(os.tmpdir(), 'overtime_bot_output');
  await mkdir(outputDirectory, { recursive: true });

  const fileName = `overtime-template-${monthLabel}.xlsx`;
  const filePath = path.join(outputDirectory, fileName);
  await workbook.xlsx.writeFile(filePath);

  return {
    fileName,
    filePath,
    daysInMonth,
    month: monthLabel
  };
}