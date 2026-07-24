import path from 'node:path';
import os from 'node:os';
import { mkdir } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { parse as parseCsv } from 'csv-parse/sync';
import { UserNaccModel } from '../models/user-nacc.model';
import { HolidayNaccModel } from '../models/holiday-nacc.model';

export type NaccParsedRow = {
  employee_id: string;
  datetime: string;
};

export type NaccReadResult = {
  fileName: string;
  fileType: 'dat' | 'csv' | 'xlsx';
  headers: string[];
  rowCount: number;
  rows: NaccParsedRow[];
};

export type NaccOvertimeExcelResult = {
  fileName: string;
  filePath: string;
  rowCount: number;
};

type DailyMinMax = {
  employee_id: string;
  date: string;
  checkInDateTime: string;
  checkOutDateTime: string;
};

type NaccOvertimeReportRow = {
  employeeName: string;
  employeeId: string;
  checkInDateTime: string;
  checkOutDateTime: string;
  workDate: string;
  isOffDay: boolean;
  otHours: number;
  amount: number;
};

type NaccOvertimeMonthlySummaryRow = {
  month: string;
  employeeName: string;
  employeeId: string;
  totalOtHours: number;
  totalAmount: number;
  rate50Hours: number;
  rate50Amount: number;
  rate200Hours: number;
  rate200Amount: number;
  rate400Hours: number;
  rate400Amount: number;
};

const NACC_OT_START_HOUR = 16;
const NACC_OT_START_MINUTE = 30;
const NACC_OT_CAP_HOUR = 20;
const NACC_OT_CAP_MINUTE = 0;
const NACC_OT_RATE_PER_HOUR = 50;
const NACC_OT_CAP_AMOUNT = 200;
const OFF_DAY_ROW_FILL = 'FFC4E2';
const NACC_WEEKEND_OT_START_HOUR = 8;
const NACC_WEEKEND_OT_START_MINUTE = 30;
const NACC_WEEKEND_OT_END_HOUR = 16;
const NACC_WEEKEND_OT_END_MINUTE = 30;
const NACC_WEEKEND_LUNCH_BREAK_MINUTES = 60;
const NACC_LUNCH_WINDOW_START_MINUTES = 12 * 60;
const NACC_LUNCH_WINDOW_END_MINUTES = 13 * 60;
const NACC_WEEKEND_FULL_DAY_AMOUNT = 400;
const SUMMARY_RATE_50_FILL = 'FFEBCBFF';
const SUMMARY_RATE_200_FILL = 'FFFF5AA5';
const SUMMARY_RATE_400_FILL = 'FFF8C8C8';

function normalizeCell(value: string | undefined): string {
  return (value ?? '').replace(/^\uFEFF/, '').trim();
}

function isEmployeeId(value: string): boolean {
  return /^\d{3,}$/.test(value);
}

function isDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function isTime(value: string): boolean {
  return /^\d{2}:\d{2}:\d{2}$/.test(value);
}

function isDateTime(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}$/.test(value);
}

function findDateTimeAndIndex(normalized: string[]): { datetime: string; index: number } | null {
  for (let index = 0; index < normalized.length; index += 1) {
    if (isDateTime(normalized[index])) {
      return {
        datetime: normalized[index].replace('T', ' '),
        index
      };
    }
  }

  for (let index = 0; index < normalized.length - 1; index += 1) {
    if (isDate(normalized[index]) && isTime(normalized[index + 1])) {
      return {
        datetime: `${normalized[index]} ${normalized[index + 1]}`,
        index
      };
    }
  }

  return null;
}

function findNearestEmployeeId(normalized: string[], baseIndex: number): string {
  for (let left = baseIndex - 1; left >= 0; left -= 1) {
    if (isEmployeeId(normalized[left])) {
      return normalized[left];
    }
  }

  for (let right = baseIndex + 1; right < normalized.length; right += 1) {
    if (isEmployeeId(normalized[right])) {
      return normalized[right];
    }
  }

  return '';
}

function extractEmployeeDatetime(row: string[]): NaccParsedRow | null {
  const normalized = row.map((value) => normalizeCell(value));

  const dateTimeFound = findDateTimeAndIndex(normalized);
  if (!dateTimeFound) {
    return null;
  }

  const employeeId = findNearestEmployeeId(normalized, dateTimeFound.index);

  if (!employeeId) {
    return null;
  }

  return {
    employee_id: employeeId,
    datetime: dateTimeFound.datetime
  };
}

function mapRecordsToEmployeeDatetime(records: string[][]): NaccParsedRow[] {
  return records
    .map((row) => extractEmployeeDatetime(row))
    .filter((row): row is NaccParsedRow => row !== null);
}

function getFileType(fileName: string): 'dat' | 'csv' | 'xlsx' {
  const extension = path.extname(fileName).toLowerCase();

  if (extension === '.dat') {
    return 'dat';
  }

  if (extension === '.csv') {
    return 'csv';
  }

  if (extension === '.xlsx') {
    return 'xlsx';
  }

  throw new Error('File must be .dat, .csv, or .xlsx');
}

function parseDelimitedRows(content: string, delimiter: string | RegExp): string[][] {
  const lines = content
    .replaceAll('\r', '')
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line !== '');

  return lines.map((line) => line.split(delimiter).map((value) => value.trim()));
}

function parseDatBuffer(buffer: Buffer): NaccParsedRow[] {
  const content = buffer.toString('utf-8');

  if (!content.trim()) {
    return [];
  }

  const hasPipe = content.includes('|');
  const hasTab = content.includes('\t');
  const hasSemicolon = content.includes(';');
  const hasComma = content.includes(',');

  let records: string[][];
  if (hasPipe) {
    records = parseDelimitedRows(content, '|');
  } else if (hasTab) {
    records = parseDelimitedRows(content, '\t');
  } else if (hasSemicolon) {
    records = parseDelimitedRows(content, ';');
  } else if (hasComma) {
    records = parseDelimitedRows(content, ',');
  } else {
    records = parseDelimitedRows(content, /\s{2,}/);
  }

  if (records.length === 0) {
    return [];
  }

  return mapRecordsToEmployeeDatetime(records);
}

function parseCsvBuffer(buffer: Buffer): NaccParsedRow[] {
  const content = buffer.toString('utf-8');

  const records = parseCsv(content, {
    columns: false,
    skip_empty_lines: true,
    trim: true
  }) as string[][];

  if (records.length === 0) {
    return [];
  }

  return mapRecordsToEmployeeDatetime(records);
}

async function parseXlsxBuffer(buffer: Buffer): Promise<NaccParsedRow[]> {
  const workbook = new ExcelJS.Workbook();
  const normalizedBuffer = Buffer.from(buffer);
  await workbook.xlsx.load(normalizedBuffer as unknown as Parameters<typeof workbook.xlsx.load>[0]);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    return [];
  }

  const rows: string[][] = [];
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const maxColumns = Math.max(4, row.cellCount);
    const values: string[] = [];
    for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
      values.push(row.getCell(columnIndex).text.trim());
    }

    rows.push(values);
  });

  return mapRecordsToEmployeeDatetime(rows);
}

export async function readNaccOvertimeFile(file: Express.Multer.File): Promise<NaccReadResult> {
  const fileType = getFileType(file.originalname);
  let rows: NaccParsedRow[];

  if (fileType === 'dat') {
    rows = parseDatBuffer(file.buffer);
  } else if (fileType === 'csv') {
    rows = parseCsvBuffer(file.buffer);
  } else {
    rows = await parseXlsxBuffer(file.buffer);
  }

  const headers = ['employee_id', 'datetime'];

  return {
    fileName: file.originalname,
    fileType,
    headers,
    rowCount: rows.length,
    rows
  };
}

function parseDateTimeParts(value: string): {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
} | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})$/.exec(value.trim());
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);

  if (
    Number.isNaN(year)
    || Number.isNaN(month)
    || Number.isNaN(day)
    || Number.isNaN(hour)
    || Number.isNaN(minute)
    || Number.isNaN(second)
  ) {
    return null;
  }

  return {
    year,
    month,
    day,
    hour,
    minute,
    second
  };
}

function toUtcTimestamp(value: string): number | null {
  const parts = parseDateTimeParts(value);
  if (!parts) {
    return null;
  }

  return Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
}

function isWeekendDate(year: number, month: number, day: number): boolean {
  const dayOfWeek = new Date(Date.UTC(year, month - 1, day)).getUTCDay();
  return dayOfWeek === 0 || dayOfWeek === 6;
}

function isWeekendDateTime(value: string): boolean {
  const parts = parseDateTimeParts(value);
  if (!parts) {
    return false;
  }

  return isWeekendDate(parts.year, parts.month, parts.day);
}

function isWeekendDateText(dateText: string): boolean {
  const [yearText, monthText, dayText] = dateText.split('-');
  const year = Number(yearText);
  const month = Number(monthText);
  const day = Number(dayText);

  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return false;
  }

  return isWeekendDate(year, month, day);
}

function getUtcMinutesOfDay(timestamp: number): number {
  const date = new Date(timestamp);
  return (date.getUTCHours() * 60) + date.getUTCMinutes();
}

function overlapsLunchWindow(startTimestamp: number, endTimestamp: number): boolean {
  const startMinutesOfDay = getUtcMinutesOfDay(startTimestamp);
  const endMinutesOfDay = getUtcMinutesOfDay(endTimestamp);
  return startMinutesOfDay < NACC_LUNCH_WINDOW_END_MINUTES && endMinutesOfDay > NACC_LUNCH_WINDOW_START_MINUTES;
}

function buildDailyMinMaxRows(month: string, rows: NaccParsedRow[]): DailyMinMax[] {
  const monthPrefix = `${month}-`;
  const grouped = new Map<string, DailyMinMax>();

  for (const row of rows) {
    const dateTime = row.datetime.trim();
    if (!dateTime.startsWith(monthPrefix)) {
      continue;
    }

    const datePart = dateTime.slice(0, 10);
    const mapKey = `${row.employee_id}|${datePart}`;
    const current = grouped.get(mapKey);

    if (!current) {
      grouped.set(mapKey, {
        employee_id: row.employee_id,
        date: datePart,
        checkInDateTime: dateTime,
        checkOutDateTime: dateTime
      });
      continue;
    }

    if (dateTime < current.checkInDateTime) {
      current.checkInDateTime = dateTime;
    }

    if (dateTime > current.checkOutDateTime) {
      current.checkOutDateTime = dateTime;
    }
  }

  return Array.from(grouped.values());
}

function calculateOtSummary(
  checkInDateTime: string,
  checkOutDateTime: string,
  isOffDay: boolean
): { otHours: number; amount: number } {
  const checkInParts = parseDateTimeParts(checkInDateTime);
  const checkOutParts = parseDateTimeParts(checkOutDateTime);
  if (!checkInParts || !checkOutParts) {
    return { otHours: 0, amount: 0 };
  }

  const checkInTs = toUtcTimestamp(checkInDateTime);
  const checkOutTs = toUtcTimestamp(checkOutDateTime);
  if (checkInTs === null || checkOutTs === null || checkOutTs <= checkInTs) {
    return { otHours: 0, amount: 0 };
  }

  if (isOffDay) {
    const weekendWindowStartTs = Date.UTC(
      checkOutParts.year,
      checkOutParts.month - 1,
      checkOutParts.day,
      NACC_WEEKEND_OT_START_HOUR,
      NACC_WEEKEND_OT_START_MINUTE,
      0
    );
    const weekendWindowEndTs = Date.UTC(
      checkOutParts.year,
      checkOutParts.month - 1,
      checkOutParts.day,
      NACC_WEEKEND_OT_END_HOUR,
      NACC_WEEKEND_OT_END_MINUTE,
      0
    );

    const effectiveStartTs = Math.max(checkInTs, weekendWindowStartTs);
    const effectiveEndTs = Math.min(checkOutTs, weekendWindowEndTs);

    if (effectiveEndTs <= effectiveStartTs) {
      return { otHours: 0, amount: 0 };
    }

    // Business rule: full weekend shift in allowed window pays fixed 400 THB.
    if (checkInTs <= weekendWindowStartTs && checkOutTs >= weekendWindowEndTs) {
      return {
        otHours: NACC_WEEKEND_FULL_DAY_AMOUNT / NACC_OT_RATE_PER_HOUR,
        amount: NACC_WEEKEND_FULL_DAY_AMOUNT
      };
    }

    const workedMinutesInWindow = Math.floor((effectiveEndTs - effectiveStartTs) / 60_000);
    const shouldDeductLunchBreak = overlapsLunchWindow(effectiveStartTs, effectiveEndTs);
    const payableMinutes = Math.max(
      0,
      workedMinutesInWindow - (shouldDeductLunchBreak ? NACC_WEEKEND_LUNCH_BREAK_MINUTES : 0)
    );

    if (payableMinutes < 60) {
      return { otHours: 0, amount: 0 };
    }

    const fullHours = Math.floor(payableMinutes / 60);

    return {
      otHours: fullHours,
      amount: fullHours * NACC_OT_RATE_PER_HOUR
    };
  }

  const otStartTs = Date.UTC(
    checkOutParts.year,
    checkOutParts.month - 1,
    checkOutParts.day,
    NACC_OT_START_HOUR,
    NACC_OT_START_MINUTE,
    0
  );

  const otCapTs = Date.UTC(
    checkOutParts.year,
    checkOutParts.month - 1,
    checkOutParts.day,
    NACC_OT_CAP_HOUR,
    NACC_OT_CAP_MINUTE,
    0
  );

  if (checkOutTs <= otStartTs) {
    return { otHours: 0, amount: 0 };
  }

  if (checkOutTs >= otCapTs) {
    return {
      otHours: NACC_OT_CAP_AMOUNT / NACC_OT_RATE_PER_HOUR,
      amount: NACC_OT_CAP_AMOUNT
    };
  }

  const rawMinutes = Math.floor((checkOutTs - otStartTs) / 60_000);
  if (rawMinutes < 60) {
    return { otHours: 0, amount: 0 };
  }

  const fullHours = Math.floor(rawMinutes / 60);

  return {
    otHours: fullHours,
    amount: fullHours * NACC_OT_RATE_PER_HOUR
  };
}

async function buildOvertimeReportRows(month: string, rows: NaccParsedRow[]): Promise<NaccOvertimeReportRow[]> {
  const dailyRows = buildDailyMinMaxRows(month, rows);
  if (dailyRows.length === 0) {
    return [];
  }

  const enabledHolidayRows = await HolidayNaccModel.find({
    enabled: true,
    date: { $regex: `^${month}-` }
  }).lean();
  const holidayDateSet = new Set(enabledHolidayRows.map((holiday) => holiday.date));

  const employeeIds = Array.from(new Set(dailyRows.map((row) => row.employee_id)));
  const users = await UserNaccModel.find({
    employee_id: { $in: employeeIds }
  }).lean();

  const userMap = new Map(users.map((user) => [user.employee_id, user.name]));

  const reportRows = dailyRows
    .map((row) => {
      const isOffDay = isWeekendDateText(row.date) || holidayDateSet.has(row.date);
      const otSummary = calculateOtSummary(row.checkInDateTime, row.checkOutDateTime, isOffDay);
      if (otSummary.amount <= 0) {
        return null;
      }
      const employeeName = userMap.get(row.employee_id) ?? '';

      return {
        employeeName,
        employeeId: row.employee_id,
        checkInDateTime: row.checkInDateTime,
        checkOutDateTime: row.checkOutDateTime,
        workDate: row.date,
        isOffDay,
        otHours: otSummary.otHours,
        amount: otSummary.amount
      } as NaccOvertimeReportRow;
    })
    .filter((row): row is NaccOvertimeReportRow => row !== null)
    .sort((left, right) => {
      if (left.employeeId !== right.employeeId) {
        return left.employeeId.localeCompare(right.employeeId);
      }

      return left.checkInDateTime.localeCompare(right.checkInDateTime);
    });

  return reportRows;
}

function applyCellBorder(cell: ExcelJS.Cell): void {
  cell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    right: { style: 'thin' },
    bottom: { style: 'thin' }
  };
}

function applyHeaderStyle(cell: ExcelJS.Cell): void {
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF69B4' }
  };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
  applyCellBorder(cell);
}

function getSummaryBucketFill(columnIndex: number): string | null {
  if (columnIndex === 5 || columnIndex === 6) {
    return SUMMARY_RATE_50_FILL;
  }

  if (columnIndex === 7 || columnIndex === 8) {
    return SUMMARY_RATE_200_FILL;
  }

  if (columnIndex === 9 || columnIndex === 10) {
    return SUMMARY_RATE_400_FILL;
  }

  return null;
}

function buildMonthlySummaryRows(month: string, reportRows: NaccOvertimeReportRow[]): NaccOvertimeMonthlySummaryRow[] {
  const summaryMap = new Map<string, NaccOvertimeMonthlySummaryRow>();

  for (const row of reportRows) {
    const key = row.employeeId;
    const current = summaryMap.get(key);

    if (!current) {
      summaryMap.set(key, {
        month,
        employeeName: row.employeeName,
        employeeId: row.employeeId,
        totalOtHours: row.otHours,
        totalAmount: row.amount,
        rate50Hours: 0,
        rate50Amount: 0,
        rate200Hours: 0,
        rate200Amount: 0,
        rate400Hours: 0,
        rate400Amount: 0
      });
    } else {
      current.totalOtHours += row.otHours;
      current.totalAmount += row.amount;
    }

    const summaryRow = summaryMap.get(key);
    if (!summaryRow) {
      continue;
    }

    const isWeekdayWorkDate = !isWeekendDateText(row.workDate);

    if (row.isOffDay && row.amount === NACC_WEEKEND_FULL_DAY_AMOUNT) {
      summaryRow.rate400Hours += 1;
      summaryRow.rate400Amount += row.amount;
      continue;
    }

    if (isWeekdayWorkDate && !row.isOffDay && row.amount === NACC_OT_CAP_AMOUNT) {
      summaryRow.rate200Hours += 1;
      summaryRow.rate200Amount += row.amount;
      continue;
    }

    summaryRow.rate50Hours += row.otHours;
    summaryRow.rate50Amount += row.amount;
  }

  return Array.from(summaryMap.values()).sort((left, right) => left.employeeId.localeCompare(right.employeeId));
}

export async function generateNaccOvertimeExcel(month: string, rows: NaccParsedRow[]): Promise<NaccOvertimeExcelResult> {
  const reportRows = await buildOvertimeReportRows(month, rows);
  const summaryRows = buildMonthlySummaryRows(month, reportRows);

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('NACC Overtime');

  worksheet.columns = [
    { key: 'employeeName', width: 30 },
    { key: 'employeeId', width: 12 },
    { key: 'checkInDateTime', width: 22 },
    { key: 'checkOutDateTime', width: 22 },
    { key: 'otHours', width: 10 },
    { key: 'amount', width: 12 }
  ];

  const headers = ['ชื่อ พนักงาน', 'รหัส', 'เวลาเข้า', 'เวลาออก', 'ชม ot', 'จำนวนเงิน'];
  worksheet.addRow(headers);

  const headerRow = worksheet.getRow(1);
  headerRow.height = 24;
  for (let columnIndex = 1; columnIndex <= headers.length; columnIndex += 1) {
    applyHeaderStyle(headerRow.getCell(columnIndex));
  }

  for (const row of reportRows) {
    const excelRow = worksheet.addRow([
      row.employeeName,
      row.employeeId,
      row.checkInDateTime,
      row.checkOutDateTime,
      row.otHours,
      row.amount
    ]);
    const isOffDayRow = row.isOffDay || isWeekendDateTime(row.checkInDateTime);

    for (let columnIndex = 1; columnIndex <= headers.length; columnIndex += 1) {
      const cell = excelRow.getCell(columnIndex);
      applyCellBorder(cell);
      cell.alignment = {
        vertical: 'middle',
        horizontal: columnIndex === 1 ? 'left' : 'center'
      };

      if (isOffDayRow) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: OFF_DAY_ROW_FILL }
        };
      }
    }
  }

  const summaryWorksheet = workbook.addWorksheet('NACC OT Summary');
  summaryWorksheet.columns = [
    { key: 'sequence', width: 10 },
    { key: 'month', width: 12 },
    { key: 'employeeName', width: 30 },
    { key: 'employeeId', width: 12 },
    { key: 'rate50Hours', width: 12 },
    { key: 'rate50Amount', width: 12 },
    { key: 'rate200Hours', width: 12 },
    { key: 'rate200Amount', width: 12 },
    { key: 'rate400Hours', width: 12 },
    { key: 'rate400Amount', width: 12 },
    { key: 'totalOtHours', width: 14 },
    { key: 'totalAmount', width: 14 }
  ];

  const summaryHeaders = [
    'ลำดับ',
    'เดือน',
    'ชื่อ พนักงาน',
    'รหัส',
    '50/ชม',
    '50/บาท',
    '200/ชม',
    '200/บาท',
    '400/ชม',
    '400/บาท',
    'รวมชม ot',
    'รวมจำนวนเงิน'
  ];
  summaryWorksheet.addRow(summaryHeaders);

  const summaryHeaderRow = summaryWorksheet.getRow(1);
  summaryHeaderRow.height = 24;
  for (let columnIndex = 1; columnIndex <= summaryHeaders.length; columnIndex += 1) {
    const headerCell = summaryHeaderRow.getCell(columnIndex);
    applyHeaderStyle(headerCell);

    const bucketFill = getSummaryBucketFill(columnIndex);
    if (bucketFill) {
      headerCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: bucketFill }
      };
      headerCell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    }
  }

  for (const [index, row] of summaryRows.entries()) {
    const summaryExcelRow = summaryWorksheet.addRow([
      index + 1,
      row.month,
      row.employeeName,
      row.employeeId,
      row.rate50Hours,
      row.rate50Amount,
      row.rate200Hours,
      row.rate200Amount,
      row.rate400Hours,
      row.rate400Amount,
      row.totalOtHours,
      row.totalAmount
    ]);

    for (let columnIndex = 1; columnIndex <= summaryHeaders.length; columnIndex += 1) {
      const cell = summaryExcelRow.getCell(columnIndex);
      applyCellBorder(cell);

      const bucketFill = getSummaryBucketFill(columnIndex);
      if (bucketFill) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bucketFill }
        };
      }

      cell.alignment = {
        vertical: 'middle',
        horizontal: columnIndex === 3 ? 'left' : 'center'
      };
    }
  }

  const outputDir = path.join(os.tmpdir(), 'nacc-overtime');
  await mkdir(outputDir, { recursive: true });

  const fileName = `nacc-overtime-${month}.xlsx`;
  const filePath = path.join(outputDir, fileName);

  await workbook.xlsx.writeFile(filePath);

  return {
    fileName,
    filePath,
    rowCount: reportRows.length
  };
}
