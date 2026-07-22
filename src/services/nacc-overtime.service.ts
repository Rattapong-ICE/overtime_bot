import path from 'node:path';
import os from 'node:os';
import { mkdir } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { parse as parseCsv } from 'csv-parse/sync';
import { UserNaccModel } from '../models/user-nacc.model';

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
  otHours: number;
  amount: number;
};

const NACC_OT_START_HOUR = 16;
const NACC_OT_START_MINUTE = 30;
const NACC_OT_CAP_HOUR = 20;
const NACC_OT_CAP_MINUTE = 0;
const NACC_OT_RATE_PER_HOUR = 50;
const NACC_OT_CAP_AMOUNT = 200;

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

function calculateOtSummary(checkOutDateTime: string): { otHours: number; amount: number } {
  const parsed = parseDateTimeParts(checkOutDateTime);
  if (!parsed) {
    return { otHours: 0, amount: 0 };
  }

  const otStartTs = Date.UTC(
    parsed.year,
    parsed.month - 1,
    parsed.day,
    NACC_OT_START_HOUR,
    NACC_OT_START_MINUTE,
    0
  );

  const otCapTs = Date.UTC(
    parsed.year,
    parsed.month - 1,
    parsed.day,
    NACC_OT_CAP_HOUR,
    NACC_OT_CAP_MINUTE,
    0
  );

  const checkOutTs = toUtcTimestamp(checkOutDateTime);
  if (checkOutTs === null || checkOutTs <= otStartTs) {
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

  const employeeIds = Array.from(new Set(dailyRows.map((row) => row.employee_id)));
  const users = await UserNaccModel.find({
    employee_id: { $in: employeeIds }
  }).lean();

  const userMap = new Map(users.map((user) => [user.employee_id, user.name]));

  const reportRows = dailyRows
    .map((row) => {
      const otSummary = calculateOtSummary(row.checkOutDateTime);
      if (otSummary.amount <= 0) {
        return null;
      }
      const employeeName = userMap.get(row.employee_id) ?? '';

      return {
        employeeName,
        employeeId: row.employee_id,
        checkInDateTime: row.checkInDateTime,
        checkOutDateTime: row.checkOutDateTime,
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
    fgColor: { argb: 'FF1D4E89' }
  };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
  applyCellBorder(cell);
}

export async function generateNaccOvertimeExcel(month: string, rows: NaccParsedRow[]): Promise<NaccOvertimeExcelResult> {
  const reportRows = await buildOvertimeReportRows(month, rows);

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

    for (let columnIndex = 1; columnIndex <= headers.length; columnIndex += 1) {
      const cell = excelRow.getCell(columnIndex);
      applyCellBorder(cell);
      cell.alignment = {
        vertical: 'middle',
        horizontal: columnIndex === 1 ? 'left' : 'center'
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
