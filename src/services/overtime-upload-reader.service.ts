import path from 'node:path';
import ExcelJS from 'exceljs';
import pdfParse from 'pdf-parse';
import { parse as parseCsv } from 'csv-parse/sync';

type TimesheetJsonRow = Record<string, string>;

type PdfTableRow = {
  day: string;
  from: string;
  to: string;
  hours: string;
  description: string;
  remark: string;
  rawLine: string;
};

type UploadedData = {
  timesheet: {
    fileName: string;
    fileType: 'csv' | 'xlsx';
    rows: TimesheetJsonRow[];
  };
  sheetOtPdf: {
    fileName: string;
    text: string;
    tableRows: PdfTableRow[];
  };
};

function getTimesheetFileType(fileName: string): 'csv' | 'xlsx' {
  const extension = path.extname(fileName).toLowerCase();

  if (extension === '.csv') {
    return 'csv';
  }

  if (extension === '.xlsx') {
    return 'xlsx';
  }

  throw new Error('timesheet must be a .csv or .xlsx file');
}

function normalizeHeader(value: string, index: number): string {
  const sanitized = value.replace(/^\uFEFF/, '').trim();
  return sanitized || `column_${index + 1}`;
}

function mapRowsToJson(headers: string[], rows: string[][]): TimesheetJsonRow[] {
  return rows
    .map((row) => {
      const rowObject: TimesheetJsonRow = {};

      headers.forEach((header, index) => {
        rowObject[header] = (row[index] ?? '').trim();
      });

      return rowObject;
    })
    .filter((rowObject) => Object.values(rowObject).some((value) => value !== ''));
}

function parseCsvTimesheet(buffer: Buffer): TimesheetJsonRow[] {
  const content = buffer.toString('utf-8');
  const records = parseCsv(content, {
    columns: false,
    skip_empty_lines: true,
    trim: true
  }) as string[][];

  if (records.length === 0) {
    return [];
  }

  const [headerRow, ...dataRows] = records;
  const headers = (headerRow ?? []).map((value, index) => normalizeHeader(value, index));

  return mapRowsToJson(headers, dataRows);
}

async function parseXlsxTimesheet(buffer: Buffer): Promise<TimesheetJsonRow[]> {
  const workbook = new ExcelJS.Workbook();
  const normalizedBuffer = Buffer.from(buffer);
  await workbook.xlsx.load(normalizedBuffer as unknown as Parameters<typeof workbook.xlsx.load>[0]);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('timesheet xlsx has no worksheet');
  }

  const headerRow = worksheet.getRow(1);
  const maxColumns = headerRow.cellCount;
  const headers: string[] = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) {
      for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
        headers.push(normalizeHeader(row.getCell(columnIndex).text, columnIndex - 1));
      }
    }
  });

  const dataRows: string[][] = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    const values: string[] = [];
    for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
      values.push(row.getCell(columnIndex).text.trim());
    }

    dataRows.push(values);
  });

  return mapRowsToJson(headers, dataRows);
}

async function parsePdfText(buffer: Buffer): Promise<string> {
  const textResult = await pdfParse(buffer);
  return textResult.text.trim();
}

function normalizeTime(value: string): string {
  const normalized = value.replace('.', ':').trim();
  const match = /^(\d{1,2}):(\d{2})$/.exec(normalized);
  if (!match) {
    return normalized;
  }

  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (!Number.isInteger(hour) || !Number.isInteger(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    return normalized;
  }

  return `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
}

function extractDayNumber(value: string): number | null {
  const token = value.replace(/[.,;:()\[\]{}]/g, '').trim();

  const dayOnlyMatch = /^(\d{1,2})$/.exec(token);
  if (dayOnlyMatch) {
    const day = Number(dayOnlyMatch[1]);
    return day >= 1 && day <= 31 ? day : null;
  }

  const dateMatch = /^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/.exec(token);
  if (dateMatch) {
    const day = Number(dateMatch[1]);
    return day >= 1 && day <= 31 ? day : null;
  }

  return null;
}

function isTimeToken(value: string): boolean {
  return /^(\d{1,2})[:.](\d{2})$/.test(value.trim());
}

function extractCompactTripleTimes(value: string): { from: string; to: string; hours: string } | null {
  const token = value.trim();
  const match = /^(\d{1,2}[:.]\d{2})(\d{1,2}[:.]\d{2})(\d{1,2}[:.]\d{2})$/.exec(token);
  if (!match) {
    return null;
  }

  return {
    from: normalizeTime(match[1]),
    to: normalizeTime(match[2]),
    hours: normalizeTime(match[3])
  };
}

function parsePdfTableRowsFromWholeText(pdfText: string): PdfTableRow[] {
  const rows: PdfTableRow[] = [];
  const tokens = pdfText
    .replace(/\r?\n/g, ' ')
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token !== '');

  for (let index = 0; index < tokens.length; index += 1) {
    const dayNumber = extractDayNumber(tokens[index]);
    if (dayNumber === null) {
      continue;
    }

    const next1 = tokens[index + 1] ?? '';
    const next2 = tokens[index + 2] ?? '';
    const next3 = tokens[index + 3] ?? '';

    let from = '';
    let to = '';
    let hours = '';

    const compactTriple = extractCompactTripleTimes(next1);
    const timeRangeMatch = /^(\d{1,2}[:.]\d{2})-(\d{1,2}[:.]\d{2})$/.exec(next1);
    if (compactTriple) {
      from = compactTriple.from;
      to = compactTriple.to;
      hours = compactTriple.hours;
    } else if (timeRangeMatch && isTimeToken(next2)) {
      from = normalizeTime(timeRangeMatch[1]);
      to = normalizeTime(timeRangeMatch[2]);
      hours = normalizeTime(next2);
    } else if (isTimeToken(next1) && isTimeToken(next2) && isTimeToken(next3)) {
      from = normalizeTime(next1);
      to = normalizeTime(next2);
      hours = normalizeTime(next3);
    }

    if (!from || !to || !hours) {
      continue;
    }

    rows.push({
      day: String(dayNumber),
      from,
      to,
      hours,
      description: '',
      remark: '',
      rawLine: `${dayNumber} ${from} ${to} ${hours}`
    });
  }

  return rows;
}

function parsePdfTableRows(pdfText: string): PdfTableRow[] {
  const lines = pdfText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line !== '');

  const rows: PdfTableRow[] = [];
  let currentRow: PdfTableRow | null = null;
  const rowWithTimePattern = /^(?<day>\d{1,2})\s+(?<from>\d{1,2}:\d{2})\s+(?<to>\d{1,2}:\d{2})\s+(?<hours>\d{1,2}:\d{2})(?:\s+(?<description>.*))?$/;
  const dayOnlyPattern = /^(?<day>\d{1,2})$/;
  const stopLinePattern = /(ลงชื่อ|signature|พนักงาน|employee|หัวหน้างาน|supervisor|client authorized approval|classification|employee over time document|as of month|client project name|department|date\s*\/\s*time|description|remark|หมายเหตุ)/i;

  for (const line of lines) {
    const compactTriple = extractCompactTripleTimes(line);
    if (compactTriple && currentRow && !currentRow.from && !currentRow.to && !currentRow.hours) {
      currentRow.from = compactTriple.from;
      currentRow.to = compactTriple.to;
      currentRow.hours = compactTriple.hours;
      currentRow.rawLine = `${currentRow.day} ${compactTriple.from} ${compactTriple.to} ${compactTriple.hours}`;
      continue;
    }

    const rowWithTimeMatch = line.match(rowWithTimePattern);
    if (rowWithTimeMatch && rowWithTimeMatch.groups) {
      const dayNumber = Number(rowWithTimeMatch.groups.day);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > 31) {
        continue;
      }

      currentRow = {
        day: rowWithTimeMatch.groups.day,
        from: normalizeTime((rowWithTimeMatch.groups.from ?? '').trim()),
        to: normalizeTime((rowWithTimeMatch.groups.to ?? '').trim()),
        hours: normalizeTime((rowWithTimeMatch.groups.hours ?? '').trim()),
        description: (rowWithTimeMatch.groups.description ?? '').trim(),
        remark: '',
        rawLine: `${rowWithTimeMatch.groups.day} ${rowWithTimeMatch.groups.from} ${rowWithTimeMatch.groups.to} ${rowWithTimeMatch.groups.hours}`
      };

      rows.push(currentRow);
      continue;
    }

    const dayOnlyMatch = line.match(dayOnlyPattern);
    if (dayOnlyMatch && dayOnlyMatch.groups) {
      const dayNumber = Number(dayOnlyMatch.groups.day);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > 31) {
        continue;
      }

      currentRow = {
        day: dayOnlyMatch.groups.day,
        from: '',
        to: '',
        hours: '',
        description: '',
        remark: '',
        rawLine: dayOnlyMatch.groups.day
      };

      rows.push(currentRow);
      continue;
    }

    if (!currentRow) {
      continue;
    }

    if (stopLinePattern.test(line)) {
      currentRow = null;
      continue;
    }

    if (!currentRow.from || !currentRow.to || !currentRow.hours) {
      continue;
    }

    const continuation = line.trim();
    if (continuation === '') {
      continue;
    }

    currentRow.description = currentRow.description
      ? `${currentRow.description} ${continuation}`
      : continuation;
  }

  const hasTimedRows = rows.some((row) => row.from || row.to || row.hours);
  if (!hasTimedRows) {
    return parsePdfTableRowsFromWholeText(pdfText);
  }

  return rows;
}

export async function readUploadedOvertimeFiles(timesheetFile: Express.Multer.File, sheetOtPdfFile: Express.Multer.File): Promise<UploadedData> {
  const timesheetType = getTimesheetFileType(timesheetFile.originalname);
  const timesheetRows = timesheetType === 'csv'
    ? parseCsvTimesheet(timesheetFile.buffer)
    : await parseXlsxTimesheet(timesheetFile.buffer);

  const pdfText = await parsePdfText(sheetOtPdfFile.buffer);
  const pdfTableRows = parsePdfTableRows(pdfText);

  return {
    timesheet: {
      fileName: timesheetFile.originalname,
      fileType: timesheetType,
      rows: timesheetRows
    },
    sheetOtPdf: {
      fileName: sheetOtPdfFile.originalname,
      text: pdfText,
      tableRows: pdfTableRows
    }
  };
}
