import path from 'node:path';
import ExcelJS from 'exceljs';
import { parse as parseCsv } from 'csv-parse/sync';

type TimesheetJsonRow = Record<string, string>;

type TextOtRow = {
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
  textOt: {
    fileName: string;
    text: string;
    tableRows: TextOtRow[];
  };
};

function normalizeClock(value: string): string {
  const match = /^(\d{1,2}):(\d{2})$/.exec(value.trim());
  if (!match) {
    return value.trim();
  }

  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (!Number.isInteger(hour) || !Number.isInteger(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    return value.trim();
  }

  return `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
}

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

function parseTextOtRows(textOt: string): TextOtRow[] {
  const lines = textOt
    .split(/\r?\n/)
    .map((line) => line.replace(/\r/g, '').trimEnd());

  const rows: TextOtRow[] = [];
  let currentRow: TextOtRow | null = null;

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      continue;
    }

    const tabColumns = rawLine.split('\t').map((column) => column.trim());
    const dayFromTab = Number(tabColumns[0] ?? '');
    if (Number.isInteger(dayFromTab) && dayFromTab >= 1 && dayFromTab <= 31 && tabColumns.length >= 4) {
      currentRow = {
        day: String(dayFromTab),
        from: normalizeClock(tabColumns[1] ?? ''),
        to: normalizeClock(tabColumns[2] ?? ''),
        hours: normalizeClock(tabColumns[3] ?? ''),
        description: (tabColumns.slice(4).join(' ') ?? '').trim(),
        remark: '',
        rawLine: line
      };
      rows.push(currentRow);
      continue;
    }

    const spacedMatch = /^(\d{1,2})\s+([0-2]?\d:\d{2})?\s+([0-2]?\d:\d{2})?\s+([0-2]?\d:\d{2})?\s*(.*)$/.exec(line);
    if (spacedMatch) {
      const dayNumber = Number(spacedMatch[1]);
      if (Number.isInteger(dayNumber) && dayNumber >= 1 && dayNumber <= 31) {
        currentRow = {
          day: String(dayNumber),
          from: normalizeClock(spacedMatch[2] ?? ''),
          to: normalizeClock(spacedMatch[3] ?? ''),
          hours: normalizeClock(spacedMatch[4] ?? ''),
          description: (spacedMatch[5] ?? '').trim(),
          remark: '',
          rawLine: line
        };
        rows.push(currentRow);
        continue;
      }
    }

    if (currentRow) {
      currentRow.description = currentRow.description
        ? `${currentRow.description} ${line}`
        : line;
    }
  }

  return rows;
}

export async function readUploadedOvertimeFiles(
  timesheetFile: Express.Multer.File,
  textOt: string
): Promise<UploadedData> {
  const timesheetType = getTimesheetFileType(timesheetFile.originalname);
  const timesheetRows = timesheetType === 'csv'
    ? parseCsvTimesheet(timesheetFile.buffer)
    : await parseXlsxTimesheet(timesheetFile.buffer);

  const normalizedTextOt = textOt.trim();
  const textOtRows = parseTextOtRows(normalizedTextOt);

  return {
    timesheet: {
      fileName: timesheetFile.originalname,
      fileType: timesheetType,
      rows: timesheetRows
    },
    textOt: {
      fileName: 'text_ot',
      text: normalizedTextOt,
      tableRows: textOtRows
    }
  };
}