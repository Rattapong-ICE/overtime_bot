import path from 'node:path';
import ExcelJS from 'exceljs';
import PDFParse from 'pdf-parse';
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

async function parsePdfText(buffer: Buffer): Promise<string> {
  const pdfParseAny = PDFParse as unknown as {
    (data: Buffer | Uint8Array): Promise<{ text: string }>;
    prototype?: {
      getText?: unknown;
      destroy?: unknown;
    };
    new (options: { data: Uint8Array }): {
      getText: () => Promise<{ text: string }>;
      destroy: () => Promise<void>;
    };
  };

  // Keep old behavior for versions that expose a class API.
  if (typeof pdfParseAny === 'function' && typeof pdfParseAny.prototype?.getText === 'function') {
    const parser = new pdfParseAny({ data: new Uint8Array(buffer) });

    try {
      const textResult = await parser.getText();
      return textResult.text.trim();
    } finally {
      await parser.destroy();
    }
  }

  // Fallback for versions that export a plain function.
  if (typeof pdfParseAny === 'function') {
    const textResult = await pdfParseAny(buffer);
    return textResult.text.trim();
  }

  throw new Error('Unsupported pdf-parse module format.');
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
    const rowWithTimeMatch = line.match(rowWithTimePattern);
    if (rowWithTimeMatch && rowWithTimeMatch.groups) {
      const dayNumber = Number(rowWithTimeMatch.groups.day);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > 31) {
        continue;
      }

      currentRow = {
        day: rowWithTimeMatch.groups.day,
        from: (rowWithTimeMatch.groups.from ?? '').trim(),
        to: (rowWithTimeMatch.groups.to ?? '').trim(),
        hours: (rowWithTimeMatch.groups.hours ?? '').trim(),
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

  return rows;
}

function parseTextOtRows(textOt: string): PdfTableRow[] {
  const lines = textOt
    .split(/\r?\n/)
    .map((line) => line.replace(/\r/g, '').trimEnd());

  const rows: PdfTableRow[] = [];
  let currentRow: PdfTableRow | null = null;

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
  sheetOtPdfFile?: Express.Multer.File,
  textOt?: string
): Promise<UploadedData> {
  const timesheetType = getTimesheetFileType(timesheetFile.originalname);
  const timesheetRows = timesheetType === 'csv'
    ? parseCsvTimesheet(timesheetFile.buffer)
    : await parseXlsxTimesheet(timesheetFile.buffer);

  const normalizedTextOt = textOt?.trim() ?? '';
  let sourceText = '';
  let sourceFileName = sheetOtPdfFile?.originalname ?? 'text_ot';
  let pdfTableRows: PdfTableRow[] = [];

  if (normalizedTextOt) {
    sourceText = normalizedTextOt;
    pdfTableRows = parseTextOtRows(normalizedTextOt);
  }

  if (pdfTableRows.length === 0 && sheetOtPdfFile) {
    sourceText = await parsePdfText(sheetOtPdfFile.buffer);
    sourceFileName = sheetOtPdfFile.originalname;
    pdfTableRows = parsePdfTableRows(sourceText);
  }

  return {
    timesheet: {
      fileName: timesheetFile.originalname,
      fileType: timesheetType,
      rows: timesheetRows
    },
    sheetOtPdf: {
      fileName: sourceFileName,
      text: sourceText,
      tableRows: pdfTableRows
    }
  };
}