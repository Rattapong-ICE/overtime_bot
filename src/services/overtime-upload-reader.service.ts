import path from 'node:path';
import ExcelJS from 'exceljs';
import { PDFParse } from 'pdf-parse';
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
  const parser = new PDFParse({ data: new Uint8Array(buffer) });

  try {
    const textResult = await parser.getText();
    return textResult.text.trim();
  } finally {
    await parser.destroy();
  }
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
