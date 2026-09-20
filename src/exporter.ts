import * as XLSX from 'xlsx';
import type { IRecord } from '@lark-base-open/js-sdk';
import { cellToText } from './bitable-helper';
import type { ExportColumn } from './bitable-helper';

export type ExportFormat = 'xlsx' | 'csv' | 'json' | 'clipboard';

function buildMatrix(columns: ExportColumn[], rows: IRecord[]): string[][] {
  return rows.map((record) =>
    columns.map((column) => cellToText(record.fields[column.id] ?? null, column.type)),
  );
}

function csvEscape(value: string): string {
  if (/[",\n\r]/.test(value)) return `"${value.replace(/"/g, '""')}"`;
  return value;
}

function toCSV(columns: ExportColumn[], rows: IRecord[]): Blob {
  const lines = [
    columns.map((column) => csvEscape(column.name)).join(','),
    ...buildMatrix(columns, rows).map((line) => line.map(csvEscape).join(',')),
  ];
  // BOM 让 Excel 正确识别中文编码
  return new Blob(['\uFEFF', lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
}

function sanitizeSheetName(name: string): string {
  const cleaned = name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
  return cleaned || 'Sheet1';
}

function toXLSX(columns: ExportColumn[], rows: IRecord[], sheetName: string): Blob {
  const aoa: unknown[][] = [
    columns.map((column) => column.name),
    ...buildMatrix(columns, rows),
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(aoa);

  // 简单设置列宽，便于直接查看
  worksheet['!cols'] = columns.map(
    (column) => ({ wch: Math.min(40, Math.max(10, column.name.length * 2 + 2)) }),
  );

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeSheetName(sheetName));
  const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  return new Blob([output], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

function toJSON(columns: ExportColumn[], rows: IRecord[]): Blob {
  const data = rows.map((record) => {
    const obj: Record<string, string> = {};
    for (const column of columns) {
      obj[column.name] = cellToText(record.fields[column.id] ?? null, column.type);
    }
    return obj;
  });
  return new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8' });
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * 以制表符分隔复制（TSV），可直接粘贴进 Excel / 表格
 */
export async function copyToClipboard(columns: ExportColumn[], rows: IRecord[]): Promise<void> {
  const tsvEscape = (value: string) => value.replace(/\t/g, ' ').replace(/[\r\n]+/g, ' ');
  const lines = [
    columns.map((column) => tsvEscape(column.name)).join('\t'),
    ...buildMatrix(columns, rows).map((line) => line.map(tsvEscape).join('\t')),
  ];
  const text = lines.join('\r\n');

  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand('copy');
  textarea.remove();
}

export function buildFileBlob(
  format: ExportFormat,
  columns: ExportColumn[],
  rows: IRecord[],
  sheetName: string,
): Blob {
  switch (format) {
    case 'xlsx':
      return toXLSX(columns, rows, sheetName);
    case 'csv':
      return toCSV(columns, rows);
    case 'json':
      return toJSON(columns, rows);
    default:
      throw new Error(`unsupported format: ${format}`);
  }
}

export function buildFilename(base: string, ext: string): string {
  const now = new Date();
  const stamp = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(
    now.getHours(),
  )}${pad2(now.getMinutes())}`;
  return `${base.replace(/[\\/:*?"<>|]/g, '_')}_${stamp}.${ext}`;
}

function pad2(value: number): string {
  return String(value).padStart(2, '0');
}
