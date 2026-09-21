import * as XLSX from 'xlsx';
import { FieldType } from '@lark-base-open/js-sdk';
import type { IRecord } from '@lark-base-open/js-sdk';
import { cellToText } from './bitable-helper';
import type { AttachmentInfo, AttachmentMap, ExportColumn } from './bitable-helper';
import type * as ExcelJSTypes from 'exceljs';

export type ExportFormat = 'xlsx' | 'csv' | 'json' | 'clipboard';

/** 图片在单元格中的显示尺寸（px），限制在单元格宽度内 */
const IMG_TARGET_WIDTH = 110;
const IMG_MAX_HEIGHT = 80;
/** 图片与单元格边缘、图片之间的间距（px） */
const IMG_GAP = 4;
/** 附件列列宽（Excel 字符单位，1 字符约 7px，16 字符约 117px > 图片宽 + 间距） */
const ATTACHMENT_COL_WIDTH = 16;
/** px → EMU（ExcelJS 锚点偏移单位，1px = 9525 EMU） */
const PX_TO_EMU = 9525;
/** 同时下载图片的并发数 */
const IMAGE_CONCURRENCY = 4;

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

type SupportedImageExt = 'jpeg' | 'png' | 'gif';

function resolveImageExtension(name: string, mime: string): SupportedImageExt | null {
  const lower = name.toLowerCase();
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg') || mime === 'image/jpeg') return 'jpeg';
  if (lower.endsWith('.png') || mime === 'image/png') return 'png';
  if (lower.endsWith('.gif') || mime === 'image/gif') return 'gif';
  return null;
}

interface DownloadedImage {
  buffer: ArrayBuffer;
  width: number;
  height: number;
  ext: SupportedImageExt;
}

/** 下载图片并读取原始尺寸；格式不支持或下载失败时返回 null */
async function downloadImage(info: AttachmentInfo): Promise<DownloadedImage | null> {
  const response = await fetch(info.url);
  if (!response.ok) return null;
  const blob = await response.blob();
  const ext = resolveImageExtension(info.name, blob.type);
  if (!ext) return null;

  const objectUrl = URL.createObjectURL(blob);
  try {
    const dimensions = await new Promise<{ width: number; height: number }>((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
      img.onerror = reject;
      img.src = objectUrl;
    });
    if (!dimensions.width || !dimensions.height) return null;
    const buffer = await blob.arrayBuffer();
    return { ...dimensions, buffer, ext };
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

interface CellImageJob {
  /** Excel 行号（1-based，含表头） */
  excelRow: number;
  /** 列索引（0-based） */
  colIndex: number;
  infos: AttachmentInfo[];
}

/** 已下载待放置的图片（统一排版，保证锚定在所属单元格内） */
interface PlacedImage {
  excelRow: number;
  colIndex: number;
  buffer: ArrayBuffer;
  ext: SupportedImageExt;
  width: number;
  height: number;
  /** 相对单元格顶部的像素偏移 */
  offsetY: number;
}

/**
 * 导出内嵌真实图片的 xlsx（附件列）。
 * 图片按单元格等比缩放、纵向排列并锁定在所属单元格内；
 * 附件取 URL 失败或格式不支持（如 webp）时，该格降级显示文件名。
 */
export async function exportXlsxWithImages(
  columns: ExportColumn[],
  rows: IRecord[],
  attachmentMap: AttachmentMap,
  sheetName: string,
  onProgress?: (done: number, total: number) => void,
): Promise<Blob> {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sanitizeSheetName(sheetName));

  const attachmentColIds = new Set(
    columns.filter((col) => col.type === FieldType.Attachment).map((col) => col.id),
  );

  // 表头
  worksheet.addRow(columns.map((col) => col.name));
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).alignment = { vertical: 'middle' };

  // 列宽：附件列固定宽度，确保图片完全落在单元格内
  worksheet.columns = columns.map((col) => ({
    width: attachmentColIds.has(col.id)
      ? ATTACHMENT_COL_WIDTH
      : Math.min(40, Math.max(10, col.name.length * 2 + 2)),
  }));

  // 数据行文本（附件列先留空）
  rows.forEach((record) => {
    const values = columns.map((col) =>
      col.type === FieldType.Attachment
        ? ''
        : cellToText(record.fields[col.id] ?? null, col.type),
    );
    const row = worksheet.addRow(values);
    row.alignment = { vertical: 'top', wrapText: true };
  });

  // 构建附件任务；无可用 URL 的单元格降级为文件名文本
  const fallbackTexts = new Map<string, string>();
  const jobs: CellImageJob[] = [];
  let totalImages = 0;

  columns.forEach((col, colIndex) => {
    if (col.type !== FieldType.Attachment) return;
    const perRecord = attachmentMap[col.id] ?? {};
    rows.forEach((record, rIdx) => {
      const excelRow = rIdx + 2;
      const infos = perRecord[record.recordId];
      if (infos?.length) {
        jobs.push({ excelRow, colIndex, infos });
        totalImages += infos.length;
      } else {
        const text = cellToText(record.fields[col.id] ?? null, col.type);
        if (text) fallbackTexts.set(`${excelRow}:${colIndex}`, text);
      }
    });
  });

  // 先并发下载收集全部图片，再统一排版，避免锚点依赖未定的行高
  const placed: PlacedImage[] = [];
  const rowPixelHeights = new Map<number, number>();
  let done = 0;

  const handleJob = async (job: CellImageJob): Promise<void> => {
    let y = IMG_GAP;
    const failedNames: string[] = [];

    for (const info of job.infos) {
      try {
        const image = await downloadImage(info);
        if (image) {
          const scale = Math.min(
            1,
            IMG_TARGET_WIDTH / image.width,
            IMG_MAX_HEIGHT / image.height,
          );
          const width = Math.round(image.width * scale);
          const height = Math.round(image.height * scale);
          placed.push({
            excelRow: job.excelRow,
            colIndex: job.colIndex,
            buffer: image.buffer,
            ext: image.ext,
            width,
            height,
            offsetY: y,
          });
          y += height + IMG_GAP;
        } else {
          failedNames.push(info.name);
        }
      } catch {
        failedNames.push(info.name);
      }
      done += 1;
      onProgress?.(done, totalImages);
    }

    if (failedNames.length) {
      fallbackTexts.set(`${job.excelRow}:${job.colIndex}`, failedNames.join(', '));
    }
    const prevHeight = rowPixelHeights.get(job.excelRow) ?? 0;
    if (y > prevHeight) rowPixelHeights.set(job.excelRow, y);
  };

  // 并发池
  let cursor = 0;
  const workers = Array.from({ length: Math.min(IMAGE_CONCURRENCY, jobs.length) }, async () => {
    while (cursor < jobs.length) {
      const job = jobs[cursor];
      cursor += 1;
      await handleJob(job);
    }
  });
  await Promise.all(workers);

  // 根据图片高度设置行高（px → point：×0.75）
  for (const [excelRow, pixels] of rowPixelHeights) {
    worksheet.getRow(excelRow).height = Math.max(20, pixels * 0.75 + 4);
  }

  // 图片锚定到所属单元格内部：单元格索引 + EMU 像素偏移（oneCellAnchor）
  for (const img of placed) {
    const imageId = workbook.addImage({
      buffer: img.buffer as never,
      extension: img.ext,
    });
    worksheet.addImage(imageId, {
      tl: {
        col: img.colIndex,
        row: img.excelRow - 1,
        nativeCol: IMG_GAP * PX_TO_EMU,
        nativeRow: img.offsetY * PX_TO_EMU,
      } as unknown as ExcelJSTypes.Anchor,
      ext: { width: img.width, height: img.height },
    });
  }

  // 降级文本写入
  for (const [key, text] of fallbackTexts) {
    const [excelRow, colIndex] = key.split(':').map(Number);
    worksheet.getRow(excelRow).getCell(colIndex + 1).value = text;
  }

  const output = (await workbook.xlsx.writeBuffer()) as BlobPart;
  return new Blob([output], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
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
