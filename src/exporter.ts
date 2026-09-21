import { FieldType } from '@lark-base-open/js-sdk';
import type { IRecord } from '@lark-base-open/js-sdk';
import { cellToText } from './bitable-helper';
import type { AttachmentInfo, AttachmentMap, ExportColumn } from './bitable-helper';
import type * as ExcelJSTypes from 'exceljs';

export type ExportFormat = 'xlsx' | 'csv';

/** 图片在单元格中的显示尺寸（px），限制在单元格宽度内 */
const IMG_TARGET_WIDTH = 64;
const IMG_MAX_HEIGHT = 48;
/** 图片与单元格边缘、图片之间的间距（px） */
const IMG_GAP = 4;
/** 附件列列宽（Excel 字符单位，1 字符约 7px，10 字符约 75px > 图片宽 + 间距） */
const ATTACHMENT_COL_WIDTH = 10;
/** px → EMU（ExcelJS 锚点偏移单位，1px = 9525 EMU） */
const PX_TO_EMU = 9525;
/** 同时下载图片的并发数 */
const IMAGE_CONCURRENCY = 4;

/* ===== 美化样式（参考 DMS_DA「导出美化Excel」） ===== */
const FONT_NAME = '微软雅黑';
/** 一级文字深灰蓝 */
const TEXT_COLOR = 'FF2C3E50';
/** 表头填充与边框 */
const HEADER_FILL = 'FFDCE3E8';
const HEADER_BORDER = 'FFB0BEC5';
/** 数据行边框与斑马纹 */
const DATA_BORDER = 'FFE0E0E0';
const ZEBRA_FILL = 'FFF8F9FA';
const WHITE_FILL = 'FFFFFFFF';
/** 表头/数据行高（pt） */
const HEADER_ROW_HEIGHT = 30;
const DATA_ROW_HEIGHT = 22;

/** 双行合并组表头（与 DMS_DA 美化台账一致）：组名列横跨子列，子列名在第二行 */
const MERGE_GROUPS = [
  { groupName: '制约因素', fields: ['方案', '备件', '条件', '人员', '窗口'] },
  { groupName: '消项确认记录', fields: ['消缺人', '班长', '消项时间'] },
];
/** 合并组子列统一列宽 */
const GROUP_COL_WIDTH = 12;

interface GroupInfo {
  groupName: string;
  /** 起止列索引（0-based，含） */
  start: number;
  end: number;
}

/** 在导出列中匹配合并组：按字段名定位，仅取连续列段 */
function matchMergeGroups(columns: ExportColumn[]): GroupInfo[] {
  const infos: GroupInfo[] = [];
  for (const group of MERGE_GROUPS) {
    const indices = columns
      .map((col, idx) => (group.fields.includes(col.name) ? idx : -1))
      .filter((idx) => idx >= 0);
    if (!indices.length) continue;
    let runStart = indices[0];
    let prev = indices[0];
    for (let k = 1; k <= indices.length; k += 1) {
      const cur = indices[k];
      if (cur !== prev + 1) {
        infos.push({ groupName: group.groupName, start: runStart, end: prev });
        runStart = cur;
      }
      prev = cur;
    }
  }
  return infos;
}

/**
 * 单元格文本统一出口：「计划期限」列只保留日期（到日），不显示时分。
 */
function formatCellText(record: IRecord, col: ExportColumn): string {
  const text = cellToText(record.fields[col.id] ?? null, col.type);
  if (col.name.includes('计划期限')) {
    return text.replace(/[ T]\d{1,2}:\d{2}(:\d{2})?$/, '');
  }
  return text;
}

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

/** 展示宽度：中文等全角字符按 2 个字符计 */
function displayWidth(text: string): number {
  return text.replace(/[^\x00-\xff]/g, 'XX').length;
}

/** 统一设置行边框 */
function setRowBorder(row: ExcelJSTypes.Row, color: string): void {
  const side = { style: 'thin' as const, color: { argb: color } };
  row.border = { top: side, bottom: side, left: side, right: side };
}

/**
 * 构建带美化样式的工作簿（参考 DMS_DA「导出美化Excel」）：
 * 表头加粗浅灰蓝底（含「制约因素/消项确认记录」双行合并组表头）、数据行斑马纹居中、
 * 细边框、列宽按内容自适应、冻结表头、自动筛选。
 * getText 决定每个单元格文本；attachmentColWidth 传入时附件列使用固定宽度（用于图片排版）。
 */
async function buildStyledWorkbook(
  columns: ExportColumn[],
  rows: IRecord[],
  sheetName: string,
  getText: (record: IRecord, col: ExportColumn) => string,
  attachmentColWidth?: number,
): Promise<{ workbook: ExcelJSTypes.Workbook; headerRows: number }> {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sanitizeSheetName(sheetName));

  // 匹配合并组（如「制约因素」「消项确认记录」），存在时使用双行表头
  const groupInfos = matchMergeGroups(columns);
  const hasGroups = groupInfos.length > 0;
  const headerRows = hasGroups ? 2 : 1;
  const groupColMap = new Map<number, GroupInfo>();
  for (const group of groupInfos) {
    for (let i = group.start; i <= group.end; i += 1) groupColMap.set(i, group);
  }

  const headerStyle = {
    font: { name: FONT_NAME, size: 12, bold: true, color: { argb: TEXT_COLOR } },
    fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: HEADER_FILL } },
    alignment: { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true },
  };

  // 第一行表头：组内首列显示组名（其余留空），非组列显示字段名
  const header1 = worksheet.addRow(
    columns.map((col, idx) => {
      const group = groupColMap.get(idx);
      if (group) return idx === group.start ? group.groupName : '';
      return col.name;
    }),
  );
  header1.height = HEADER_ROW_HEIGHT;
  header1.font = headerStyle.font;
  header1.fill = headerStyle.fill;
  header1.alignment = headerStyle.alignment;
  setRowBorder(header1, HEADER_BORDER);

  // 第二行表头：组内列显示子字段名，其余留空（稍后与第一行纵向合并）
  if (hasGroups) {
    const header2 = worksheet.addRow(columns.map((col, idx) => (groupColMap.has(idx) ? col.name : '')));
    header2.height = HEADER_ROW_HEIGHT;
    header2.font = headerStyle.font;
    header2.fill = headerStyle.fill;
    header2.alignment = headerStyle.alignment;
    setRowBorder(header2, HEADER_BORDER);
  }

  // 列宽：按内容自适应（中文按 2 字符），上限 40；附件列/合并组子列可固定宽度
  worksheet.columns = columns.map((col, idx) => {
    if (col.type === FieldType.Attachment && attachmentColWidth) {
      return { width: attachmentColWidth };
    }
    if (groupColMap.has(idx)) {
      return { width: GROUP_COL_WIDTH };
    }
    let maxLen = displayWidth(col.name);
    for (const record of rows) {
      const len = displayWidth(getText(record, col));
      if (len > maxLen) maxLen = len;
    }
    return { width: Math.min(40, Math.max(10, maxLen + 4)) };
  });

  // 合并单元格：组名横向合并（第一行），非组列纵向合并（两行表头）
  if (hasGroups) {
    for (const group of groupInfos) {
      if (group.end > group.start) {
        worksheet.mergeCells(1, group.start + 1, 1, group.end + 1);
      }
    }
    for (let idx = 0; idx < columns.length; idx += 1) {
      if (!groupColMap.has(idx)) {
        worksheet.mergeCells(1, idx + 1, 2, idx + 1);
      }
    }
  }

  // 数据行：斑马纹 + 居中 + 细边框
  rows.forEach((record, idx) => {
    const row = worksheet.addRow(columns.map((col) => getText(record, col)));
    row.height = DATA_ROW_HEIGHT;
    row.font = { name: FONT_NAME, size: 11, color: { argb: TEXT_COLOR } };
    row.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: idx % 2 === 0 ? ZEBRA_FILL : WHITE_FILL },
    };
    row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    setRowBorder(row, DATA_BORDER);
  });

  // 冻结表头 + 自动筛选
  worksheet.views = [{ state: 'frozen', ySplit: headerRows }];
  worksheet.autoFilter = {
    from: { row: headerRows, column: 1 },
    to: { row: rows.length + headerRows, column: columns.length },
  };

  return { workbook, headerRows };
}

async function toXLSX(columns: ExportColumn[], rows: IRecord[], sheetName: string): Promise<Blob> {
  const { workbook } = await buildStyledWorkbook(columns, rows, sheetName, formatCellText);
  const output = (await workbook.xlsx.writeBuffer()) as ArrayBuffer;
  return new Blob([output], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
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

export async function buildFileBlob(
  format: ExportFormat,
  columns: ExportColumn[],
  rows: IRecord[],
  sheetName: string,
): Promise<Blob> {
  if (format === 'csv') return toCSV(columns, rows);
  return toXLSX(columns, rows, sheetName);
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
  const { workbook, headerRows } = await buildStyledWorkbook(
    columns,
    rows,
    sheetName,
    (record, col) =>
      col.type === FieldType.Attachment ? '' : formatCellText(record, col),
    ATTACHMENT_COL_WIDTH,
  );
  const worksheet = workbook.worksheets[0];

  // 构建附件任务；无可用 URL 的单元格降级为文件名文本
  const fallbackTexts = new Map<string, string>();
  const jobs: CellImageJob[] = [];
  let totalImages = 0;

  columns.forEach((col, colIndex) => {
    if (col.type !== FieldType.Attachment) return;
    const perRecord = attachmentMap[col.id] ?? {};
    rows.forEach((record, rIdx) => {
      const excelRow = rIdx + headerRows + 1;
      const infos = perRecord[record.recordId];
      if (infos?.length) {
        jobs.push({ excelRow, colIndex, infos });
        totalImages += infos.length;
      } else {
        const text = formatCellText(record, col);
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

  // 根据图片高度设置行高（px → point：×0.75），不低于标准数据行高
  for (const [excelRow, pixels] of rowPixelHeights) {
    worksheet.getRow(excelRow).height = Math.max(DATA_ROW_HEIGHT, pixels * 0.75 + 4);
  }

  // 图片锚定到所属单元格内部：单元格索引 + EMU 像素偏移（oneCellAnchor）
  for (const img of placed) {
    const imageId = workbook.addImage({
      buffer: img.buffer as never,
      extension: img.ext,
    });
    worksheet.addImage(imageId, {
      tl: {
        // ExcelJS 约定：nativeCol/nativeRow 为单元格索引，nativeColOff/nativeRowOff 为 EMU 像素偏移
        nativeCol: img.colIndex,
        nativeColOff: IMG_GAP * PX_TO_EMU,
        nativeRow: img.excelRow - 1,
        nativeRowOff: img.offsetY * PX_TO_EMU,
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
