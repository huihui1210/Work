import { bitable, FieldType } from '@lark-base-open/js-sdk';
import type { IGridView, IRecord } from '@lark-base-open/js-sdk';

/** 批量读取记录的批次大小（接口单次上限 1000） */
const CHUNK_SIZE = 500;

export class ExportError extends Error {
  code: string;

  constructor(code: string) {
    super(code);
    this.code = code;
  }
}

export interface ExportColumn {
  id: string;
  name: string;
  type: FieldType;
}

export interface ExportData {
  columns: ExportColumn[];
  rows: IRecord[];
  tableName: string;
  viewName: string;
}

export interface SelectionInfo {
  /** 当前选中的记录数 */
  count: number;
  tableName: string;
  viewName: string;
  /** 当前视图是否支持读取多选记录（仅表格视图支持） */
  multiSelectSupported: boolean;
}

/**
 * 读取当前选中状态（用于面板上的实时提示）
 */
export async function getSelectionInfo(): Promise<SelectionInfo> {
  const selection = await bitable.base.getSelection();
  if (!selection.tableId || !selection.viewId) {
    throw new ExportError('NO_CONTEXT');
  }

  const table = await bitable.base.getTableById(selection.tableId);
  const view = await table.getViewById(selection.viewId);
  const gridView = view as unknown as IGridView;

  let count = 0;
  let multiSelectSupported = false;
  if (typeof gridView.getSelectedRecordIdList === 'function') {
    multiSelectSupported = true;
    const recordIds = await gridView.getSelectedRecordIdList();
    count = recordIds.length;
  } else if (selection.recordId) {
    count = 1;
  }

  const [tableName, viewName] = await Promise.all([table.getName(), view.getName()]);
  return { count, tableName, viewName, multiSelectSupported };
}

/**
 * 读取选中的记录及当前视图的可见列（按视图列顺序）
 */
export async function getSelectedData(): Promise<ExportData> {
  const selection = await bitable.base.getSelection();
  if (!selection.tableId || !selection.viewId) {
    throw new ExportError('NO_CONTEXT');
  }

  const table = await bitable.base.getTableById(selection.tableId);
  const view = await table.getViewById(selection.viewId);
  const gridView = view as unknown as IGridView;

  let recordIds: string[] = [];
  if (typeof gridView.getSelectedRecordIdList === 'function') {
    recordIds = await gridView.getSelectedRecordIdList();
  } else if (selection.recordId) {
    recordIds = [selection.recordId];
  }

  const [tableName, viewName, fieldMetas, visibleFieldIds] = await Promise.all([
    table.getName(),
    view.getName(),
    view.getFieldMetaList(),
    view.getVisibleFieldIdList(),
  ]);

  const visibleSet = new Set(visibleFieldIds);
  const columns: ExportColumn[] = fieldMetas
    .filter((field) => visibleSet.has(field.id))
    .map((field) => ({ id: field.id, name: field.name, type: field.type }));

  const rows: IRecord[] = [];
  for (let i = 0; i < recordIds.length; i += CHUNK_SIZE) {
    const chunkIds = recordIds.slice(i, i + CHUNK_SIZE);
    const values = await table.getRecordsByIds(chunkIds);
    values.forEach((value, index) => {
      rows.push({ recordId: chunkIds[index], fields: value.fields });
    });
  }

  return { columns, rows, tableName, viewName };
}

interface RawAttachment {
  name: string;
  size: number;
  token: string;
}

export interface AttachmentInfo {
  name: string;
  url: string;
}

/** fieldId -> recordId -> 该单元格的附件信息 */
export type AttachmentMap = Record<string, Record<string, AttachmentInfo[]>>;

function extractAttachments(value: unknown): RawAttachment[] {
  if (!Array.isArray(value)) return [];
  return value.filter(
    (item): item is RawAttachment =>
      !!item &&
      typeof item === 'object' &&
      typeof (item as RawAttachment).token === 'string' &&
      typeof (item as RawAttachment).size === 'number',
  );
}

/**
 * 收集所有附件单元格的可下载 URL（用于 Excel 内嵌图片）
 */
export async function getAttachmentMap(
  columns: ExportColumn[],
  rows: IRecord[],
): Promise<AttachmentMap> {
  const attachmentColumns = columns.filter((col) => col.type === FieldType.Attachment);
  if (attachmentColumns.length === 0 || rows.length === 0) return {};

  const selection = await bitable.base.getSelection();
  if (!selection.tableId) return {};
  const table = await bitable.base.getTableById(selection.tableId);

  const result: AttachmentMap = {};
  for (const col of attachmentColumns) {
    const perRecord: Record<string, AttachmentInfo[]> = {};
    for (const record of rows) {
      const attachments = extractAttachments(record.fields[col.id]);
      if (attachments.length === 0) continue;
      try {
        const urls = await table.getCellAttachmentUrls(
          attachments.map((item) => item.token),
          col.id,
          record.recordId,
        );
        perRecord[record.recordId] = urls.map((url, index) => ({
          name: attachments[index]?.name ?? `图片${index + 1}`,
          url,
        }));
      } catch {
        // 单格取 URL 失败时忽略，导出时该格降级为文件名文本
      }
    }
    result[col.id] = perRecord;
  }
  return result;
}

const SEGMENT_TYPES = new Set(['text', 'url', 'mention']);
const DATE_FIELD_TYPES = new Set<FieldType>([
  FieldType.DateTime,
  FieldType.CreatedTime,
  FieldType.ModifiedTime,
]);

/**
 * 将单元格的值转换为可读文本（用于 CSV / Excel / 剪贴板）
 */
export function cellToText(value: unknown, type: FieldType): string {
  if (value === null || value === undefined) return '';

  if (typeof value === 'boolean') return value ? '是' : '否';

  if (typeof value === 'number') {
    if (DATE_FIELD_TYPES.has(type)) return formatDate(value);
    return String(value);
  }

  if (typeof value === 'string') return value;

  if (Array.isArray(value)) return stringifyArray(value, type);

  return stringifyObject(value, type);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function asRecordArray(arr: unknown[]): Record<string, unknown>[] | null {
  return arr.every(isRecord) ? (arr as Record<string, unknown>[]) : null;
}

function stringifyArray(arr: unknown[], type: FieldType): string {
  if (arr.length === 0) return '';

  const records = asRecordArray(arr);
  if (records) {
    // 多行文本 / 超链接字段：文本片段数组
    if (records.every((item) => SEGMENT_TYPES.has(String(item.type)))) {
      return records
        .map((item) => {
          if (item.type === 'url') {
            const text = String(item.text ?? '');
            const link = String(item.link ?? '');
            return text && link && text !== link ? `${text} ${link}` : text || link;
          }
          return String(item.text ?? '');
        })
        .join('');
    }

    // 附件：名称列表
    if (records.every((item) => 'token' in item && 'size' in item)) {
      return records.map((item) => String(item.name ?? '')).filter(Boolean).join(', ');
    }

    // 多选：选项文本列表
    if (records.every((item) => 'text' in item && 'id' in item && !('recordIds' in item))) {
      return records.map((item) => String(item.text ?? '')).filter(Boolean).join(', ');
    }

    // 人员
    if (records.every((item) => 'email' in item || 'enName' in item)) {
      return records.map((item) => String(item.name ?? '')).filter(Boolean).join(', ');
    }

    // 群组
    if (records.every((item) => 'avatarUrl' in item)) {
      return records.map((item) => String(item.name ?? '')).filter(Boolean).join(', ');
    }
  }

  // 公式 / 查找引用等：递归展开
  return arr.map((item) => cellToText(item, type)).filter(Boolean).join(', ');
}

function stringifyObject(obj: unknown, type: FieldType): string {
  if (!isRecord(obj)) return String(obj ?? '');

  // 自动编号 { value, status }
  if ('value' in obj && 'status' in obj) return String(obj.value ?? '');

  // 关联记录 { text, recordIds, tableId }
  if ('recordIds' in obj && 'tableId' in obj) return String(obj.text ?? '');

  // 单选 { id, text }
  if ('text' in obj && 'id' in obj) return String(obj.text ?? '');

  // 地理位置
  if ('fullAddress' in obj || 'location' in obj) {
    return String(obj.fullAddress || obj.name || obj.location || '');
  }

  // 文本片段作为单值出现
  if (SEGMENT_TYPES.has(String(obj.type))) {
    return String(obj.text ?? obj.link ?? '');
  }

  // 其它对象兜底
  if ('text' in obj) return String(obj.text);
  if ('name' in obj) return String(obj.name);
  void type;
  return '';
}

function pad2(value: number): string {
  return String(value).padStart(2, '0');
}

function formatDate(ms: number): string {
  const date = new Date(ms);
  const datePart = `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  if (date.getHours() || date.getMinutes() || date.getSeconds()) {
    return `${datePart} ${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
  }
  return datePart;
}
