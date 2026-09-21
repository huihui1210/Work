import type { IRecord } from '@lark-base-open/js-sdk';
import { cellToText } from './bitable-helper';
import type { ExportColumn } from './bitable-helper';

/**
 * 工单（PDF / Word）共享数据层：
 * 负责把多维表格记录转换为与渲染无关的「工单视图模型」，
 * DOM 渲染（work-order-pdf）与 OOXML 渲染（work-order-docx）只消费同一份模型。
 */

export type RowMap = Record<string, string>;

interface FieldDef {
  key: string;
  /** 匹配字段名的关键词（去空格、忽略大小写后包含匹配） */
  match: string[];
  /** 为 true 时按字段名完全相等匹配（避免「负责人」误匹配「消缺负责人」等） */
  exact?: boolean;
}

/** 工单固定字段定义；未匹配到的字段作为独立行显示，不丢数据 */
const FIELD_DEFS: FieldDef[] = [
  { key: 'defectNo', match: ['缺陷编号', '缺陷单号', '工单编号', '工单号', '编号', 'id'] },
  { key: 'finder', match: ['发现人', '上报人', '报告人'] },
  { key: 'post', match: ['所属岗位', '岗位', '班值'] },
  { key: 'level', match: ['分类定级', '缺陷等级', '定级', '分类', '等级'] },
  { key: 'desc', match: ['缺陷描述', '缺陷内容', '设备缺陷', '描述'] },
  { key: 'deadline', match: ['计划期限', '计划完成', '期限'] },
  { key: 'repairNote', match: ['消缺情况', '处理情况', '消缺简记', '处理记录', '维修记录'] },
  { key: 'partyMember', match: ['党员责任人'], exact: true },
  { key: 'manager', match: ['负责人'], exact: true },
  { key: 'handler', match: ['处理人'], exact: true },
  { key: 'dept', match: ['责任部门'] },
  { key: 'specialty', match: ['专业'], exact: true },
  { key: 'risk', match: ['风险控制措施', '风险措施', '安全措施'] },
  { key: 'repairer', match: ['消缺人'] },
  { key: 'leader', match: ['班长', '班组长'] },
  { key: 'closeTime', match: ['消项时间', '消缺完成日期', '完成日期', '消缺时间', '完成时间', '关闭时间'] },
  { key: 'remark', match: ['备注', '说明'] },
];

/** 与缺陷描述合并显示的字段（区域部位 / 设备名称 / 位号） */
const DESC_PART_FIELDS = ['区域部位', '设备名称', '位号', '区域', '位置'];

/** 制约因素字段（填写形式显示） */
export const CONSTRAINT_FIELDS = ['方案', '备件', '条件', '人员', '窗口'];

function normalize(name: string): string {
  return name.replace(/[\s/*\\]/g, '').toLowerCase();
}

/** 按字段名模糊匹配列 id */
export function resolveFields(columns: ExportColumn[]): Record<string, string | undefined> {
  const result: Record<string, string | undefined> = {};
  const usedColumnIds = new Set<string>();

  for (const def of FIELD_DEFS) {
    const hit = columns.find((col) => {
      if (usedColumnIds.has(col.id)) return false;
      return def.match.some((keyword) =>
        def.exact
          ? normalize(col.name) === normalize(keyword)
          : normalize(col.name).includes(normalize(keyword)),
      );
    });
    if (hit) {
      result[def.key] = hit.id;
      usedColumnIds.add(hit.id);
    }
  }

  // 制约因素
  for (const name of CONSTRAINT_FIELDS) {
    const hit = columns.find(
      (col) => !usedColumnIds.has(col.id) && normalize(col.name).includes(normalize(name)),
    );
    if (hit) {
      result[`constraint_${name}`] = hit.id;
      usedColumnIds.add(hit.id);
    }
  }

  // 是否类字段（是否消缺 / 是否逾期等）→ 勾选形式显示
  const yesNoHits = columns.filter(
    (col) => !usedColumnIds.has(col.id) && normalize(col.name).startsWith('是否'),
  );
  result.yesNoIds = JSON.stringify(yesNoHits.map((col) => col.id));
  for (const col of yesNoHits) usedColumnIds.add(col.id);

  // 缺陷描述合并来源字段（区域部位 / 设备名称 / 位号等）
  const descPartHits = columns.filter(
    (col) =>
      !usedColumnIds.has(col.id) &&
      DESC_PART_FIELDS.some((keyword) => normalize(col.name).includes(normalize(keyword))),
  );
  result.descPartIds = JSON.stringify(descPartHits.map((col) => col.id));
  for (const col of descPartHits) usedColumnIds.add(col.id);

  // 剩余字段作为其他信息
  result.leftovers = JSON.stringify(
    columns.filter((col) => !usedColumnIds.has(col.id)).map((col) => col.id),
  );

  return result;
}

export function toRowMap(columns: ExportColumn[], record: IRecord): RowMap {
  const map: RowMap = {};
  for (const column of columns) {
    map[column.id] = cellToText(record.fields[column.id] ?? null, column.type).trim();
  }
  return map;
}

/** 从缺陷编号前 8 位解析发现日期 YYYY-MM-DD */
export function parseDiscoverDate(defectNo: string): string {
  const match = /^(\d{4})(\d{2})(\d{2})/.exec(defectNo);
  if (!match) return '';
  return `${match[1]}-${match[2]}-${match[3]}`;
}

/** 按缺陷编号升序排序（编号缺失时保持原顺序） */
export function sortRows(rows: RowMap[], defectNoId: string | undefined): RowMap[] {
  if (!defectNoId) return rows;
  return rows
    .map((row, index) => ({ row, index }))
    .sort((a, b) => {
      const va = a.row[defectNoId];
      const vb = b.row[defectNoId];
      if (!va && !vb) return a.index - b.index;
      if (!va) return 1;
      if (!vb) return -1;
      const cmp = va.localeCompare(vb, undefined, { numeric: true });
      return cmp !== 0 ? cmp : a.index - b.index;
    })
    .map((item) => item.row);
}

/** 日期文本只保留到日（去掉时分） */
export function dateOnly(text: string): string {
  return text.replace(/[ T]\d{1,2}:\d{2}(:\d{2})?$/, '').trim();
}

/** 是否类字段 → 勾选文本：是/否 二选一勾选；空值返回空串（渲染端自行决定是否显示） */
export function yesNoText(value: string): string {
  const v = value.trim();
  if (!v) return '';
  const isYes = /^(是|已|完|y|true|1)/i.test(v);
  return isYes ? '☑ 是　　☐ 否' : '☐ 是　　☑ 否';
}

/** 是否类字段的渲染数据 */
export interface YesNoItem {
  name: string;
  value: string;
}

/** 其他信息（未归入固定位置且有值的字段） */
export interface LeftoverItem {
  name: string;
  value: string;
}

/** 与渲染无关的单张工单视图模型 */
export interface OrderModel {
  /** 按字段语义 key 取值（key 未匹配时返回空串） */
  get: (key: string) => string;
  /** 该语义字段在表中是否存在（区别于「有值」，用于需要留空待填的行） */
  has: (key: string) => boolean;
  defectNo: string;
  discoverDate: string;
  /** 计划期限（只到日） */
  deadlineDate: string;
  closeDate: string;
  /** 区域部位 / 设备名称 / 位号 与缺陷描述合并后的一行文本 */
  descText: string;
  /** 是否类字段（按视图顺序，可能含空值） */
  yesNoItems: YesNoItem[];
  /** 本表实际存在的制约因素字段名 */
  constraintItems: string[];
  /** 制约因素填写文本，如「方案：xxx　　备件：」 */
  constraintText: string;
  /** 有值的其他字段 */
  leftovers: LeftoverItem[];
}

/** 由一行记录构建工单视图模型（PDF / Word 共用，保证两边内容一致） */
export function buildOrderModel(
  row: RowMap,
  fieldIds: Record<string, string | undefined>,
  columnsById: Map<string, ExportColumn>,
): OrderModel {
  const get = (key: string): string => (fieldIds[key] ? row[fieldIds[key]!] ?? '' : '');
  const has = (key: string): boolean => !!fieldIds[key];

  const defectNo = get('defectNo');
  const discoverDate = parseDiscoverDate(defectNo);
  const deadlineDate = dateOnly(get('deadline'));
  const closeDate = dateOnly(get('closeTime'));

  // 缺陷描述：区域部位 / 设备名称 / 位号 与描述合并为一行
  const descPartIds: string[] = JSON.parse(fieldIds.descPartIds ?? '[]');
  const parts = descPartIds.map((id) => row[id] ?? '').filter(Boolean);
  const descText = [parts.join(' / '), get('desc')].filter(Boolean).join('　');

  // 是否类字段
  const yesNoIds: string[] = JSON.parse(fieldIds.yesNoIds ?? '[]');
  const yesNoItems = yesNoIds.map((id) => ({
    name: columnsById.get(id)?.name ?? '状态',
    value: yesNoText(row[id] ?? ''),
  }));

  // 制约因素（字段存在即列出，无值留空供手写）
  const constraintItems = CONSTRAINT_FIELDS.filter((name) => fieldIds[`constraint_${name}`]);
  const constraintText = constraintItems
    .map((name) => {
      const v = get(`constraint_${name}`);
      return v ? `${name}：${v}` : `${name}：`;
    })
    .join('　　');

  // 其他信息：仅含非空字段
  const leftoverIds: string[] = JSON.parse(fieldIds.leftovers ?? '[]');
  const leftovers = leftoverIds
    .map((id) => ({ name: columnsById.get(id)?.name ?? '其他', value: (row[id] ?? '').trim() }))
    .filter((item) => item.value);

  return {
    get,
    has,
    defectNo,
    discoverDate,
    deadlineDate,
    closeDate,
    descText,
    yesNoItems,
    constraintItems,
    constraintText,
    leftovers,
  };
}
