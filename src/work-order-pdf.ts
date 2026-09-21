import type { IRecord } from '@lark-base-open/js-sdk';
import { cellToText } from './bitable-helper';
import type { ExportColumn } from './bitable-helper';

/**
 * 设备缺陷处理工单 PDF 生成（样式参考「工作派工单」模板）：
 * A4 一页上下两张工单、中间虚线分隔；全局黑白简约、无背景填充；
 * 严格按照字段显示——表格中不存在的字段不渲染；
 * 区域部位 / 设备名称 / 位号 与缺陷描述合并为一个「缺陷描述」单元格。
 */

type RowMap = Record<string, string>;

interface FieldDef {
  key: string;
  /** 匹配字段名的关键词（去空格、忽略大小写后包含匹配） */
  match: string[];
}

/** 工单固定字段定义；未匹配到的字段作为独立行显示，不丢数据 */
const FIELD_DEFS: FieldDef[] = [
  { key: 'defectNo', match: ['缺陷编号', '缺陷单号', '工单编号', '工单号', '编号', 'id'] },
  { key: 'finder', match: ['发现人', '上报人', '报告人'] },
  { key: 'post', match: ['所属岗位', '岗位', '班值'] },
  { key: 'level', match: ['分类定级', '缺陷等级', '定级', '分类', '等级'] },
  { key: 'desc', match: ['缺陷描述', '缺陷内容', '设备缺陷', '描述'] },
  { key: 'repairNote', match: ['消缺情况', '处理情况', '消缺简记', '处理记录', '维修记录'] },
  { key: 'repairer', match: ['消缺人', '消缺负责人', '维修负责人', '处理人'] },
  { key: 'leader', match: ['班长', '班组长'] },
  { key: 'closeTime', match: ['消项时间', '消缺完成日期', '完成日期', '消缺时间', '完成时间', '关闭时间'] },
  { key: 'remark', match: ['备注', '说明'] },
];

/** 与缺陷描述合并显示的字段（区域部位 / 设备名称 / 位号） */
const DESC_PART_FIELDS = ['区域部位', '设备名称', '位号', '区域', '位置'];

/** 制约因素字段（有值则勾选） */
const CONSTRAINT_FIELDS = ['方案', '备件', '条件', '人员', '窗口'];

const META_EXTRA_FIELDS = ['党员责任人', '是否消缺', '是否逾期'];

function normalize(name: string): string {
  return name.replace(/[\s/*\\]/g, '').toLowerCase();
}

/** 按字段名模糊匹配列 id */
function resolveFields(columns: ExportColumn[]): Record<string, string | undefined> {
  const result: Record<string, string | undefined> = {};
  const usedColumnIds = new Set<string>();

  for (const def of FIELD_DEFS) {
    const hit = columns.find(
      (col) =>
        !usedColumnIds.has(col.id) &&
        def.match.some((keyword) => normalize(col.name).includes(normalize(keyword))),
    );
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

  // 其他固定小字段
  for (const name of META_EXTRA_FIELDS) {
    const hit = columns.find(
      (col) => !usedColumnIds.has(col.id) && normalize(col.name) === normalize(name),
    );
    if (hit) {
      result[`extra_${name}`] = hit.id;
      usedColumnIds.add(hit.id);
    }
  }

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

function toRowMap(columns: ExportColumn[], record: IRecord): RowMap {
  const map: RowMap = {};
  for (const column of columns) {
    map[column.id] = cellToText(record.fields[column.id] ?? null, column.type).trim();
  }
  return map;
}

/** 从缺陷编号前 8 位解析发现日期 YYYY-MM-DD */
function parseDiscoverDate(defectNo: string): string {
  const match = /^(\d{4})(\d{2})(\d{2})/.exec(defectNo);
  if (!match) return '';
  return `${match[1]}-${match[2]}-${match[3]}`;
}

/** 按缺陷编号升序排序（编号缺失时保持原顺序） */
function sortRows(rows: RowMap[], defectNoId: string | undefined): RowMap[] {
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

/* ---------------- DOM 渲染 ---------------- */

function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== undefined) element.textContent = text;
  return element;
}

/** 标签-值对行：多个字段并排一行，无值的字段不显示；全部无值时返回 null */
function pairRow(pairs: { label: string; value: string }[]): HTMLElement | null {
  const filtered = pairs.filter((pair) => pair.value);
  if (!filtered.length) return null;
  const row = h('div', 'wo-row');
  filtered.forEach((pair, idx) => {
    row.append(h('div', 'wo-label wo-pair-label', pair.label));
    const value = h('div', 'wo-value', pair.value);
    if (idx < filtered.length - 1) value.classList.add('wo-bordered');
    row.append(value);
  });
  return row;
}

/** 整行区块：左侧标签 + 右侧内容；无值时返回 null（严格按照字段显示） */
function fieldRow(label: string, value: string, extraClass?: string): HTMLElement | null {
  if (!value) return null;
  const row = h('div', 'wo-row');
  row.append(h('div', 'wo-label', label));
  const cell = h('div', 'wo-value', value);
  if (extraClass) cell.classList.add(extraClass);
  row.append(cell);
  return row;
}

function renderOrder(
  row: RowMap,
  fieldIds: Record<string, string | undefined>,
  columnsById: Map<string, ExportColumn>,
): HTMLElement {
  const order = h('div', 'wo-order');
  const get = (key: string) => (fieldIds[key] ? row[fieldIds[key]!] ?? '' : '');

  const defectNo = get('defectNo');
  const discoverDate = parseDiscoverDate(defectNo);

  // 标题 + 顶部信息（发现时间 / 编号，存在才显示）
  order.append(h('div', 'wo-title', '设备缺陷处理工单'));
  const topMeta = h('div', 'wo-topmeta');
  if (discoverDate) topMeta.append(h('span', undefined, `发现时间：${discoverDate}`));
  if (defectNo) topMeta.append(h('span', undefined, `编号：${defectNo}`));
  if (topMeta.childElementCount) order.append(topMeta);

  // 表格主体
  const table = h('div', 'wo-table');

  const metaRow = pairRow([
    { label: '缺陷等级', value: get('level') },
    { label: '发现人', value: get('finder') },
    { label: '所属岗位', value: get('post') },
  ]);
  if (metaRow) table.append(metaRow);

  // 缺陷描述：区域部位 / 设备名称 / 位号 与缺陷描述合并为一个单元格
  const descPartIds: string[] = JSON.parse(fieldIds.descPartIds ?? '[]');
  const parts = descPartIds.map((id) => row[id] ?? '').filter(Boolean);
  const descText = [parts.join(' / '), get('desc')].filter(Boolean).join('\n');
  const descRow = fieldRow('缺陷描述', descText, 'wo-desc-value');
  if (descRow) table.append(descRow);

  // 消缺处理情况（备注并入）
  const repairText = [get('repairNote'), get('remark') ? `备注：${get('remark')}` : '']
    .filter(Boolean)
    .join('\n');
  const repairRow = fieldRow('消缺处理情况', repairText, 'wo-repair-value');
  if (repairRow) table.append(repairRow);

  // 其他信息：未归入固定位置且有值的字段，逐行显示
  const leftoverIds: string[] = JSON.parse(fieldIds.leftovers ?? '[]');
  for (const id of leftoverIds) {
    const value = (row[id] ?? '').trim();
    if (!value) continue;
    const rowEl = fieldRow(columnsById.get(id)?.name ?? '其他', value);
    if (rowEl) table.append(rowEl);
  }

  // 状态信息（党员责任人 / 是否消缺 / 是否逾期，存在且有值才显示）
  const statusRow = pairRow(
    META_EXTRA_FIELDS.map((name) => ({ label: name, value: get(`extra_${name}`) })),
  );
  if (statusRow) table.append(statusRow);

  // 制约因素（表中存在该组字段才显示整行，有值勾选）
  const constraintItems = CONSTRAINT_FIELDS.filter((name) => fieldIds[`constraint_${name}`]);
  if (constraintItems.length) {
    const rowEl = h('div', 'wo-row');
    rowEl.append(h('div', 'wo-label', '制约因素'));
    const cell = h('div', 'wo-value wo-constraints');
    for (const name of constraintItems) {
      const v = get(`constraint_${name}`);
      const chip = h('span', 'wo-constraint');
      chip.append(h('span', 'wo-checkbox', v ? '☑' : '☐'), h('span', undefined, name));
      if (v) {
        chip.append(h('span', 'wo-constraint-text', v));
        chip.classList.add('checked');
      }
      cell.append(chip);
    }
    rowEl.append(cell);
    table.append(rowEl);
  }

  // 消项时间（有值才显示）
  const closeRow = pairRow([{ label: '消项时间', value: get('closeTime') }]);
  if (closeRow) table.append(closeRow);

  // 签字区（表单固定结构，始终显示）
  const signRow = h('div', 'wo-row wo-sign');
  const signCells = [
    { role: '消缺人（处理负责人）', name: get('repairer') },
    { role: '班长（验收确认）', name: get('leader') },
  ];
  signCells.forEach((item, idx) => {
    const cell = h('div', 'wo-sign-cell');
    if (idx < signCells.length - 1) cell.classList.add('wo-bordered');
    const top = h('div', 'wo-sign-top');
    top.append(h('span', 'wo-sign-role', item.role));
    if (item.name) top.append(h('span', undefined, item.name));
    cell.append(top);
    cell.append(h('div', 'wo-sign-bottom', '签字：　　　　　　日期：　　　年　　月　　日'));
    signRow.append(cell);
  });
  table.append(signRow);

  order.append(table);
  return order;
}

const WO_STYLES = `
.wo-root { position: absolute; left: -10000px; top: 0; }
.wo-page {
  width: 794px; height: 1123px; box-sizing: border-box;
  padding: 26px 42px; background: #fff; color: #1a1a1a;
  font-family: "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", sans-serif;
  font-size: 12px; display: flex; flex-direction: column;
}
/* 一页两张工单：上下平分，中间虚线分隔 */
.wo-order { flex: 1 1 0; min-height: 0; display: flex; flex-direction: column; }
.wo-divider { flex: 0 0 auto; border-top: 2px dashed #888; margin: 10px 0; }
.wo-title { text-align: center; font-size: 21px; font-weight: 800; letter-spacing: 8px; margin-bottom: 8px; }
.wo-topmeta { display: flex; justify-content: space-between; gap: 12px; margin-bottom: 5px; }
.wo-table { border: 1.2px solid #333; }
.wo-row { display: flex; border-bottom: 1px solid #333; min-height: 26px; }
.wo-row:last-child { border-bottom: none; }
.wo-label {
  flex: 0 0 92px; display: flex; align-items: center; justify-content: center;
  padding: 5px 6px; font-weight: 700; border-right: 1px solid #333;
  text-align: center; line-height: 1.4;
}
.wo-value {
  flex: 1; min-width: 0; padding: 5px 8px; display: flex; align-items: center;
  white-space: pre-wrap; word-break: break-word; line-height: 1.55;
}
.wo-bordered { border-right: 1px solid #333; }
.wo-desc-value { align-items: flex-start; min-height: 78px; }
.wo-repair-value { align-items: flex-start; min-height: 56px; }
.wo-constraints { flex-wrap: wrap; gap: 4px 14px; }
.wo-constraint { display: inline-flex; align-items: center; gap: 3px; }
.wo-checkbox { font-size: 13px; }
.wo-constraint.checked { font-weight: 700; }
.wo-constraint-text { color: #444; font-weight: 400; }
.wo-sign-cell {
  flex: 1; display: flex; flex-direction: column; justify-content: space-between;
  gap: 8px; padding: 7px 10px; min-height: 58px;
}
.wo-sign-top { display: flex; gap: 6px; align-items: baseline; }
.wo-sign-role { font-weight: 700; }
.wo-sign-bottom { letter-spacing: 1px; }
`;

/* ---------------- 主流程 ---------------- */

export async function buildWorkOrderPdfBlob(
  columns: ExportColumn[],
  rows: IRecord[],
  onProgress?: (done: number, total: number) => void,
): Promise<Blob> {
  const fieldIds = resolveFields(columns);
  const rowMaps = sortRows(
    rows.map((record) => toRowMap(columns, record)),
    fieldIds.defectNo,
  );

  // 离屏容器
  const root = h('div', 'wo-root');
  const style = document.createElement('style');
  style.textContent = WO_STYLES;
  root.append(style);

  const columnsById = new Map(columns.map((col) => [col.id, col]));
  const orders = rowMaps.map((row) => renderOrder(row, fieldIds, columnsById));

  // 一页 A4 放两张工单，中间虚线分隔（奇数张时最后一页只放一张）
  const pages: HTMLElement[] = [];
  for (let i = 0; i < orders.length; i += 2) {
    const page = h('div', 'wo-page');
    page.append(orders[i]);
    if (orders[i + 1]) {
      page.append(h('div', 'wo-divider'));
      page.append(orders[i + 1]);
    }
    pages.push(page);
  }
  pages.forEach((page) => root.append(page));
  document.body.append(root);

  try {
    const [{ jsPDF }, html2canvas] = await Promise.all([
      import('jspdf'),
      import('html2canvas-pro'),
    ]);

    const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    for (let i = 0; i < pages.length; i++) {
      const canvas = await html2canvas.default(pages[i], {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
      });
      const image = canvas.toDataURL('image/jpeg', 0.92);
      if (i > 0) pdf.addPage();
      pdf.addImage(image, 'JPEG', 0, 0, pageWidth, pageHeight);
      onProgress?.(i + 1, pages.length);
    }

    return pdf.output('blob');
  } finally {
    root.remove();
  }
}
