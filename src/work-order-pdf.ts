import type { IRecord } from '@lark-base-open/js-sdk';
import { cellToText } from './bitable-helper';
import type { ExportColumn } from './bitable-helper';

/**
 * 设备缺陷处理工单 PDF 生成
 * 参考电力行业「设备缺陷通知单」与维修工单闭环（closeout）格式：
 * 一记录一页（A4），按缺陷编号排序，底部为消缺人 / 班长确认签字位。
 */

type RowMap = Record<string, string>;

interface FieldDef {
  key: string;
  /** 匹配字段名的关键词（去空格、忽略大小写后包含匹配） */
  match: string[];
}

/** 工单固定字段定义；未匹配到的字段自动进入「其他信息」，不丢数据 */
const FIELD_DEFS: FieldDef[] = [
  { key: 'defectNo', match: ['缺陷编号', '缺陷单号', '工单编号', '工单号', '编号', 'id'] },
  { key: 'finder', match: ['发现人', '上报人', '报告人'] },
  { key: 'post', match: ['所属岗位', '岗位', '班值'] },
  { key: 'level', match: ['分类定级', '缺陷等级', '定级', '分类', '等级'] },
  { key: 'location', match: ['区域部位', '设备名称', '位号', '区域', '位置'] },
  { key: 'desc', match: ['缺陷描述', '缺陷内容', '设备缺陷', '描述'] },
  { key: 'repairNote', match: ['消缺情况', '处理情况', '消缺简记', '处理记录', '维修记录'] },
  { key: 'repairer', match: ['消缺人', '消缺负责人', '维修负责人', '处理人'] },
  { key: 'leader', match: ['班长', '班组长'] },
  { key: 'closeTime', match: ['消项时间', '消缺完成日期', '完成日期', '消缺时间', '完成时间', '关闭时间'] },
  { key: 'remark', match: ['备注', '说明'] },
];

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

function valueOrDash(cell: HTMLElement, value: string): void {
  if (value) {
    cell.textContent = value;
  } else {
    cell.textContent = '—';
    cell.classList.add('wo-empty');
  }
}

function metaRow(label1: string, value1: string, label2: string, value2: string): HTMLElement {
  const row = h('div', 'wo-meta-row');
  row.append(
    h('div', 'wo-meta-label', label1),
    h('div', 'wo-meta-value'),
    h('div', 'wo-meta-label', label2),
    h('div', 'wo-meta-value'),
  );
  const cells = row.querySelectorAll('.wo-meta-value');
  valueOrDash(cells[0] as HTMLElement, value1);
  valueOrDash(cells[1] as HTMLElement, value2);
  return row;
}

function fullRow(label: string, value: string): HTMLElement {
  const row = h('div', 'wo-meta-row');
  row.append(h('div', 'wo-meta-label wo-full-label', label));
  const valueCell = h('div', 'wo-meta-value wo-full-value');
  valueOrDash(valueCell, value);
  row.append(valueCell);
  return row;
}

function sectionBlock(title: string, value: string, minHeight: number): HTMLElement {
  const section = h('div', 'wo-section');
  const titleEl = h('div', 'wo-section-title', title);
  const body = h('div', 'wo-section-body');
  body.style.minHeight = `${minHeight}px`;
  if (value) {
    body.textContent = value;
  } else {
    body.textContent = '—';
    body.classList.add('wo-empty');
  }
  section.append(titleEl, body);
  return section;
}

function renderOrder(
  row: RowMap,
  fieldIds: Record<string, string | undefined>,
  columnsById: Map<string, ExportColumn>,
  pageIndex: number,
  total: number,
): HTMLElement {
  const page = h('div', 'wo-page');

  const get = (key: string) => (fieldIds[key] ? row[fieldIds[key]!] ?? '' : '');

  const defectNo = get('defectNo');
  const discoverDate = parseDiscoverDate(defectNo);

  // 抬头
  const header = h('div', 'wo-header');
  const org = h('div', 'wo-org', '福建LNG接收站');
  const title = h('div', 'wo-title', '设备缺陷处理工单');
  const no = h('div', 'wo-no', defectNo ? `NO. ${defectNo}` : '');
  header.append(org, title, no);
  page.append(header);

  // 基本信息
  const meta = h('div', 'wo-meta');
  meta.append(
    metaRow('发现时间', discoverDate, '发现人', get('finder')),
    metaRow('所属岗位', get('post'), '缺陷等级', get('level')),
  );

  // 状态小标签
  const statusChips: string[] = [];
  for (const name of META_EXTRA_FIELDS) {
    const v = get(`extra_${name}`);
    if (v) statusChips.push(`${name}：${v}`);
  }
  if (statusChips.length) {
    meta.append(fullRow('状态信息', statusChips.join('　')));
  }

  meta.append(fullRow('区域部位 / 设备名称 / 位号', get('location')));
  page.append(meta);

  // 缺陷描述
  page.append(sectionBlock('缺陷描述', get('desc'), 70));

  // 消缺处理情况（备注并入）
  const repairText = [get('repairNote'), get('remark') ? `备注：${get('remark')}` : '']
    .filter(Boolean)
    .join('\n');
  page.append(sectionBlock('消缺处理情况', repairText, 90));

  // 其他信息：未归入固定位置且有值的字段
  const leftoverIds: string[] = JSON.parse(fieldIds.leftovers ?? '[]');
  const leftoverPairs = leftoverIds
    .map((id) => ({ name: columnsById.get(id)?.name ?? id, value: row[id] ?? '' }))
    .filter((item) => item.value);
  if (leftoverPairs.length) {
    const extra = h('div', 'wo-section');
    extra.append(h('div', 'wo-section-title', '其他信息'));
    const body = h('div', 'wo-extra-body');
    for (const pair of leftoverPairs) {
      const item = h('div', 'wo-extra-item');
      item.append(h('span', 'wo-extra-name', pair.name), h('span', 'wo-extra-value', pair.value));
      body.append(item);
    }
    extra.append(body);
    page.append(extra);
  }

  // 制约因素
  const constraintSection = h('div', 'wo-constraints');
  constraintSection.append(h('span', 'wo-constraints-label', '制约因素：'));
  for (const name of CONSTRAINT_FIELDS) {
    const v = get(`constraint_${name}`);
    const chip = h('span', 'wo-constraint');
    chip.append(h('span', 'wo-checkbox', v ? '☑' : '☐'));
    chip.append(h('span', undefined, name));
    if (v) {
      chip.append(h('span', 'wo-constraint-text', v));
      chip.classList.add('checked');
    }
    constraintSection.append(chip);
  }
  page.append(constraintSection);

  // 确认签字区：消缺人、班长
  const sign = h('div', 'wo-sign');
  const signRows = [
    { role: '消缺人（处理负责人）', value: get('repairer') },
    { role: '班长（验收确认）', value: get('leader') },
  ];
  for (const item of signRows) {
    const rowEl = h('div', 'wo-sign-row');
    rowEl.append(h('span', 'wo-sign-role', item.role));
    const nameEl = h('span', 'wo-sign-name');
    if (item.value) {
      nameEl.textContent = item.value;
    } else {
      nameEl.textContent = '（待签字）';
      nameEl.classList.add('wo-empty');
    }
    rowEl.append(nameEl);
    rowEl.append(h('span', 'wo-sign-line', '签字：'));
    rowEl.append(h('span', 'wo-sign-date', '日期：　　 年 　 月 　 日'));
    sign.append(rowEl);
  }
  const closeRow = h('div', 'wo-sign-row wo-close-row');
  closeRow.append(h('span', 'wo-sign-role', '消项时间'), h('span', 'wo-sign-name', get('closeTime') || '—'));
  const closeNameEl = closeRow.querySelector('.wo-sign-name') as HTMLElement;
  if (!get('closeTime')) closeNameEl.classList.add('wo-empty');
  sign.append(closeRow);
  page.append(sign);

  // 页脚
  const footer = h('div', 'wo-footer');
  footer.append(
    h('span', undefined, `第 ${pageIndex} 页 / 共 ${total} 页`),
    h('span', undefined, `打印时间：${new Date().toLocaleString('zh-CN')}`),
  );
  page.append(footer);

  return page;
}

const WO_STYLES = `
.wo-root { position: absolute; left: -10000px; top: 0; }
.wo-page {
  width: 794px; min-height: 1123px; box-sizing: border-box;
  padding: 34px 40px 28px; background: #fff; color: #1a1a1a;
  font-family: "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", sans-serif;
  font-size: 14px; display: flex; flex-direction: column;
}
.wo-header { text-align: center; border-bottom: 3px solid #1f3a5f; padding-bottom: 14px; margin-bottom: 16px; position: relative; }
.wo-org { font-size: 17px; color: #1f3a5f; letter-spacing: 4px; font-weight: 700; }
.wo-title { font-size: 28px; font-weight: 800; color: #142942; letter-spacing: 10px; margin-top: 4px; }
.wo-no { position: absolute; right: 0; bottom: 12px; font-size: 14px; color: #333; font-weight: 600; }
.wo-meta { border: 1.5px solid #2c4a6e; }
.wo-meta-row { display: flex; border-bottom: 1px solid #9bb0c7; }
.wo-meta-row:last-child { border-bottom: none; }
.wo-meta-label { width: 118px; flex-shrink: 0; background: #eef3f9; color: #1f3a5f; font-weight: 600; display: flex; align-items: center; padding: 9px 12px; border-right: 1px solid #9bb0c7; }
.wo-meta-value { flex: 1; display: flex; align-items: center; padding: 9px 12px; min-width: 0; }
.wo-full-label { width: 190px; }
.wo-empty { color: #9aa5b1; }
.wo-section { border: 1.5px solid #2c4a6e; border-top: none; }
.wo-section-title { background: #eef3f9; color: #1f3a5f; font-weight: 700; padding: 7px 12px; border-bottom: 1px solid #9bb0c7; }
.wo-section-body { padding: 10px 12px; white-space: pre-wrap; line-height: 1.7; word-break: break-word; }
.wo-extra-body { display: flex; flex-wrap: wrap; padding: 6px 12px; }
.wo-extra-item { display: flex; width: 50%; box-sizing: border-box; padding: 5px 0; gap: 6px; }
.wo-extra-name { color: #1f3a5f; font-weight: 600; flex-shrink: 0; }
.wo-extra-value { color: #333; word-break: break-word; }
.wo-constraints { display: flex; flex-wrap: wrap; gap: 8px 18px; align-items: center; border: 1.5px solid #2c4a6e; border-top: none; padding: 10px 12px; }
.wo-constraints-label { color: #1f3a5f; font-weight: 700; }
.wo-constraint { display: inline-flex; align-items: center; gap: 4px; }
.wo-checkbox { font-size: 16px; color: #1f3a5f; }
.wo-constraint.checked { color: #14532d; font-weight: 600; }
.wo-constraint.checked .wo-checkbox { color: #15803d; }
.wo-constraint-text { color: #555; font-weight: 400; font-size: 12px; margin-left: 2px; }
.wo-sign { border: 1.5px solid #2c4a6e; border-top: none; padding: 6px 12px; }
.wo-sign-row { display: flex; align-items: center; gap: 8px; padding: 9px 0; border-bottom: 1px dashed #c3cedb; }
.wo-sign-row:last-child { border-bottom: none; }
.wo-sign-role { color: #1f3a5f; font-weight: 700; min-width: 170px; }
.wo-sign-name { min-width: 110px; font-weight: 600; }
.wo-sign-line { color: #444; }
.wo-sign-date { margin-left: auto; color: #444; letter-spacing: 1px; }
.wo-close-row .wo-sign-name { font-weight: 400; }
.wo-footer { margin-top: auto; padding-top: 12px; display: flex; justify-content: space-between; color: #7a8699; font-size: 12px; }
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
  const pages = rowMaps.map((row, index) =>
    renderOrder(row, fieldIds, columnsById, index + 1, rowMaps.length),
  );
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
