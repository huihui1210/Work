import type { IRecord } from '@lark-base-open/js-sdk';
import type { ExportColumn } from './bitable-helper';
import {
  buildOrderModel,
  resolveFields,
  sortRows,
  toRowMap,
} from './work-order-core';
import type { OrderModel } from './work-order-core';

/**
 * 设备缺陷处理工单 PDF 生成（样式参考「工作派工单」模板）：
 * A4 一页上下两张工单、中间虚线分隔；全局黑白简约、无背景填充；
 * 严格按照字段显示——表格中不存在的字段不渲染。
 * 字段匹配与内容组装见 work-order-core，本文件只负责 DOM 渲染与截图。
 */

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

/** 标签-值对行：多个字段并排一行；默认无值的字段不显示（keepEmpty 可强制保留），全部为空时返回 null */
function pairRow(pairs: { label: string; value: string; keepEmpty?: boolean }[]): HTMLElement | null {
  const filtered = pairs.filter((pair) => pair.value || pair.keepEmpty);
  if (!filtered.length) return null;
  const row = h('div', 'wo-row');
  filtered.forEach((pair, idx) => {
    row.append(h('div', 'wo-label wo-pair-label', pair.label));
    const value = h('div', 'wo-value', pair.value || '　');
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
  if (extraClass) cell.classList.add(...extraClass.split(/\s+/).filter(Boolean));
  row.append(cell);
  return row;
}

function appendIf(parent: HTMLElement, child: HTMLElement | null): void {
  if (child) parent.append(child);
}

function renderOrder(model: OrderModel): HTMLElement {
  const { get } = model;
  const order = h('div', 'wo-order');

  // 公司抬头 + 标题 + 顶部信息（发现时间 / 编号，存在才显示）
  order.append(h('div', 'wo-company', '福建LNG接收站'));
  order.append(h('div', 'wo-title', '设备缺陷处理工单'));
  const topMeta = h('div', 'wo-topmeta');
  if (model.discoverDate) topMeta.append(h('span', undefined, `发现时间：${model.discoverDate}`));
  if (model.defectNo) topMeta.append(h('span', undefined, `编号：${model.defectNo}`));
  if (topMeta.childElementCount) order.append(topMeta);

  // 表格主体
  const table = h('div', 'wo-table');

  // 基本信息：缺陷等级 / 发现人 / 所属岗位
  appendIf(
    table,
    pairRow([
      { label: '缺陷等级', value: get('level') },
      { label: '发现人', value: get('finder') },
      { label: '所属岗位', value: get('post') },
    ]),
  );

  // 部门与专业：一行
  appendIf(
    table,
    pairRow([
      { label: '责任部门', value: get('dept') },
      { label: '专业', value: get('specialty') },
    ]),
  );

  // 缺陷描述：区域部位 / 设备名称 / 位号 与缺陷描述合并为一个单元格（水平/垂直居中）
  appendIf(table, fieldRow('缺陷描述', model.descText, 'wo-desc-value'));

  // 日期：计划期限 / 消项时间（只到日；空值对由 pairRow 自动过滤）
  appendIf(
    table,
    pairRow([
      { label: '计划期限', value: model.deadlineDate },
      { label: '消项时间', value: model.closeDate },
    ]),
  );

  // 人员：党员责任人 / 负责人 / 处理人（一行）
  appendIf(
    table,
    pairRow([
      { label: '党员责任人', value: get('partyMember') },
      { label: '负责人', value: get('manager') },
      { label: '处理人', value: get('handler') },
    ]),
  );

  // 是否类状态：勾选形式（空值不显示）
  appendIf(
    table,
    pairRow(
      model.yesNoItems
        .filter((item) => item.value)
        .map((item) => ({ label: item.name, value: item.value })),
    ),
  );

  // 消缺处理情况
  appendIf(table, fieldRow('消缺处理情况', get('repairNote'), 'wo-repair-value wo-left'));

  // 风险控制措施 + 备注：一行（字段存在即显示，无值留空）
  if (get('risk') || get('remark')) {
    appendIf(
      table,
      pairRow([
        { label: '风险控制措施', value: get('risk'), keepEmpty: true },
        { label: '备注', value: get('remark'), keepEmpty: true },
      ]),
    );
  }

  // 制约因素：填写形式（字段存在即显示整行，供填写；有值则带出）
  if (model.constraintItems.length) {
    const rowEl = h('div', 'wo-row');
    rowEl.append(h('div', 'wo-label', '制约因素'));
    rowEl.append(h('div', 'wo-value wo-left', model.constraintText || '　'));
    table.append(rowEl);
  }

  // 其他信息：未归入固定位置且有值的字段，逐行显示
  for (const item of model.leftovers) {
    appendIf(table, fieldRow(item.name, item.value, 'wo-left'));
  }

  // 签字区（表单固定结构，始终显示）：消缺人 / 班长；姓名手写体，日期显示消项日期
  const signRow = h('div', 'wo-row wo-sign');
  const signCells = [
    { role: '消缺人', name: get('repairer') },
    { role: '班长', name: get('leader') },
  ];
  signCells.forEach((item, idx) => {
    const cell = h('div', 'wo-sign-cell');
    if (idx < signCells.length - 1) cell.classList.add('wo-bordered');
    const top = h('div', 'wo-sign-top', `${item.role}签字：`);
    if (item.name) top.append(h('span', 'wo-handwriting', item.name));
    cell.append(top);
    cell.append(h('div', 'wo-sign-bottom', `日期：${model.closeDate || '　　　年　　月　　日'}`));
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
/* 一页两张工单：上下平分，各自在半页内整体水平/垂直居中，中间虚线分隔 */
.wo-order {
  flex: 1 1 0; min-height: 0; display: flex; flex-direction: column;
  justify-content: center;
}
.wo-divider { flex: 0 0 auto; border-top: 2px dashed #888; margin: 10px 0; }
.wo-company { text-align: center; font-size: 15px; font-weight: 700; letter-spacing: 5px; margin-bottom: 2px; }
.wo-title { text-align: center; font-size: 21px; font-weight: 800; letter-spacing: 8px; margin-bottom: 8px; }
.wo-topmeta { display: flex; justify-content: space-between; gap: 12px; margin-bottom: 5px; }
.wo-table { border: 1.2px solid #333; }
.wo-row { display: flex; border-bottom: 1px solid #333; min-height: 26px; }
.wo-row:last-child { border-bottom: none; }
.wo-label {
  flex: 0 0 84px; display: flex; align-items: center; justify-content: center;
  padding: 5px 4px; font-weight: 700; border-right: 1px solid #333;
  text-align: center; line-height: 1.4; white-space: nowrap;
}
.wo-value {
  flex: 1; min-width: 0; padding: 5px 8px; display: flex; align-items: center;
  justify-content: center; text-align: center;
  white-space: pre-wrap; word-break: break-word; line-height: 1.55;
}
.wo-left { justify-content: flex-start; text-align: left; align-items: flex-start; }
.wo-bordered { border-right: 1px solid #333; }
.wo-desc-value { min-height: 34px; }
.wo-repair-value { min-height: 56px; }
.wo-sign-cell {
  flex: 1; display: flex; flex-direction: column; justify-content: center;
  align-items: center; gap: 10px; padding: 8px 10px; min-height: 58px;
}
.wo-sign-top { font-weight: 700; }
.wo-handwriting {
  font-family: "STXingkai", "Xingkai SC", "KaiTi", "楷体", "Kaiti SC", cursive;
  font-weight: 400; font-size: 17px; margin-left: 2px; color: #1a1a1a;
}
.wo-sign-bottom { letter-spacing: 1px; }
`;

/* ---------------- 主流程 ---------------- */

export async function buildWorkOrderPdfBlob(
  columns: ExportColumn[],
  rows: IRecord[],
  onProgress?: (done: number, total: number) => void,
): Promise<Blob> {
  const fieldIds = resolveFields(columns);
  const columnsById = new Map(columns.map((col) => [col.id, col]));
  const rowMaps = sortRows(
    rows.map((record) => toRowMap(columns, record)),
    fieldIds.defectNo,
  );
  const orders = rowMaps.map((row) => renderOrder(buildOrderModel(row, fieldIds, columnsById)));

  // 离屏容器
  const root = h('div', 'wo-root');
  const style = document.createElement('style');
  style.textContent = WO_STYLES;
  root.append(style);

  // 一页 A4 放两张工单，中间虚线分隔；奇数张时下半区留空占位（工单只占上半区）
  const pages: HTMLElement[] = [];
  for (let i = 0; i < orders.length; i += 2) {
    const page = h('div', 'wo-page');
    page.append(orders[i]);
    if (orders[i + 1]) {
      page.append(h('div', 'wo-divider'));
      page.append(orders[i + 1]);
    } else {
      page.append(h('div', 'wo-order wo-order-empty'));
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
        scale: 3,
        backgroundColor: '#ffffff',
        useCORS: true,
      });
      const image = canvas.toDataURL('image/jpeg', 0.98);
      if (i > 0) pdf.addPage();
      pdf.addImage(image, 'JPEG', 0, 0, pageWidth, pageHeight);
      onProgress?.(i + 1, pages.length);
    }

    return pdf.output('blob');
  } finally {
    root.remove();
  }
}
