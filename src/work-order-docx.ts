import type { IRecord } from '@lark-base-open/js-sdk';
import {
  AlignmentType,
  BorderStyle,
  Document,
  DocumentGridType,
  HeightRule,
  LineRuleType,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from 'docx';
import type { ExportColumn } from './bitable-helper';
import { buildOrderModel, resolveFields, sortRows, toRowMap } from './work-order-core';
import type { OrderModel } from './work-order-core';

/**
 * 可编辑的 Word 格式缺陷处理工单（与 PDF 内容一致）：
 * A4 一页上下两张工单；表格文本可直接编辑；签字姓名使用楷体手写风格。
 * 内容组装见 work-order-core，本文件只负责 OOXML 排版。
 */

/* A4 内容区 6 列网格（twips）：标签 1500，三列标签共 4500，值均分剩余宽度 */
const PAGE_CONTENT_WIDTH = 10466;
const GRID = [1500, 1989, 1500, 1989, 1500, 1988];

/** 字号（half-point）：正文 9pt（与 PDF 12px 视觉一致） */
const BODY_SIZE = 18;
const COMPANY_SIZE = 24;
const TITLE_SIZE = 32;
/** 签字区手写体姓名 11pt */
const SIGN_NAME_SIZE = 22;

const FONT = { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' } as const;
const HANDWRITING_FONT = { ascii: '楷体', eastAsia: '楷体', hAnsi: '楷体' } as const;

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: '333333' };
const CELL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };
const TABLE_NO_BORDERS = {
  ...NO_BORDERS,
  insideHorizontal: NO_BORDER,
  insideVertical: NO_BORDER,
};

/** 各数据行最小高度（twips） */
const ROW_H = 300;
const ROW_H_TALL = 420;
const ROW_H_CONSTRAINT = 460;
const ROW_H_SIGN = 560;

interface CellSpec {
  text?: string;
  width: number;
  span?: number;
  bold?: boolean;
  align?: (typeof AlignmentType)[keyof typeof AlignmentType];
  size?: number;
}

function bodyRun(text: string, opts: { bold?: boolean; size?: number } = {}): TextRun {
  return new TextRun({
    text,
    bold: opts.bold ?? false,
    size: opts.size ?? BODY_SIZE,
    font: FONT,
  });
}

function makeCell(spec: CellSpec): TableCell {
  return new TableCell({
    columnSpan: spec.span,
    width: { size: spec.width, type: WidthType.DXA },
    borders: CELL_BORDERS,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 30, bottom: 30, left: 70, right: 70 },
    children: [
      new Paragraph({
        alignment: spec.align ?? AlignmentType.CENTER,
        spacing: { before: 0, after: 0, line: 260 },
        children: [bodyRun(spec.text || '', { bold: spec.bold, size: spec.size })],
      }),
    ],
  });
}

function makeRow(cells: TableCell[], minHeight?: number): TableRow {
  return new TableRow({
    cantSplit: true,
    height: minHeight ? { value: minHeight, rule: HeightRule.ATLEAST } : undefined,
    children: cells,
  });
}

/** 无边框空段落（用于表格之间防合并，行高压到 20twips） */
function makeTinyParagraph(): Paragraph {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: '', size: 2 })],
  });
}

/** 标签 + 值（值跨 span 列） */
function labelValue(
  label: string,
  value: string,
  labelWidth: number,
  valueWidth: number,
  span?: number,
): [TableCell, TableCell] {
  return [
    makeCell({ text: label, width: labelWidth, bold: true }),
    makeCell({ text: value, width: valueWidth, span }),
  ];
}

/** 抬头居中段落 */
function headingParagraph(text: string, size: number, after: number): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after },
    children: [new TextRun({ text, bold: true, size, font: FONT })],
  });
}

/** 构建单张工单的内容块（标题 + 顶部信息 + 主表），由外层容器负责分页与纵向定位 */
function buildOrderBlocks(model: OrderModel): (Paragraph | Table)[] {
  const { get, has } = model;
  const blocks: (Paragraph | Table)[] = [];

  // 抬头（公司 / 标题）
  blocks.push(headingParagraph('福建LNG接收站', COMPANY_SIZE, 20));
  blocks.push(headingParagraph('设备缺陷处理工单', TITLE_SIZE, 40));

  // 顶部信息（发现时间 / 编号），无边框两列表格
  if (model.discoverDate || model.defectNo) {
    blocks.push(
      new Table({
        width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
        columnWidths: [5233, 5233],
        borders: TABLE_NO_BORDERS,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 5233, type: WidthType.DXA },
                borders: NO_BORDERS,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [bodyRun(`发现时间：${model.discoverDate}`)],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 5233, type: WidthType.DXA },
                borders: NO_BORDERS,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [bodyRun(`编号：${model.defectNo}`)],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    );
    // 防止相邻表格被 Word 合并
    blocks.push(makeTinyParagraph());
  }

  /* ------- 主表格 ------- */
  const [c0, c1, c2, c3, c4, c5] = GRID;
  const valueAll = c1 + c2 + c3 + c4 + c5;
  const rows: TableRow[] = [];

  // 基本信息：缺陷等级 / 发现人 / 所属岗位
  rows.push(
    makeRow(
      [
        makeCell({ text: '缺陷等级', width: c0, bold: true }),
        makeCell({ text: get('level'), width: c1 }),
        makeCell({ text: '发现人', width: c2, bold: true }),
        makeCell({ text: get('finder'), width: c3 }),
        makeCell({ text: '所属岗位', width: c4, bold: true }),
        makeCell({ text: get('post'), width: c5 }),
      ],
      ROW_H,
    ),
  );

  // 部门与专业
  rows.push(
    makeRow([
      ...labelValue('责任部门', get('dept'), c0, c1 + c2 + c3, 3),
      ...labelValue('专业', get('specialty'), c4, c5),
    ]),
  );

  // 缺陷描述（区域部位 / 设备名称 / 位号 已在模型中合并为一行）
  rows.push(
    makeRow(
      [
        makeCell({ text: '缺陷描述', width: c0, bold: true }),
        makeCell({ text: model.descText, width: valueAll, span: 5 }),
      ],
      ROW_H_TALL,
    ),
  );

  // 日期：计划期限 / 消项时间（只到日；字段存在即显示，无值留空）
  if (has('deadline') || has('closeTime')) {
    rows.push(
      makeRow([
        ...labelValue('计划期限', model.deadlineDate, c0, c1 + c2, 2),
        ...labelValue('消项时间', model.closeDate, c3, c4 + c5, 2),
      ]),
    );
  }

  // 人员：党员责任人 / 负责人 / 处理人
  if (has('partyMember') || has('manager') || has('handler')) {
    rows.push(
      makeRow(
        [
          makeCell({ text: '党员责任人', width: c0, bold: true }),
          makeCell({ text: get('partyMember'), width: c1 }),
          makeCell({ text: '负责人', width: c2, bold: true }),
          makeCell({ text: get('manager'), width: c3 }),
          makeCell({ text: '处理人', width: c4, bold: true }),
          makeCell({ text: get('handler'), width: c5 }),
        ],
        ROW_H_CONSTRAINT,
      ),
    );
  }

  // 是否类字段（勾选形式），两两一行
  model.yesNoItems.forEach((item, idx) => {
    if (idx % 2 === 0) {
      const next = model.yesNoItems[idx + 1];
      if (next) {
        rows.push(
          makeRow([
            ...labelValue(item.name, item.value, c0, c1 + c2, 2),
            ...labelValue(next.name, next.value, c3, c4 + c5, 2),
          ]),
        );
      } else {
        rows.push(makeRow(labelValue(item.name, item.value, c0, valueAll, 5)));
      }
    }
  });

  // 消缺处理情况
  if (get('repairNote')) {
    rows.push(
      makeRow(
        [
          makeCell({ text: '消缺处理情况', width: c0, bold: true }),
          makeCell({ text: get('repairNote'), width: valueAll, span: 5, align: AlignmentType.LEFT }),
        ],
        ROW_H_TALL,
      ),
    );
  }

  // 风险控制措施 + 备注（字段存在即显示，无值留空）
  if (has('risk') || has('remark')) {
    rows.push(
      makeRow([
        ...labelValue('风险控制措施', get('risk'), c0, c1 + c2, 2),
        ...labelValue('备注', get('remark'), c3, c4 + c5, 2),
      ]),
    );
  }

  // 制约因素（填写形式，无值留空供手写）
  if (model.constraintItems.length) {
    rows.push(
      makeRow(
        [
          makeCell({ text: '制约因素', width: c0, bold: true }),
          makeCell({ text: model.constraintText, width: valueAll, span: 5, align: AlignmentType.LEFT }),
        ],
        ROW_H_CONSTRAINT,
      ),
    );
  }

  // 其他信息
  for (const item of model.leftovers) {
    rows.push(
      makeRow(
        [
          makeCell({ text: item.name, width: c0, bold: true }),
          makeCell({ text: item.value, width: valueAll, span: 5, align: AlignmentType.LEFT }),
        ],
        ROW_H_CONSTRAINT,
      ),
    );
  }

  // 签字区：消缺人签字 / 班长签字（姓名楷体手写风格）+ 日期（取消项日期）
  const signSpecs = [
    { role: '消缺人', name: get('repairer') },
    { role: '班长', name: get('leader') },
  ];
  rows.push(
    makeRow(
      signSpecs.map((item) => {
        const signRuns: TextRun[] = [bodyRun(`${item.role}签字：`, { bold: true })];
        if (item.name) {
          signRuns.push(new TextRun({ text: item.name, size: SIGN_NAME_SIZE, font: HANDWRITING_FONT }));
        }
        return new TableCell({
          columnSpan: 3,
          width: { size: c0 + c1 + c2, type: WidthType.DXA },
          borders: CELL_BORDERS,
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 30, bottom: 30, left: 70, right: 70 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 20, line: 260 },
              children: signRuns,
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0, line: 260 },
              children: [bodyRun(`日期：${model.closeDate || '　　　年　　月　　日'}`)],
            }),
          ],
        });
      }),
      ROW_H_SIGN,
    ),
  );

  blocks.push(
    new Table({
      width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: GRID,
      borders: {
        top: BORDER,
        bottom: BORDER,
        left: BORDER,
        right: BORDER,
        insideHorizontal: BORDER,
        insideVertical: BORDER,
      },
      rows,
    }),
  );

  return blocks;
}

/* ---------------- 页面容器（分页 / 虚线居中，Word 实测校准） ---------------- */

const PAGE_HEIGHT = 16838;
const PAGE_CENTER = PAGE_HEIGHT / 2; // 8419
/**
 * 半区固定高 7732twips（2×7732 = 15464）。
 * Word 对表格后必需的空段落标记有约 120twips 最小行高、无法压缩，
 * 因此上边距取 687：虚线位置 = 687 + 7732 = 8419，恰为 A4 物理正中线；
 * 容器底 16151 距页底 687，其中下边距 537 + 段落标记约 150，上下视觉留白对称。
 */
const HALF_PAGE_HEIGHT = 7732;
const MARGIN_TOP = PAGE_CENTER - HALF_PAGE_HEIGHT; // 687
const MARGIN_BOTTOM = 537;

/**
 * 一页容器：1 列 2 行的无边框表格，每行固定半页高，工单内容在半区内垂直居中。
 * 第二张工单存在时，上半区底边显示虚线；不存在时下半区留空（工单仍位于上半区）。
 * 第 2 页起在单元格内插入带 pageBreakBefore 的零高段落分页，不额外占行。
 */
function buildPageTable(
  first: (Paragraph | Table)[],
  second: (Paragraph | Table)[] | null,
  needPageBreak: boolean,
): Table {
  const makeHalfCell = (
    content: (Paragraph | Table)[] | null,
    withDashedBottom: boolean,
    breakBefore = false,
  ) => {
    const inner: (Paragraph | Table)[] = [];
    if (breakBefore) {
      inner.push(
        new Paragraph({
          pageBreakBefore: true,
          spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT },
          children: [new TextRun({ text: '', size: 2 })],
        }),
      );
    }
    if (content && content.length) inner.push(...content);
    else inner.push(makeTinyParagraph());
    return new TableCell({
      width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
      borders: {
        ...NO_BORDERS,
        bottom: withDashedBottom
          ? { style: BorderStyle.DASHED, size: 6, color: '888888' }
          : NO_BORDER,
      },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: inner,
    });
  };

  return new Table({
    width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [PAGE_CONTENT_WIDTH],
    borders: TABLE_NO_BORDERS,
    rows: [
      new TableRow({
        cantSplit: true,
        height: { value: HALF_PAGE_HEIGHT, rule: HeightRule.EXACT },
        children: [makeHalfCell(first, second !== null, needPageBreak)],
      }),
      new TableRow({
        cantSplit: true,
        height: { value: HALF_PAGE_HEIGHT, rule: HeightRule.EXACT },
        children: [makeHalfCell(second, false)],
      }),
    ],
  });
}

/** 生成可编辑的 Word 工单 Blob */
export async function buildWorkOrderDocxBlob(
  columns: ExportColumn[],
  records: IRecord[],
  onProgress?: (done: number, total: number) => void,
): Promise<Blob> {
  const fieldIds = resolveFields(columns);
  const columnsById = new Map(columns.map((col) => [col.id, col]));
  const rowMaps = sortRows(
    records.map((record) => toRowMap(columns, record)),
    fieldIds.defectNo,
  );

  // 每张工单生成内容块，按两单一页分组为容器表格
  const blocksPerOrder = rowMaps.map((row, index) => {
    onProgress?.(index + 1, rowMaps.length);
    return buildOrderBlocks(buildOrderModel(row, fieldIds, columnsById));
  });

  const bodyChildren: (Paragraph | Table)[] = [];
  let pageIndex = 0;
  for (let i = 0; i < blocksPerOrder.length; i += 2) {
    // 第 2 页起靠容器内 pageBreakBefore 分页；页间 20twips 段落防止相邻表格合并
    if (pageIndex > 0) bodyChildren.push(makeTinyParagraph());
    bodyChildren.push(buildPageTable(blocksPerOrder[i], blocksPerOrder[i + 1] ?? null, pageIndex > 0));
    pageIndex += 1;
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          // 默认 2pt 仅作用于 Word 自动追加的空段落标记（业务文字均显式指定字号）
          run: { size: 4, font: FONT },
          paragraph: { spacing: { before: 0, after: 0, line: 240 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: PAGE_HEIGHT },
            margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: 720, right: 720 },
          },
          // 网格行距压到 1twips：避免空段落被默认 360 网格吸附拉高导致整页溢出
          grid: { type: DocumentGridType.LINES, linePitch: 1 },
        },
        children: bodyChildren,
      },
    ],
  });

  return Packer.toBlob(doc);
}
