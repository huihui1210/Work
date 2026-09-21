import type { IRecord } from '@lark-base-open/js-sdk';
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeightRule,
  Packer,
  Paragraph,
  SectionType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from 'docx';
import type { ExportColumn } from './bitable-helper';
import {
  CONSTRAINT_FIELDS,
  dateOnly,
  parseDiscoverDate,
  resolveFields,
  sortRows,
  toRowMap,
  yesNoText,
} from './work-order-pdf';
import type { RowMap } from './work-order-pdf';

/**
 * 可编辑的 Word 格式缺陷处理工单（与 PDF 版式一致）：
 * A4 一页上下两张工单；表格文本可直接编辑；姓名使用楷体手写风格。
 */

/* A4 内容区 6 列网格（twips）：标签 1500，三列标签共 4500，值均分剩余宽度 */
const PAGE_CONTENT_WIDTH = 10466;
const GRID = [1500, 1989, 1500, 1989, 1500, 1988];

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: '333333' };
const CELL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const BODY_SIZE = 18; // 9pt，与 PDF 工单 12px 视觉一致
const LABEL_SIZE = 18;

interface CellSpec {
  text?: string;
  width: number;
  span?: number;
  bold?: boolean;
  align?: (typeof AlignmentType)[keyof typeof AlignmentType];
  /** 楷体手写风格 */
  handwriting?: boolean;
  size?: number;
  minHeight?: number;
}

function makeCell(spec: CellSpec): TableCell {
  const runs = [
    new TextRun({
      text: spec.text || '',
      bold: spec.bold ?? false,
      size: spec.handwriting ? 26 : (spec.size ?? BODY_SIZE),
      font: spec.handwriting
        ? { ascii: '楷体', eastAsia: '楷体', hAnsi: '楷体' }
        : { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' },
    }),
  ];
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
        children: runs,
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

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};

/** 无边框空段落（用于表格之间防合并） */
function makeTinyParagraph(): Paragraph {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 200 },
    children: [new TextRun({ text: '', size: 2 })],
  });
}

/** 标签 + 值（跨 span 列） */
function labelValue(
  label: string,
  value: string,
  labelWidth: number,
  valueWidth: number,
  span?: number,
  minHeight?: number,
): [TableCell, TableCell] {
  return [
    makeCell({ text: label, width: labelWidth, bold: true, size: LABEL_SIZE }),
    makeCell({
      text: value,
      width: valueWidth,
      span,
      minHeight,
      align: AlignmentType.CENTER,
    }),
  ];
}

function joinNonEmpty(parts: string[], sep = '　'): string {
  return parts.filter(Boolean).join(sep);
}

/** 构建单张工单的内容块（标题 + 顶部信息 + 主表），由外层容器负责分页与纵向定位 */
function buildOrderBlocks(
  row: RowMap,
  fieldIds: Record<string, string | undefined>,
  columnsById: Map<string, ExportColumn>,
): (Paragraph | Table)[] {
  const get = (key: string) => (fieldIds[key] ? row[fieldIds[key]!] ?? '' : '');
  const defectNo = get('defectNo');
  const discoverDate = parseDiscoverDate(defectNo);
  const closeDate = dateOnly(get('closeTime'));

  const blocks: (Paragraph | Table)[] = [];

  // 抬头（公司 / 标题）
  blocks.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 20 },
      children: [
        new TextRun({
          text: '福建LNG接收站',
          bold: true,
          size: 24,
          font: { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' },
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [
        new TextRun({
          text: '设备缺陷处理工单',
          bold: true,
          size: 32,
          font: { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' },
        }),
      ],
    }),
  );

  // 顶部信息（发现时间 / 编号），无边框两列表格
  if (discoverDate || defectNo) {
    blocks.push(
      new Table({
        width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
        columnWidths: [5233, 5233],
        borders: {
          top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 5233, type: WidthType.DXA },
                borders: {
                  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [new TextRun({ text: `发现时间：${discoverDate}`, size: BODY_SIZE })],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 5233, type: WidthType.DXA },
                borders: {
                  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
                },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: `编号：${defectNo}`, size: BODY_SIZE })],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    );
  }

  // 防止相邻表格被 Word 合并：顶部信息表与主表之间加微间距
  if (discoverDate || defectNo) {
    blocks.push(makeTinyParagraph());
  }

  /* ------- 主表格 ------- */
  const [c0, c1, c2, c3, c4, c5] = GRID;
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
      300,
    ),
  );

  // 部门与专业
  const deptCell = labelValue('责任部门', get('dept'), c0, c1 + c2 + c3, 3, 300);
  const specCell = labelValue('专业', get('specialty'), c4, c5, undefined, 300);
  rows.push(makeRow([...deptCell, ...specCell]));

  // 缺陷描述（区域部位 / 设备名称 / 位号 合并一行）
  const descPartIds: string[] = JSON.parse(fieldIds.descPartIds ?? '[]');
  const parts = descPartIds.map((id) => row[id] ?? '').filter(Boolean);
  const descText = joinNonEmpty([parts.join(' / '), get('desc')]);
  rows.push(
    makeRow(
      [makeCell({ text: '缺陷描述', width: c0, bold: true }), makeCell({ text: descText, width: c1 + c2 + c3 + c4 + c5, span: 5, align: AlignmentType.CENTER })],
      420,
    ),
  );

  // 日期：计划期限 / 消项时间（只到日）
  if (fieldIds.deadline || fieldIds.closeTime) {
    const a = labelValue('计划期限', dateOnly(get('deadline')), c0, c1 + c2, 2, 300);
    const b = labelValue('消项时间', closeDate, c3, c4 + c5, 2, 300);
    rows.push(makeRow([a[0], a[1], b[0], b[1]]));
  }

  // 人员：党员责任人 / 负责人 / 处理人
  if (fieldIds.partyMember || fieldIds.manager || fieldIds.handler) {
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
        460,
      ),
    );
  }

  // 是否类字段（勾选形式）
  const yesNoIds: string[] = JSON.parse(fieldIds.yesNoIds ?? '[]');
  for (let i = 0; i < yesNoIds.length; i += 2) {
    const first = {
      name: columnsById.get(yesNoIds[i])?.name ?? '状态',
      value: yesNoText(row[yesNoIds[i]] ?? ''),
    };
    const second = yesNoIds[i + 1]
      ? {
        name: columnsById.get(yesNoIds[i + 1])?.name ?? '状态',
        value: yesNoText(row[yesNoIds[i + 1]] ?? ''),
      }
      : null;
    if (second) {
      const a = labelValue(first.name, first.value, c0, c1 + c2, 2, 300);
      const b = labelValue(second.name, second.value, c3, c4 + c5, 2, 300);
      rows.push(makeRow([a[0], a[1], b[0], b[1]]));
    } else {
      rows.push(makeRow(labelValue(first.name, first.value, c0, c1 + c2 + c3 + c4 + c5, 5, 300)));
    }
  }

  // 消缺处理情况
  const repairText = get('repairNote');
  if (repairText) {
    rows.push(
      makeRow(
        [
          makeCell({ text: '消缺处理情况', width: c0, bold: true }),
          makeCell({ text: repairText, width: c1 + c2 + c3 + c4 + c5, span: 5, align: AlignmentType.LEFT }),
        ],
        420,
      ),
    );
  }

  // 风险控制措施 + 备注
  if (fieldIds.risk || fieldIds.remark) {
    const a = labelValue('风险控制措施', get('risk'), c0, c1 + c2, 2, 300);
    const b = labelValue('备注', get('remark'), c3, c4 + c5, 2, 300);
    rows.push(makeRow([a[0], a[1], b[0], b[1]]));
  }

  // 制约因素（填写形式）
  const constraintItems = CONSTRAINT_FIELDS.filter((name) => fieldIds[`constraint_${name}`]);
  let constraintText = '';
  if (constraintItems.length) {
    constraintText = constraintItems
      .map((name) => {
        const v = get(`constraint_${name}`);
        return v ? `${name}：${v}` : `${name}：`;
      })
      .join('　　');
    rows.push(
      makeRow(
        [
          makeCell({ text: '制约因素', width: c0, bold: true }),
          makeCell({ text: constraintText, width: c1 + c2 + c3 + c4 + c5, span: 5, align: AlignmentType.LEFT }),
        ],
        460,
      ),
    );
  }

  // 其他信息
  const leftoverIds: string[] = JSON.parse(fieldIds.leftovers ?? '[]');
  for (const id of leftoverIds) {
    const value = (row[id] ?? '').trim();
    if (!value) continue;
    rows.push(
      makeRow(
        [
          makeCell({ text: columnsById.get(id)?.name ?? '其他', width: c0, bold: true }),
          makeCell({ text: value, width: c1 + c2 + c3 + c4 + c5, span: 5, align: AlignmentType.LEFT }),
        ],
        460,
      ),
    );
  }

  // 签字区：消缺人签字 / 班长签字（姓名手写体）+ 日期
  const signSpecs = [
    { role: '消缺人', name: get('repairer') },
    { role: '班长', name: get('leader') },
  ];
  rows.push(
    makeRow(
      signSpecs.map((item) => {
        const paraChildren: TextRun[] = [
          new TextRun({
            text: `${item.role}签字：`,
            bold: true,
            size: BODY_SIZE,
            font: { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' },
          }),
        ];
        if (item.name) {
          paraChildren.push(
            new TextRun({
              text: item.name,
              size: 22,
              font: { ascii: '楷体', eastAsia: '楷体', hAnsi: '楷体' },
            }),
          );
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
              children: paraChildren,
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0, line: 260 },
              children: [
                new TextRun({
                  text: `日期：${closeDate || '　　　年　　月　　日'}`,
                  size: BODY_SIZE,
                  font: { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' },
                }),
              ],
            }),
          ],
        });
      }),
      560,
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

/** A4 内容区与半页高度（twips），与页边距设置保持一致 */
const PAGE_MARGIN_Y = 567;
const PAGE_CONTENT_HEIGHT = 16838 - PAGE_MARGIN_Y * 2;
const HALF_PAGE_HEIGHT = Math.floor(PAGE_CONTENT_HEIGHT / 2);

/**
 * 一页容器：1 列 2 行的无边框表格，每行固定半页高，工单内容在半区内垂直居中。
 * 第二张工单存在时，上半区底边显示虚线分隔；不存在时下半区留空（工单仍位于上半区）。
 */
function buildPageTable(first: (Paragraph | Table)[], second: (Paragraph | Table)[] | null): Table {
  const halfCellBorders = (withDashedBottom: boolean) => ({
    ...NO_BORDERS,
    bottom: withDashedBottom
      ? { style: BorderStyle.DASHED, size: 6, color: '888888' }
      : NO_BORDER,
  });

  const makeHalfCell = (content: (Paragraph | Table)[] | null, withDashedBottom: boolean) => {
    // 空单元格至少要有一个空段落
    const inner = content && content.length ? content : [makeTinyParagraph()];
    return new TableCell({
      width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
      borders: halfCellBorders(withDashedBottom),
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: inner,
    });
  };

  return new Table({
    width: { size: PAGE_CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [PAGE_CONTENT_WIDTH],
    borders: {
      top: NO_BORDER,
      bottom: NO_BORDER,
      left: NO_BORDER,
      right: NO_BORDER,
      insideHorizontal: NO_BORDER,
      insideVertical: NO_BORDER,
    },
    rows: [
      new TableRow({
        height: { value: HALF_PAGE_HEIGHT, rule: HeightRule.EXACT },
        children: [makeHalfCell(first, second !== null)],
      }),
      new TableRow({
        height: { value: PAGE_CONTENT_HEIGHT - HALF_PAGE_HEIGHT, rule: HeightRule.EXACT },
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
  const rowMaps = sortRows(
    records.map((record) => toRowMap(columns, record)),
    fieldIds.defectNo,
  );
  const columnsById = new Map(columns.map((col) => [col.id, col]));

  // 每张工单生成内容块，按两单一页分组，每节一页（分节符天然分页，不留空行）
  const blocksPerOrder = rowMaps.map((row, index) => {
    onProgress?.(index + 1, rowMaps.length);
    return buildOrderBlocks(row, fieldIds, columnsById);
  });

  const pageTables: Table[] = [];
  for (let i = 0; i < blocksPerOrder.length; i += 2) {
    pageTables.push(buildPageTable(blocksPerOrder[i], blocksPerOrder[i + 1] ?? null));
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            size: BODY_SIZE,
            font: { ascii: '微软雅黑', eastAsia: '微软雅黑', hAnsi: '微软雅黑' },
          },
        },
      },
    },
    sections: pageTables.map((pageTable) => ({
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: PAGE_MARGIN_Y, bottom: PAGE_MARGIN_Y, left: 720, right: 720 },
        },
      },
      children: [pageTable],
    })),
  });

  return Packer.toBlob(doc);
}
