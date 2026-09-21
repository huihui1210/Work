import type { IRecord } from '@lark-base-open/js-sdk';
import {
  AlignmentType,
  BorderStyle,
  Document,
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

/** 当前工单的行高累加器（用于估算工单高度，把虚线定位到页面中线） */
let activeRowHeightSink: ((height: number) => void) | null = null;

function makeRow(cells: TableCell[], minHeight?: number): TableRow {
  activeRowHeightSink?.(minHeight ?? 300);
  return new TableRow({
    cantSplit: true,
    height: minHeight ? { value: minHeight, rule: HeightRule.ATLEAST } : undefined,
    children: cells,
  });
}

/** 固定高度的空白段落（用于页面纵向定位） */
function makeExactSpacer(height: number): Paragraph {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: Math.max(1, Math.round(height)), lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: '', size: 2 })],
  });
}

/** 估算长文本在指定宽度（twips）内换行产生的额外高度 */
function wrapExtra(text: string, widthTwips: number): number {
  if (!text) return 0;
  const charsPerLine = Math.max(1, Math.floor(widthTwips / 180));
  const lines = text
    .split('\n')
    .reduce((sum, seg) => sum + Math.max(1, Math.ceil([...seg].length / charsPerLine)), 0);
  return Math.max(0, lines - 1) * 260;
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

function buildOrderChildren(
  row: RowMap,
  fieldIds: Record<string, string | undefined>,
  columnsById: Map<string, ExportColumn>,
  orderIndex: number,
  totalCount: number,
): (Paragraph | Table)[] {
  const get = (key: string) => (fieldIds[key] ? row[fieldIds[key]!] ?? '' : '');
  const defectNo = get('defectNo');
  const discoverDate = parseDiscoverDate(defectNo);
  const closeDate = dateOnly(get('closeTime'));

  const children: (Paragraph | Table)[] = [];

  // 抬头（公司 / 标题）
  children.push(
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
    children.push(
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
    children.push(
      new Paragraph({
        spacing: { before: 10, after: 10 },
        children: [new TextRun({ text: '', size: 2 })],
      }),
    );
  }

  /* ------- 主表格 ------- */
  const [c0, c1, c2, c3, c4, c5] = GRID;
  const rows: TableRow[] = [];
  let tableH = 0;
  let wrapH = 0;
  activeRowHeightSink = (h) => {
    tableH += h;
  };

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
    wrapH += wrapExtra(constraintText, c1 + c2 + c3 + c4 + c5);
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
    wrapH += wrapExtra(value, c1 + c2 + c3 + c4 + c5);
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
  activeRowHeightSink = null;

  // 统计长文本换行带来的额外行高（描述 / 处理情况 / 风险 / 备注）
  const fullWidth = c1 + c2 + c3 + c4 + c5;
  wrapH += wrapExtra(descText, fullWidth);
  wrapH += wrapExtra(get('repairNote'), fullWidth);
  wrapH += wrapExtra(get('risk'), c1 + c2);
  wrapH += wrapExtra(get('remark'), c1 + c2);

  children.push(
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

  /* ------- 纵向定位：与 PDF 一致，虚线位于页面中线，上下留白对称 ------- */
  const PAGE_CONTENT_HEIGHT = 16838 - 567 * 2;
  const HALF_PAGE = PAGE_CONTENT_HEIGHT / 2;
  const TITLE_BLOCK_H = 700;
  const META_BLOCK_H = 320;
  // 签字行两行文字实际略高于最小行高
  const SIGN_EXTRA_H = 60;
  const orderHeight =
    TITLE_BLOCK_H +
    (discoverDate || defectNo ? META_BLOCK_H : 0) +
    tableH +
    wrapH +
    SIGN_EXTRA_H;

  const isFirstOfPage = orderIndex % 2 === 0;
  const hasPartner = orderIndex + 1 < totalCount;

  // 页首元素：奇数页起始先分页；单张工单独占页时在顶部补空白整体居中
  const head: Paragraph[] = [];
  if (isFirstOfPage && orderIndex > 0) {
    head.push(new Paragraph({ pageBreakBefore: true, children: [new TextRun({ text: '', size: 2 })] }));
  }
  if (isFirstOfPage && !hasPartner) {
    const topPad = Math.max(0, Math.round((PAGE_CONTENT_HEIGHT - orderHeight) / 2));
    head.push(makeExactSpacer(topPad));
  }
  if (head.length) children.unshift(...head);

  if (isFirstOfPage && hasPartner) {
    // 同页两张：第一张后补空白，把虚线推到页面正中
    const gap = Math.max(40, Math.round(HALF_PAGE - orderHeight - 80));
    children.push(makeExactSpacer(gap));
    children.push(
      new Paragraph({
        spacing: { before: 0, after: 0 },
        border: { bottom: { style: BorderStyle.DASHED, size: 6, color: '888888', space: 1 } },
        children: [new TextRun({ text: '', size: 2 })],
      }),
    );
    children.push(makeExactSpacer(80));
  }

  return children;
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

  const children: (Paragraph | Table)[] = [];
  rowMaps.forEach((row, index) => {
    children.push(...buildOrderChildren(row, fieldIds, columnsById, index, rowMaps.length));
    onProgress?.(index + 1, rowMaps.length);
  });

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
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 567, bottom: 567, left: 720, right: 720 },
          },
        },
        children,
      },
    ],
  });

  return Packer.toBlob(doc);
}
