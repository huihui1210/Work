import { useCallback, useEffect, useRef, useState } from 'react';
import { bitable, FieldType } from '@lark-base-open/js-sdk';
import {
  ExportError,
  cellToText,
  getAttachmentMap,
  getSelectionInfo,
  getSelectedData,
} from './bitable-helper';
import type { SelectionInfo } from './bitable-helper';
import { buildFileBlob, buildFilename, downloadBlob, exportXlsxWithImages } from './exporter';
import type { ExportFormat } from './exporter';
import * as LZString from 'lz-string';
import { buildWorkOrderPdfBlob } from './work-order-pdf';
import { buildWorkOrderDocxBlob } from './work-order-docx';
import './styles.css';

type AppFormat = ExportFormat | 'dms' | 'pdf';

interface FormatOption {
  value: AppFormat;
  label: string;
  desc: string;
}

const DMS_URL = 'https://huihui1210.github.io/DMS_DA/';
/** 压缩后锚点的安全长度上限，超出则退回 Excel 下载 */
const DMS_MAX_PAYLOAD_CHARS = 1_500_000;

const FORMAT_OPTIONS: FormatOption[] = [
  { value: 'xlsx', label: 'Excel', desc: '.xlsx 推荐' },
  { value: 'csv', label: 'CSV', desc: '.csv 通用' },
  { value: 'pdf', label: '工单PDF', desc: '缺陷处理工单' },
  { value: 'dms', label: '缺陷分析系统', desc: '一键发送' },
];

type MessageType = 'success' | 'error' | 'info';

export default function App() {
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [checking, setChecking] = useState(true);
  const [notInHost, setNotInHost] = useState(false);
  const [format, setFormat] = useState<AppFormat>('xlsx');
  const [withImages, setWithImages] = useState(false);
  const [withWord, setWithWord] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [message, setMessage] = useState<{ type: MessageType; text: string } | null>(null);

  const refreshingRef = useRef(false);
  const pendingRefreshRef = useRef(false);

  /** silent=true 时为后台轮询：不切换 loading/警告状态，仅在结果变化时更新界面 */
  const refresh = useCallback(async (silent = false) => {
    // 刷新进行中不丢弃请求，标记后排队补刷一次
    if (refreshingRef.current) {
      pendingRefreshRef.current = true;
      return;
    }
    refreshingRef.current = true;
    try {
      const info = await getSelectionInfo();
      setSelection((prev) => (isSameSelection(prev, info) ? prev : info));
      if (!silent) setNotInHost(false);
    } catch (error) {
      if (!silent) setNotInHost(window.top === window || error instanceof ExportError);
    } finally {
      refreshingRef.current = false;
      if (!silent) setChecking(false);
      if (pendingRefreshRef.current) {
        pendingRefreshRef.current = false;
        void refresh();
      }
    }
  }, []);

  useEffect(() => {
    void refresh();
    let unsubscribe: () => void = () => undefined;
    try {
      // 事件能触发时立即响应；若宿主对勾选不抛事件，由下面的定时轮询兜底
      unsubscribe = bitable.base.onSelectionChange(() => void refresh());
    } catch {
      // 非多维表格环境时忽略
    }
    // 定时静默轮询：勾选事件在部分宿主场景下不触发，轮询保证数量在 1 秒内自动同步
    const pollTimer = window.setInterval(() => void refresh(true), 800);
    return () => {
      unsubscribe();
      window.clearInterval(pollTimer);
    };
  }, [refresh]);

  const handleExport = async () => {
    setExporting(true);
    setMessage(null);
    // 发送到缺陷系统时，必须在任何 await 之前同步打开窗口，否则会被弹窗拦截
    let dmsWindow: Window | null = null;
    if (format === 'dms') {
      dmsWindow = window.open(DMS_URL, '_blank');
    }
    try {
      const data = await getSelectedData();
      if (data.rows.length === 0) {
        setMessage({ type: 'info', text: '请先在表格视图中勾选要导出的记录。' });
        return;
      }

      if (format === 'dms') {
        if (!dmsWindow) {
          setMessage({
            type: 'error',
            text: '浏览器拦截了弹窗。请允许本页面弹出窗口后重试，或改用 Excel 导出后手动导入。',
          });
          return;
        }
        const payloadRows = data.rows.map((record) => {
          const obj: Record<string, string> = {};
          for (const column of data.columns) {
            obj[column.name] = cellToText(record.fields[column.id] ?? null, column.type);
          }
          return obj;
        });
        const compressed = LZString.compressToEncodedURIComponent(JSON.stringify(payloadRows));

        if (compressed.length > DMS_MAX_PAYLOAD_CHARS) {
          await downloadSelectedAsXlsx(data);
          setMessage({
            type: 'info',
            text: `勾选数据较多（${data.rows.length} 条），已自动改为下载 Excel，请在缺陷系统中手动导入该文件。`,
          });
          return;
        }

        try {
          // 带数据锚点重新导航，DMS 加载时自动读取导入（沙箱无法拦截页面导航）
          dmsWindow.location.href = `${DMS_URL}?t=${Date.now()}#dms=${compressed}`;
          setMessage({
            type: 'success',
            text: `已发送 ${data.rows.length} 条记录到缺陷管理系统，请查看新标签页。`,
          });
          void refresh();
        } catch {
          await downloadSelectedAsXlsx(data);
          setMessage({
            type: 'error',
            text: '无法自动打开缺陷系统，已改为下载 Excel，请手动导入。',
          });
        }
        return;
      }

      if (format === 'pdf') {
        // 工单不包含缺陷图片：过滤附件列
        const columns = data.columns.filter((col) => col.type !== FieldType.Attachment);
        if (withWord) {
          setMessage({ type: 'info', text: '正在生成工单 Word…' });
          const blob = await buildWorkOrderDocxBlob(columns, data.rows, (done, total) => {
            setMessage({ type: 'info', text: `正在生成工单 Word（${done}/${total} 张）…` });
          });
          downloadBlob(blob, buildFilename(`缺陷处理工单_${data.tableName}`, 'docx'));
          setMessage({
            type: 'success',
            text: `已生成 ${data.rows.length} 张可编辑工单 Word（按缺陷编号排序）。`,
          });
        } else {
          setMessage({ type: 'info', text: '正在生成工单 PDF…' });
          const blob = await buildWorkOrderPdfBlob(columns, data.rows, (done, total) => {
            setMessage({ type: 'info', text: `正在生成工单 PDF（${done}/${total} 页）…` });
          });
          downloadBlob(blob, buildFilename(`缺陷处理工单_${data.tableName}`, 'pdf'));
          setMessage({
            type: 'success',
            text: `已生成 ${data.rows.length} 张缺陷处理工单 PDF（按缺陷编号排序）。`,
          });
        }
        void refresh();
        return;
      }

      if (format === 'xlsx') {
        if (withImages) {
          setMessage({ type: 'info', text: '正在读取缺陷图片…' });
          const attachmentMap = await getAttachmentMap(data.columns, data.rows);
          const blob = await exportXlsxWithImages(
            data.columns,
            data.rows,
            attachmentMap,
            data.viewName,
            (done, total) =>
              setMessage(
                total
                  ? { type: 'info', text: `正在下载缺陷图片（${done}/${total}）…` }
                  : { type: 'info', text: '正在生成 Excel…' },
              ),
          );
          downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'xlsx'));
          setMessage({
            type: 'success',
            text: `已导出 ${data.rows.length} 条记录（${data.columns.length} 列），缺陷图片已内嵌。`,
          });
        } else {
          // 未勾选图片：不导出附件列（其余字段全部保留，即使选中记录中该列暂无数据）
          const columns = data.columns.filter((col) => col.type !== FieldType.Attachment);
          const blob = await buildFileBlob('xlsx', columns, data.rows, data.viewName);
          downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'xlsx'));
          setMessage({
            type: 'success',
            text: `已导出 ${data.rows.length} 条记录（${columns.length} 列，未含附件列）。`,
          });
        }
        void refresh();
        return;
      }

      // CSV：不导出附件列（其余字段全部保留，即使暂无数据）；计划期限只到日
      const columns = data.columns.filter((col) => col.type !== FieldType.Attachment);
      const blob = await buildFileBlob('csv', columns, data.rows, data.viewName);
      downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'csv'));
      setMessage({
        type: 'success',
        text: `已导出 ${data.rows.length} 条记录（${columns.length} 列，未含附件列）。`,
      });
      void refresh();
    } catch (error) {
      setMessage({ type: 'error', text: errorToMessage(error) });
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="app">
      <header className="header">
        <div className="header-icon">⇩</div>
        <div>
          <h1>一键导出选中内容</h1>
          <p className="subtitle">勾选记录，导出为 Excel / CSV / 工单PDF</p>
        </div>
      </header>

      {notInHost && (
        <div className="banner banner-warning">
          当前页面未运行在多维表格边栏中。请在多维表格里通过「插件 → 自定义插件」添加本服务地址，不要直接用浏览器打开。
        </div>
      )}

      <section className="card">
        <div className="card-top">
          <span className="card-label">当前选中</span>
          <button type="button" className="text-btn" onClick={() => void refresh()}>
            刷新
          </button>
        </div>
        <div className="selection-main">
          <span className="count-badge">{checking ? '…' : selection?.count ?? 0}</span>
          <span className="count-unit">条记录</span>
        </div>
        <p className="location">
          {selection ? `${selection.tableName} / ${selection.viewName}` : '未检测到数据表'}
        </p>
        {selection && !selection.multiSelectSupported && (
          <p className="hint">当前视图不支持读取多选，切换到「表格」视图可勾选多条记录。</p>
        )}
      </section>

      <section className="card">
        <span className="card-label">导出格式</span>
        <div className="format-grid">
          {FORMAT_OPTIONS.map((option) => (
            <button
              type="button"
              key={option.value}
              className={`format-item ${format === option.value ? 'active' : ''}`}
              onClick={() => setFormat(option.value)}
            >
              <span className="format-name">{option.label}</span>
              <span className="format-desc">{option.desc}</span>
            </button>
          ))}
        </div>
        {format === 'xlsx' && (
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={withImages}
              onChange={(e) => setWithImages(e.target.checked)}
            />
            <span className="checkbox-text">导出图片</span>
            <span className="checkbox-hint">不勾选则不导出缺陷图片</span>
          </label>
        )}
        {format === 'pdf' && (
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={withWord}
              onChange={(e) => setWithWord(e.target.checked)}
            />
            <span className="checkbox-text">导出Word</span>
            <span className="checkbox-hint">导出可编辑的 Word 格式工单</span>
          </label>
        )}
      </section>

      <button type="button" className="primary-btn" disabled={exporting} onClick={() => void handleExport()}>
        {exporting
          ? '处理中…'
          : format === 'dms'
            ? '一键发送到缺陷分析系统'
            : format === 'pdf'
              ? withWord
                ? '一键导出工单Word'
                : '一键导出工单PDF'
              : '一键导出'}
      </button>

      {message && <div className={`banner banner-${message.type}`}>{message.text}</div>}

      <p className="tips">
        使用方法：在「表格」视图中勾选记录 → 选择格式 → 点击导出。Excel 可勾选内嵌缺陷图片；
        工单PDF 可勾选导出可编辑的 Word 版工单。仅导出当前视图可见的列，数据不会离开当前页面。
      </p>
    </div>
  );
}

/** DMS 发送失败或数据过大时的兜底：下载美化 Excel 供手动导入 */
async function downloadSelectedAsXlsx(
  data: Awaited<ReturnType<typeof getSelectedData>>,
): Promise<void> {
  const blob = await buildFileBlob('xlsx', data.columns, data.rows, data.viewName);
  downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'xlsx'));
}

/** 判断两次读取的选中状态是否一致，避免轮询造成无谓的界面刷新 */
function isSameSelection(a: SelectionInfo | null, b: SelectionInfo): boolean {
  return (
    !!a &&
    a.count === b.count &&
    a.tableName === b.tableName &&
    a.viewName === b.viewName &&
    a.multiSelectSupported === b.multiSelectSupported
  );
}

function errorToMessage(error: unknown): string {
  if (error instanceof ExportError && error.code === 'NO_CONTEXT') {
    return '请先在多维表格中打开一个数据表和视图。';
  }
  console.error('[一键导出] 导出失败，原始错误：', error);
  const detail = error instanceof Error ? error.message : String(error ?? '');
  return detail
    ? `导出失败：${detail}`
    : '导出失败，请确认插件运行在多维表格边栏中，且你对这些记录有查看权限。';
}
