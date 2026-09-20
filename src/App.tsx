import { useCallback, useEffect, useRef, useState } from 'react';
import { bitable } from '@lark-base-open/js-sdk';
import { ExportError, cellToText, getSelectionInfo, getSelectedData } from './bitable-helper';
import type { SelectionInfo } from './bitable-helper';
import { buildFileBlob, buildFilename, copyToClipboard, downloadBlob } from './exporter';
import type { ExportFormat } from './exporter';
import * as LZString from 'lz-string';
import './styles.css';

type AppFormat = ExportFormat | 'dms';

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
  { value: 'json', label: 'JSON', desc: '.json 数据' },
  { value: 'clipboard', label: '剪贴板', desc: '直接粘贴' },
  { value: 'dms', label: '缺陷系统', desc: '一键发送' },
];

type MessageType = 'success' | 'error' | 'info';

export default function App() {
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [checking, setChecking] = useState(true);
  const [notInHost, setNotInHost] = useState(false);
  const [format, setFormat] = useState<AppFormat>('xlsx');
  const [exporting, setExporting] = useState(false);
  const [message, setMessage] = useState<{ type: MessageType; text: string } | null>(null);

  const refreshingRef = useRef(false);

  const refresh = useCallback(async () => {
    if (refreshingRef.current) return;
    refreshingRef.current = true;
    try {
      const info = await getSelectionInfo();
      setSelection(info);
      setNotInHost(false);
    } catch (error) {
      setNotInHost(window.top === window || error instanceof ExportError);
    } finally {
      refreshingRef.current = false;
      setChecking(false);
    }
  }, []);

  useEffect(() => {
    void refresh();
    let unsubscribe: () => void = () => undefined;
    try {
      unsubscribe = bitable.base.onSelectionChange(() => void refresh());
    } catch {
      // 非多维表格环境时忽略
    }
    return unsubscribe;
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
          const blob = buildFileBlob('xlsx', data.columns, data.rows, data.viewName);
          downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'xlsx'));
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
          const blob = buildFileBlob('xlsx', data.columns, data.rows, data.viewName);
          downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'xlsx'));
          setMessage({
            type: 'error',
            text: '无法自动打开缺陷系统，已改为下载 Excel，请手动导入。',
          });
        }
        return;
      }

      if (format === 'clipboard') {
        await copyToClipboard(data.columns, data.rows);
        setMessage({ type: 'success', text: `已复制 ${data.rows.length} 条记录，可直接粘贴到 Excel。` });
        return;
      }

      const ext = format;
      const blob = buildFileBlob(format, data.columns, data.rows, data.viewName);
      downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, ext));
      setMessage({ type: 'success', text: `已导出 ${data.rows.length} 条记录（${data.columns.length} 列）。` });
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
          <p className="subtitle">勾选记录，导出为 Excel / CSV / JSON</p>
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
      </section>

      <button type="button" className="primary-btn" disabled={exporting} onClick={() => void handleExport()}>
        {exporting
          ? '处理中…'
          : format === 'clipboard'
            ? '一键复制到剪贴板'
            : format === 'dms'
              ? '一键发送到缺陷系统'
              : '一键导出'}
      </button>

      {message && <div className={`banner banner-${message.type}`}>{message.text}</div>}

      <p className="tips">
        使用方法：在「表格」视图中勾选记录 → 选择格式 → 点击导出。仅导出当前视图可见的列，数据不会离开当前页面。
      </p>
    </div>
  );
}

function errorToMessage(error: unknown): string {
  if (error instanceof ExportError && error.code === 'NO_CONTEXT') {
    return '请先在多维表格中打开一个数据表和视图。';
  }
  return '导出失败，请确认插件运行在多维表格边栏中，且你对这些记录有查看权限。';
}
