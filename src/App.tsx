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
/** localStorage 自动同步数据键：Work插件写、DMS 轮询读 */
const DMS_SYNC_PAYLOAD_KEY = 'dms_sync_payload';
/** localStorage 自动同步版本号键：每次同步 +1，DMS 靠版本号变化触发 */
const DMS_SYNC_VER_KEY = 'dms_sync_ver';
/** 设备缺陷分析助手分享链接（内网地址） */
const AGENT_URL =
  'https://agents.cnooc.com.cn/share/eOzEl9oq6nZb80Yzp8yy5Bc1VhRnOdjd6exWvQ1NCGSGIyXP_aM7uA-sTCMJbK92LCrMEfUf-9PmQJ8mkLf5fQ';

const FORMAT_OPTIONS: FormatOption[] = [
  { value: 'xlsx', label: 'Excel', desc: '推荐' },
  { value: 'csv', label: 'CSV', desc: '通用' },
  { value: 'pdf', label: '工单PDF', desc: '可打印' },
  { value: 'dms', label: '缺陷分析系统', desc: '发送数据' },
];

type MessageType = 'success' | 'error' | 'info';

/** 自动同步间隔选项（分钟） */
const AUTO_INTERVAL_OPTIONS = [1, 3, 5, 10, 30];
/** localStorage 持久化键：自动同步设置 */
const AUTO_SYNC_STORAGE_KEY = 'dms-auto-sync-config';
/** localStorage 持久化键：上次同步时间（跨会话保留） */
const AUTO_SYNC_LAST_KEY = 'dms-auto-sync-last';

interface AutoSyncConfig {
  enabled: boolean;
  intervalMinutes: number;
}

export default function App() {
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [checking, setChecking] = useState(true);
  const [notInHost, setNotInHost] = useState(false);
  const [format, setFormat] = useState<AppFormat>('xlsx');
  const [withImages, setWithImages] = useState(false);
  const [withWord, setWithWord] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [message, setMessage] = useState<{ type: MessageType; text: string } | null>(null);
  /** 提示条自动消失：成功/信息 3 秒，错误 6 秒 */
  useEffect(() => {
    if (!message) return;
    const timer = window.setTimeout(
      () => setMessage(null),
      message.type === 'error' ? 6000 : 3000,
    );
    return () => window.clearTimeout(timer);
  }, [message]);
  /** 顶部页签：导出 / 智能助手 */
  const [tab, setTab] = useState<'export' | 'assistant'>('export');
  /** 首次进入助手页后才挂载 iframe，避免不必要地请求内网地址 */
  const [assistantMounted, setAssistantMounted] = useState(false);
  const [assistantKey, setAssistantKey] = useState(0);
  /** 复制数据到剪贴板进行中 */
  const [copying, setCopying] = useState(false);
  /** 浏览器拦截剪贴板时的降级文本 */
  const [copyFallback, setCopyFallback] = useState<string | null>(null);

  const refreshingRef = useRef(false);
  const pendingRefreshRef = useRef(false);

  // ======= 自动同步状态 =======
  const loadAutoConfig = (): AutoSyncConfig => {
    try {
      const raw = localStorage.getItem(AUTO_SYNC_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        return {
          enabled: false,
          intervalMinutes:
            typeof parsed.intervalMinutes === 'number' &&
              AUTO_INTERVAL_OPTIONS.includes(parsed.intervalMinutes)
              ? parsed.intervalMinutes
              : 5,
        };
      }
    } catch { }
    return { enabled: false, intervalMinutes: 5 };
  };
  const [autoConfig, setAutoConfig] = useState<AutoSyncConfig>(loadAutoConfig);
  const [nextSyncAt, setNextSyncAt] = useState<number | null>(null);
  const [lastSyncAt, setLastSyncAt] = useState<number | null>(() => {
    try {
      const raw = localStorage.getItem(AUTO_SYNC_LAST_KEY);
      const ts = raw ? Number(raw) : NaN;
      return Number.isFinite(ts) && ts > 0 ? ts : null;
    } catch {
      return null;
    }
  });
  /** 每秒刷新的当前时间，仅用于驱动倒计时显示 */
  const [autoNow, setAutoNow] = useState(Date.now());
  const autoTimerRef = useRef<number | null>(null);
  const autoCountdownRef = useRef<number | null>(null);
  const dmsWindowRef = useRef<Window | null>(null);

  /** 记录上次同步时间并持久化，重开插件后仍可显示 */
  const markSynced = () => {
    const ts = Date.now();
    setLastSyncAt(ts);
    try {
      localStorage.setItem(AUTO_SYNC_LAST_KEY, String(ts));
    } catch { }
  };

  /** 持久化自动同步间隔（enabled 不持久化，每次打开插件默认关闭） */
  useEffect(() => {
    localStorage.setItem(
      AUTO_SYNC_STORAGE_KEY,
      JSON.stringify({ intervalMinutes: autoConfig.intervalMinutes }),
    );
  }, [autoConfig.intervalMinutes]);

  /** 格式化剩余时间（ms → "X分Y秒"） */
  const formatCountdown = (ms: number): string => {
    if (ms < 0) return '同步中…';
    if (ms === 0) return '即将同步';
    const totalSec = Math.ceil(ms / 1000);
    const min = Math.floor(totalSec / 60);
    const sec = totalSec % 60;
    return min > 0 ? `${min}分${sec}秒` : `${sec}秒`;
  };

  /** 自动同步的主循环：读选中 → 压缩 → 写 localStorage 给 DMS
   *  同属 github.io origin，localStorage 天然共享，DMS 端轮询版本号就会收到。
   *  优点：完全不改 DMS URL、不触发 hashchange、不跳页面、不依赖 postMessage。
   */
  const runAutoSync = useCallback(async () => {
    try {
      const data = await getSelectedData();
      if (data.rows.length === 0) {
        setMessage({
          type: 'info',
          text: `[自动同步] 当前无选中记录（${data.tableName} / ${data.viewName}），跳过本次。`,
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
        setMessage({
          type: 'error',
          text: `[自动同步] 数据量过大（${compressed.length} chars），建议拆分视图或改用手动发送。`,
        });
        return;
      }

      // ✅ 写 localStorage：DMS 端轮询版本号变化就会读到
      try {
        const ver = String(Date.now());
        localStorage.setItem(DMS_SYNC_PAYLOAD_KEY, compressed);
        localStorage.setItem(DMS_SYNC_VER_KEY, ver);
        console.log('[Work插件→DMS] localStorage 已写入，ver=' + ver + '，长度=' + compressed.length);
      } catch (e) {
        console.warn('[Work插件→DMS] localStorage 写入失败，降级到 URL hash:', e);
        // 降级：新开窗口导航过去
        const targetWindow = window.open(DMS_URL, 'dms-da-main');
        if (targetWindow) {
          targetWindow.location.href = `${DMS_URL}?t=${Date.now()}#dms=${compressed}`;
          dmsWindowRef.current = targetWindow;
        }
      }

      markSynced();
      setMessage({
        type: 'success',
        text: `[自动同步] 已发送 ${data.rows.length} 条记录（${data.tableName} / ${data.viewName}）。`,
      });
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error ?? '');
      setMessage({ type: 'error', text: `[自动同步] 失败：${detail}` });
    }
  }, []);

  /** 启动/停止自动同步定时器（首次同步由勾选动作触发，这里只负责后续定时） */
  useEffect(() => {
    // 清理上一次
    if (autoTimerRef.current !== null) {
      window.clearTimeout(autoTimerRef.current);
      autoTimerRef.current = null;
    }
    if (autoCountdownRef.current !== null) {
      window.clearInterval(autoCountdownRef.current);
      autoCountdownRef.current = null;
    }

    if (!autoConfig.enabled || format !== 'dms') {
      setNextSyncAt(null);
      return;
    }

    let cancelled = false;
    const intervalMs = autoConfig.intervalMinutes * 60 * 1000;

    // 倒计时：目标时间固定，每秒只刷新“当前时间”
    autoCountdownRef.current = window.setInterval(() => setAutoNow(Date.now()), 1000);

    const scheduleNext = () => {
      if (cancelled) return;
      setAutoNow(Date.now());
      setNextSyncAt(Date.now() + intervalMs);
      autoTimerRef.current = window.setTimeout(async () => {
        if (cancelled) return;
        await runAutoSync();
        if (!cancelled) scheduleNext();
      }, intervalMs);
    };

    scheduleNext();

    return () => {
      cancelled = true;
      if (autoTimerRef.current !== null) {
        window.clearTimeout(autoTimerRef.current);
        autoTimerRef.current = null;
      }
      if (autoCountdownRef.current !== null) {
        window.clearInterval(autoCountdownRef.current);
        autoCountdownRef.current = null;
      }
    };
  }, [autoConfig, format, runAutoSync]);

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
      // 如果已有未关闭的 DMS 窗口，复用它，避免多开标签
      if (dmsWindowRef.current && !dmsWindowRef.current.closed) {
        dmsWindow = dmsWindowRef.current;
      } else {
        dmsWindow = window.open(DMS_URL, 'dms-da-main');
        if (dmsWindow) {
          dmsWindowRef.current = dmsWindow;
        }
      }
    }
    try {
      const data = await getSelectedData();
      const skipNote =
        data.failedRecordCount > 0
          ? `（另有 ${data.failedRecordCount} 条记录读取失败已跳过，可能包含异常单元格）`
          : '';
      if (data.rows.length === 0) {
        setMessage({
          type: data.failedRecordCount > 0 ? 'error' : 'info',
          text:
            data.failedRecordCount > 0
              ? `选中的 ${data.failedRecordCount} 条记录读取失败，可能包含异常单元格（如错误公式），已无法导出。`
              : '请先勾选记录。',
        });
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
          markSynced();
          setMessage({
            type: 'success',
            text: `已发送 ${data.rows.length} 条记录到缺陷管理系统。${skipNote}${autoConfig.enabled ? '（自动同步已启用）' : ''}`,
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
            text: `已生成 ${data.rows.length} 张可编辑工单 Word（按缺陷编号排序）。${skipNote}`,
          });
        } else {
          setMessage({ type: 'info', text: '正在生成工单 PDF…' });
          const blob = await buildWorkOrderPdfBlob(columns, data.rows, (done, total) => {
            setMessage({ type: 'info', text: `正在生成工单 PDF（${done}/${total} 页）…` });
          });
          downloadBlob(blob, buildFilename(`缺陷处理工单_${data.tableName}`, 'pdf'));
          setMessage({
            type: 'success',
            text: `已生成 ${data.rows.length} 张缺陷处理工单 PDF（按缺陷编号排序）。${skipNote}`,
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
            text: `已导出 ${data.rows.length} 条记录（${data.columns.length} 列），缺陷图片已内嵌。${skipNote}`,
          });
        } else {
          // 未勾选图片：不导出附件列（其余字段全部保留，即使选中记录中该列暂无数据）
          const columns = data.columns.filter((col) => col.type !== FieldType.Attachment);
          const blob = await buildFileBlob('xlsx', columns, data.rows, data.viewName);
          downloadBlob(blob, buildFilename(`${data.tableName}_${data.viewName}`, 'xlsx'));
          setMessage({
            type: 'success',
            text: `已导出 ${data.rows.length} 条记录（${columns.length} 列，未含附件列）。${skipNote}`,
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
        text: `已导出 ${data.rows.length} 条记录（${columns.length} 列，未含附件列）。${skipNote}`,
      });
      void refresh();
    } catch (error) {
      setMessage({ type: 'error', text: errorToMessage(error) });
    } finally {
      setExporting(false);
    }
  };

  /** 切到助手页；首次进入时挂载 iframe，之后保留对话状态 */
  const openAssistant = () => {
    setAssistantMounted(true);
    setTab('assistant');
  };

  /** 写入剪贴板；失败时返回 false，由界面弹窗兜底 */
  const writeClipboard = async (text: string): Promise<boolean> => {
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch {
      // 继续尝试降级方案
    }
    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      const ok = document.execCommand('copy');
      document.body.removeChild(textarea);
      return ok;
    } catch {
      return false;
    }
  };

  /** 把选中记录复制成文本（CSV 或列式 JSON），粘贴到对话框即可提问 */
  const handleCopyData = async (kind: 'csv' | 'json') => {
    if (copying) return;
    setCopying(true);
    setMessage(null);
    try {
      const data = await getSelectedData();
      if (data.rows.length === 0) {
        setMessage({ type: 'info', text: '请先勾选记录。' });
        return;
      }
      const columns = data.columns.filter((col) => col.type !== FieldType.Attachment);
      let body: string;
      if (kind === 'csv') {
        const blob = await buildFileBlob('csv', columns, data.rows, data.viewName);
        body = (await blob.text()).replace(/^\uFEFF/, '');
      } else {
        body = JSON.stringify({
          columns: columns.map((col) => col.name),
          rows: data.rows.map((record) =>
            columns.map((col) => cellToText(record.fields[col.id] ?? null, col.type)),
          ),
        });
      }
      const payload = [
        '请根据以下表格数据回答我的问题。',
        `表名：${data.tableName}；视图：${data.viewName}；共 ${data.rows.length} 条记录，${columns.length} 列。`,
        '',
        body,
        '',
        '我的问题：',
      ].join('\n');
      const ok = await writeClipboard(payload);
      if (!ok) {
        setCopyFallback(payload);
        setMessage({ type: 'info', text: '无法自动写入剪贴板，请在弹窗中按 Ctrl+C 复制。' });
        return;
      }
      const skipNote =
        data.failedRecordCount > 0 ? `（另有 ${data.failedRecordCount} 条读取失败已跳过）` : '';
      const sizeNote =
        payload.length > 60000 ? ' 内容较长，超限时请缩小勾选范围。' : '';
      setMessage({
        type: 'success',
        text: `已复制 ${data.rows.length} 条（约 ${payload.length.toLocaleString()} 字符），在对话框按 Ctrl+V 粘贴。${skipNote}${sizeNote}`,
      });
    } catch (error) {
      setMessage({ type: 'error', text: errorToMessage(error) });
    } finally {
      setCopying(false);
    }
  };

  const renderMessage = () =>
    message ? (
      <div className={`banner banner-${message.type}`} role="status">
        <span className="banner-text">{message.text}</span>
        <button
          type="button"
          className="banner-close"
          aria-label="关闭提示"
          onClick={() => setMessage(null)}
        >
          ×
        </button>
      </div>
    ) : null;

  return (
    <div className={`app${tab === 'assistant' ? ' app-assistant' : ''}`}>
      <div className="tabs">
        <button
          type="button"
          className={`tab${tab === 'export' ? ' active' : ''}`}
          onClick={() => setTab('export')}
        >
          导出
        </button>
        <button
          type="button"
          className={`tab${tab === 'assistant' ? ' active' : ''}`}
          onClick={openAssistant}
        >
          智能助手
        </button>
      </div>

      <div className="export-body" style={{ display: tab === 'export' ? 'flex' : 'none' }}>
      <header className="header">
        <div className="header-icon">⇩</div>
        <div>
          <h1>导出选中记录</h1>
        </div>
      </header>

      {notInHost && (
        <div className="banner banner-warning">
          未在多维表格中运行。请通过「插件 → 自定义插件」加载本页面。
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
          <p className="hint">当前视图不支持多选，请切换到「表格」视图。</p>
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
          </label>
        )}
        {format === 'pdf' && (
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={withWord}
              onChange={(e) => setWithWord(e.target.checked)}
            />
            <span className="checkbox-text">导出 Word</span>
          </label>
        )}

        {/* ====== 自动同步（仅 DMS 模式）====== */}
        {format === 'dms' && (
          <>
            <div className="divider" />
            <label className="checkbox-row">
              <input
                type="checkbox"
                checked={autoConfig.enabled}
                onChange={(e) => {
                  if (e.target.checked) {
                    // 在用户点击手势内同步打开窗口，避免被浏览器拦截；勾选后立即发送一次
                    if (!dmsWindowRef.current || dmsWindowRef.current.closed) {
                      const win = window.open(DMS_URL, 'dms-da-main');
                      if (!win) {
                        setMessage({
                          type: 'error',
                          text: '浏览器拦截了 DMS 弹窗，请允许本页面弹出窗口后重试。',
                        });
                        return;
                      }
                      dmsWindowRef.current = win;
                    }
                    setAutoConfig((c) => ({ ...c, enabled: true }));
                    void runAutoSync();
                  } else {
                    setAutoConfig((c) => ({ ...c, enabled: false }));
                    setMessage({ type: 'info', text: '自动同步已关闭。' });
                  }
                }}
              />
              <span className="checkbox-text">自动同步</span>
            </label>

            {autoConfig.enabled && (
              <div className="auto-sync-panel">
                <div className="auto-sync-row">
                  <span className="auto-sync-label">同步间隔</span>
                  <select
                    className="interval-select"
                    value={autoConfig.intervalMinutes}
                    onChange={(e) =>
                      setAutoConfig((c) => ({ ...c, intervalMinutes: Number(e.target.value) }))
                    }
                  >
                    {AUTO_INTERVAL_OPTIONS.map((min) => (
                      <option key={min} value={min}>
                        {min} 分钟
                      </option>
                    ))}
                  </select>
                  <span className="auto-sync-last">
                    上次同步：
                    {lastSyncAt ? new Date(lastSyncAt).toLocaleTimeString('zh-CN') : '—'}
                  </span>
                </div>
                <div className="auto-sync-status">
                  {dmsWindowRef.current && !dmsWindowRef.current.closed ? (
                    <>
                      <span className="status-dot status-active" />
                      <span className="status-text">接收窗口已就绪</span>
                    </>
                  ) : (
                    <>
                      <span className="status-dot status-warn" />
                      <span className="status-text">接收窗口未打开</span>
                    </>
                  )}
                  {nextSyncAt && (
                    <span className="status-countdown">
                      · 下次同步 {formatCountdown(nextSyncAt - autoNow)}
                    </span>
                  )}
                </div>
              </div>
            )}
          </>
        )}
      </section>

      {!(format === 'dms' && autoConfig.enabled) && (
        <button type="button" className="primary-btn" disabled={exporting} onClick={() => void handleExport()}>
          {exporting
            ? '处理中…'
            : format === 'dms'
              ? '发送到缺陷分析系统'
              : format === 'pdf'
                ? withWord
                  ? '导出工单 Word'
                  : '导出工单 PDF'
                : '导出'}
        </button>
      )}

      {renderMessage()}

      <div style={{ fontSize: '11px', color: '#bbb', marginTop: '4px' }}>v2.2.0</div>
      </div>

      {assistantMounted && (
        <div className="assistant-body" style={{ display: tab === 'assistant' ? 'flex' : 'none' }}>
          <div className="assistant-toolbar">
            <span className="assistant-title">缺陷分析助手</span>
            <button
              type="button"
              className="text-btn"
              disabled={copying}
              onClick={() => void handleCopyData('csv')}
            >
              {copying ? '复制中…' : '复制CSV'}
            </button>
            <button
              type="button"
              className="text-btn"
              disabled={copying}
              onClick={() => void handleCopyData('json')}
            >
              复制JSON
            </button>
            <button type="button" className="text-btn" onClick={() => setAssistantKey((k) => k + 1)}>
              刷新
            </button>
            <a className="text-btn" href={AGENT_URL} target="_blank" rel="noreferrer">
              新窗口
            </a>
          </div>

          {renderMessage()}

          <iframe
            key={assistantKey}
            className="assistant-frame"
            src={AGENT_URL}
            title="设备缺陷分析助手"
            allow="clipboard-read; clipboard-write"
          />

          <p className="assistant-hint">
            勾选记录 → 复制CSV → 在对话框粘贴提问；空白请点「新窗口」。
          </p>
        </div>
      )}

      {copyFallback !== null && (
        <div className="copy-modal-mask" onClick={() => setCopyFallback(null)}>
          <div className="copy-modal" onClick={(e) => e.stopPropagation()}>
            <div className="copy-modal-title">复制数据</div>
            <p className="copy-modal-hint">浏览器拦截了自动复制，请全选后按 Ctrl+C。</p>
            <textarea
              className="copy-modal-text"
              readOnly
              autoFocus
              value={copyFallback}
              onFocus={(e) => e.currentTarget.select()}
            />
            <div className="copy-modal-actions">
              <button type="button" className="primary-btn" onClick={() => setCopyFallback(null)}>
                关闭
              </button>
            </div>
          </div>
        </div>
      )}
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
  // 插件刚更新或网络抖动时按需模块会偶发拉取失败，给出可操作的提示
  if (/Failed to fetch dynamically imported module/i.test(detail)) {
    return '导出失败：模块加载失败（网络波动或插件刚更新），请重新打开插件后重试。';
  }
  return detail
    ? `导出失败：${detail}`
    : '导出失败，请确认插件运行在多维表格边栏中，且你对这些记录有查看权限。';
}
