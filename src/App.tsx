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

/** 自动同步间隔选项（分钟） */
const AUTO_INTERVAL_OPTIONS = [1, 3, 5, 10, 30];
/** localStorage 持久化键：自动同步设置 */
const AUTO_SYNC_STORAGE_KEY = 'dms-auto-sync-config';

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
  const [lastSyncAt, setLastSyncAt] = useState<number | null>(null);
  /** 每秒刷新的当前时间，仅用于驱动倒计时显示 */
  const [autoNow, setAutoNow] = useState(Date.now());
  const autoTimerRef = useRef<number | null>(null);
  const autoCountdownRef = useRef<number | null>(null);
  const dmsWindowRef = useRef<Window | null>(null);

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

  /** 自动同步的主循环：读选中 → 压缩 → 发 hash 给 DMS
   *  窗口策略：插件自管一个名为 dms-da-main 的 DMS 标签页。
   *  - 已有可用引用 → 直接复用（只改 hash，不聚焦/不跳转标签页）
   *  - 引用不存在或已关闭 → 新开一个
   *  首次勾选时已在点击手势内同步打开窗口，后续定时同步复用该窗口。
   *  DMS 端 checkAndRedirect 会按缺陷编号长度自动跳到正确界面。
   */
  const runAutoSync = useCallback(async () => {
    const existing = dmsWindowRef.current;
    const targetWindow =
      existing && !existing.closed ? existing : window.open(DMS_URL, 'dms-da-main');
    if (!targetWindow) {
      setMessage({
        type: 'error',
        text: '浏览器拦截了 DMS 弹窗。请允许本页面弹出窗口后重试，或改用手动发送。',
      });
      return;
    }
    dmsWindowRef.current = targetWindow;
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
      // 给 DMS 窗口导航到带 hash 的新地址，DMS 监听 hashchange 自动刷新
      targetWindow.location.href = `${DMS_URL}?t=${Date.now()}#dms=${compressed}`;
      setLastSyncAt(Date.now());
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
        if (dmsWindow) dmsWindowRef.current = dmsWindow;
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
              : '请先在表格视图中勾选要导出的记录。',
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
          setLastSyncAt(Date.now());
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
              <span className="checkbox-hint">勾选后立即发送一次，之后按所选间隔自动发送当前选中记录</span>
            </label>

            {autoConfig.enabled && (
              <div className="auto-sync-panel">
                <div className="auto-sync-row">
                  <span className="auto-sync-label">同步间隔</span>
                  <div className="interval-group">
                    {AUTO_INTERVAL_OPTIONS.map((min) => (
                      <button
                        key={min}
                        type="button"
                        className={`interval-btn ${autoConfig.intervalMinutes === min ? 'active' : ''}`}
                        onClick={() => setAutoConfig((c) => ({ ...c, intervalMinutes: min }))}
                      >
                        {min}分
                      </button>
                    ))}
                  </div>
                </div>
                <div className="auto-sync-status">
                  {dmsWindowRef.current && !dmsWindowRef.current.closed ? (
                    <>
                      <span className="status-dot status-active" />
                      <span className="status-text">DMS 接收窗口就绪</span>
                    </>
                  ) : (
                    <>
                      <span className="status-dot status-warn" />
                      <span className="status-text">DMS 窗口未打开</span>
                    </>
                  )}
                  {nextSyncAt && (
                    <span className="status-countdown">
                      · 下次同步 {formatCountdown(nextSyncAt - autoNow)}
                    </span>
                  )}
                  {lastSyncAt && (
                    <span className="status-last">
                      · 上次 {new Date(lastSyncAt).toLocaleTimeString('zh-CN')}
                    </span>
                  )}
                </div>
              </div>
            )}
          </>
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
      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: '#bbb', marginTop: '4px' }}>
        <span>v2.1.0</span>
        <span>{new Date().toLocaleDateString('zh-CN')}</span>
      </div>
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
