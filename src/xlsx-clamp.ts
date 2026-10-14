/**
 * xlsx 字节层硬截断：按给定列数锁定文件 used_range。
 * 解包 zip，对每个 sheet XML 做手术：删除 expectedCols 之后的所有
 * 单元格、<col> 定义、合并区域，收拢 dimension / autoFilter / spans，再重新打包。
 * 与 js-sdk 无任何依赖，便于在 Node 中独立测试。
 * 任何异常都回退原始 buffer，保证导出永不因截断失败而中断。
 */

/** 列字母 → 1-based 序号（A=1, Z=26, AA=27…） */
export function colToIdx(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** 1-based 序号 → 列字母 */
export function idxToCol(index: number): string {
  let s = '';
  let i = index;
  while (i > 0) {
    const rem = (i - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

/** 收拢 "A1:XY3" 形式的引用范围，使终止列不超过 expectedCols */
function clampRef(ref: string, expectedCols: number): string {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return ref;
  const start = Math.min(colToIdx(m[1]), expectedCols);
  const end = Math.min(colToIdx(m[3]), expectedCols);
  return `${idxToCol(start)}${m[2]}:${idxToCol(Math.max(start, end))}${m[4]}`;
}

/** 对单个 sheet XML 做截断手术 */
function clampSheetXml(xml: string, expectedCols: number): string {
  let out = xml;

  // dimension / autoFilter 收拢
  out = out.replace(/(<dimension ref=")([^"]+)(")/, (_m, p1, ref, p3) => `${p1}${clampRef(ref, expectedCols)}${p3}`);
  out = out.replace(/(<autoFilter ref=")([^"]+)(")/, (_m, p1, ref, p3) => `${p1}${clampRef(ref, expectedCols)}${p3}`);

  // <col> 定义：起始列超界整条删除，终止列超界收拢
  out = out.replace(/<col [^>]*?\/>/g, (el) => {
    const min = Number(/min="(\d+)"/.exec(el)?.[1] ?? 0);
    const max = Number(/max="(\d+)"/.exec(el)?.[1] ?? 0);
    if (!min || !max) return el;
    if (min > expectedCols) return '';
    if (max > expectedCols) return el.replace(/max="\d+"/, `max="${expectedCols}"`);
    return el;
  });

  // 合并区域：终止列超界的整条删除，并修正 count
  out = out.replace(/<mergeCells count="\d+">([\s\S]*?)<\/mergeCells>/, (_m, inner: string) => {
    const kept = (inner.match(/<mergeCell ref="[^"]+"\/>/g) ?? []).filter((mc) => {
      const ref = /ref="([^"]+)"/.exec(mc)?.[1] ?? '';
      const end = ref.split(':')[1] ?? ref;
      return colToIdx(/^[A-Z]+/.exec(end)?.[0] ?? 'A') <= expectedCols;
    });
    return kept.length ? `<mergeCells count="${kept.length}">${kept.join('')}</mergeCells>` : '';
  });

  // 删除 expectedCols 之后的全部单元格。
  // 属性段 (?:[^">]|"[^"]*")* 不允许跨过 ">"；自闭合分支 \/> 与成对分支 (?<!\/)> 互斥，
  // 杜绝自闭合标签匹配失败后落入 ">[\s\S]*?</c>" 分支而跨越多个单元格，
  // 导致越界单元格被合法单元格"连带保留"。
  out = out.replace(
    /<c r="([A-Z]+)(\d+)"(?:[^">]|"[^"]*")*(?:\/>|(?<!\/)>(?:(?!<\/c>)[\s\S])*?<\/c>)/g,
    (el, letters: string) => (colToIdx(letters) > expectedCols ? '' : el),
  );

  // 行 spans 收拢
  out = out.replace(/spans="\d+:\d+"/g, `spans="1:${expectedCols}"`);

  return out;
}

/**
 * 截断 xlsx 的 used_range 到 expectedCols 列。
 * 失败时静默回退原始 buffer（不影响导出主流程）。
 */
export async function clampXlsxColumns(buffer: ArrayBuffer, expectedCols: number): Promise<ArrayBuffer> {
  if (!(expectedCols > 0)) return buffer;
  try {
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import('fflate');
    const files = unzipSync(new Uint8Array(buffer));
    let touched = false;
    for (const key of Object.keys(files)) {
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(key)) continue;
      const xml = strFromU8(files[key]);
      const clamped = clampSheetXml(xml, expectedCols);
      if (clamped !== xml) {
        files[key] = strToU8(clamped);
        touched = true;
      }
    }
    if (!touched) return buffer;
    const out = zipSync(files, { level: 6 });
    return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer;
  } catch {
    return buffer;
  }
}
