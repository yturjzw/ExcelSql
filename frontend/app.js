// Excel SQL 查询台 —— 前端逻辑（单文件 / 多源关联 两种模式）
(() => {
  "use strict";

  const API_BASE = ""; // 同源部署；如前后端分离可改为 http://127.0.0.1:8000

  // 查询模式：'single' | 'multi'
  let mode = "single";
  // 单个数据源（单文件模式）—— 路径
  let singlePath = null;
  // 多源模式的源列表 [{id, alias, path, sheet, sheets, columns}]
  let sources = [];
  // 当前选中的源（决定字段区显示哪个源）
  let activeSourceId = null;
  // 字段递增计数器，保证源 id 唯一
  let sourceSeq = 0;

  // 与 backend/app.py 的 DEFAULT_LIMIT / MAX_RESULT_ROWS 保持一致（单点常量，避免漂移）。
  // 后端在 /api/query 会把超上限的 limit 夹到 MAX_RESULT_ROWS。
  const DEFAULT_LIMIT = 100;
  const MAX_RESULT_ROWS = 10000;
  // SQL 编辑器最大高度：内容超过此高度时编辑器内部滚动，避免单条长 SQL 撑满整屏。
  const SQL_EDITOR_MAX_H = 320;

  // ---------- 会话持久化（localStorage） ----------
  const STORE_KEY = "excel_sql_session_v1";
  function persistSession() {
    try {
      const payload = {
        mode,
        cells: cells.map((c) => ({
          sql: c.dom.cm ? c.dom.cm.getValue() : "",
          collapsed: c.dom.cellBody ? c.dom.cellBody.classList.contains("collapsed") : false,
          limit: c.limit ?? DEFAULT_LIMIT,
        })),
      };
      localStorage.setItem(STORE_KEY, JSON.stringify(payload));
    } catch (_) { /* 存储不可用/超限时静默降级 */ }
  }
  function restoreSession() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (!raw) return null;
      const p = JSON.parse(raw);
      if (!p || !Array.isArray(p.cells)) return null;
      return p;
    } catch (_) { return null; }
  }

  // ---------- 最近打开的历史文件 ----------
  const HIST_KEY = "excel_sql_history_v1";
  const HIST_MAX = 15;       // 历史最多保留条数
  const HIST_SHOWN = 6;      // 浮层内默认展开条数，其余折叠在「更多」
  const COMBOS_KEY = "excel_sql_combos_v1";   // 多源「组合」快照
  const COMBO_MAX = 5;       // 组合最多保留条数
  let recentFiles = loadHistory();
  let combos = loadCombos();
  let recentOpen = false;      // 浮层当前是否展开
  let recentAnchorEl = null;   // 当前浮层的锚点元素（最近打开按钮 / 源卡「选择」按钮）
  let recentTargetSrc = null;  // 浮层选中后的目标源（多源模式源卡打开时；null=单文件）
  let missingPaths = {};       // path -> true（文件已被移除，用于标红）
  let checkingMissing = false; // 防止并发检查

  function loadCombos() {
    try {
      const raw = localStorage.getItem(COMBOS_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr.filter((c) => c && Array.isArray(c.sources) && c.sources.length) : [];
    } catch (_) { return []; }
  }

  // 记录一次多源组合：查询成功后调用，同组合去重顶到最前，超限截断。
  function recordCombo(sourcesArr) {
    if (!sourcesArr || !sourcesArr.length) return;
    const snap = sourcesArr
      .filter((s) => s && s.path)
      .map((s) => ({ path: s.path, sheet: s.sheet || "", alias: s.alias || "" }));
    if (!snap.length) return;
    // 组合指纹：路径序列（稳定排序，忽略顺序差异）
    const fp = snap.map((s) => s.path).sort().join("|");
    combos = combos.filter((c) => c.fp !== fp);
    combos.unshift({ fp, sources: snap, time: Date.now() });
    if (combos.length > COMBO_MAX) combos.length = COMBO_MAX;
    try { localStorage.setItem(COMBOS_KEY, JSON.stringify(combos)); } catch (_) {}
  }

  // 加载缺失状态：用后端 exists 检查历史路径，缺失的标红
  async function checkMissingPaths() {
    if (checkingMissing) return;
    checkingMissing = true;
    const paths = recentFiles.map((e) => e.path);
    if (!paths.length) { checkingMissing = false; return; }
    try {
      const res = await request("/api/check_paths", { paths });
      const ex = (res && res.exists) || {};
      missingPaths = {};
      paths.forEach((p) => { if (ex[p] === false) missingPaths[p] = true; });
      if (recentOpen && recentFiles.length) renderRecentPop();
    } catch (_) { /* 后端不可用时静默，不标红 */ }
    checkingMissing = false;
  }

  function loadHistory() {
    try {
      const raw = localStorage.getItem(HIST_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr.filter((e) => e && typeof e.path === "string") : [];
    } catch (_) { return []; }
  }

  // 记录一次打开：同路径去重并顶到最前，超限截断，再重绘。
  function recordHistory(path, sheet) {
    if (!path) return;
    const now = Date.now();
    recentFiles = recentFiles.filter((e) => e.path !== path);
    recentFiles.unshift({ path, sheet: sheet || "", time: now });
    if (recentFiles.length > HIST_MAX) recentFiles.length = HIST_MAX;
    try { localStorage.setItem(HIST_KEY, JSON.stringify(recentFiles)); } catch (_) {}
    renderRecent();
  }

  // 从历史移除一条（文件被删/手动删除）
  function removeHistory(path) {
    recentFiles = recentFiles.filter((e) => e.path !== path);
    try { localStorage.setItem(HIST_KEY, JSON.stringify(recentFiles)); } catch (_) {}
    renderRecent();
  }

  function clearHistory() {
    recentFiles = [];
    try { localStorage.removeItem(HIST_KEY); } catch (_) {}
    renderRecent();
  }

  // 历史条目的相对时间：刚刚 / N 分钟前 / N 小时前 / 月-日 / 年-月-日
  function historyTime(ts) {
    if (!ts) return "";
    const diff = Date.now() - ts;
    const min = Math.floor(diff / 60000);
    if (min < 1) return "刚刚";
    if (min < 60) return `${min} 分钟前`;
    const h = Math.floor(min / 60);
    if (h < 24) return `${h} 小时前`;
    const d = new Date(ts);
    const y = d.getFullYear();
    const nowY = new Date().getFullYear();
    const md = `${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    return y === nowY ? md : `${y}-${md}`;
  }

  // 渲染「最近打开」：按钮显示条数徽标；有历史才显示按钮，浮层内容 → recentPop。
  function renderRecent() {
    const n = recentFiles.length;
    recentBtn.classList.toggle("hidden", !n);
    recentCount.textContent = n ? `(${n})` : "";
    if (!n && recentOpen) closeRecentPop();
    else if (recentOpen) renderRecentPop();
  }

  // 重建浮层内容（组合区 + 标题 + 清空 + 列表 + 更多），整体由 JS 构建
  function renderRecentPop() {
    recentPop.innerHTML = "";
    // 「最近组合」只在全局「最近打开」浮层展示（多源模式）：
    // 源卡文件夹图标打开的浮层是「选择历史文件」语义（给单个源换文件），
    // 混入会清空全部源的「恢复组合」会误操作，故排除。
    const isMultiPopup = !recentTargetSrc && mode === "multi";

    // 多源模式下、全局打开时：顶部展示「最近使用过的组合」区
    if (isMultiPopup) {
      if (combos.length) {
        const comboHead = document.createElement("div");
        comboHead.className = "recent-combo-head";
        comboHead.textContent = "最近组合";
        recentPop.appendChild(comboHead);

        const comboList = document.createElement("div");
        comboList.className = "recent-combo-list";
        combos.forEach((combo) => {
          const item = document.createElement("div");
          item.className = "recent-combo-item";
          item.title = combo.sources.map((s) => s.path).join("\n");

          // 摘要：源个数 + 每个文件名（省略中间）
          const names = combo.sources.map((s) => basename(s.path));
          const label = names.length <= 3 ? names.join(" · ") : `${names[0]} … +${names.length - 1} 个`;
          const nameEl = document.createElement("span");
          nameEl.className = "recent-combo-name";
          nameEl.textContent = label;

          const timeEl = document.createElement("span");
          timeEl.className = "recent-time";
          timeEl.textContent = historyTime(combo.time);

          item.appendChild(nameEl);
          item.appendChild(timeEl);

          // 点击恢复组合
          item.addEventListener("click", (e) => {
            e.stopPropagation();
            closeRecentPop();
            restoreCombo(combo);
          });
          comboList.appendChild(item);
        });
        recentPop.appendChild(comboList);
      } else {
        // 空态引导：让用户知道何时会出现组合
        const hint = document.createElement("div");
        hint.className = "recent-combo-hint";
        hint.textContent = "多源执行一次查询后，这里会自动记住你的源组合";
        recentPop.appendChild(hint);
      }
    }

    // 头部：标题 + 清空
    const head = document.createElement("div");
    head.className = "recent-head";
    const title = document.createElement("span");
    title.className = "recent-title";
    title.textContent = recentTargetSrc ? "选择历史文件" : "最近打开";
    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "recent-clear";
    clearBtn.textContent = "清空";
    clearBtn.title = "清空全部历史记录";
    clearBtn.addEventListener("click", (e) => {
      e.stopPropagation();   // 不触发外部点击关闭
      clearHistory();
    });
    head.appendChild(title);
    head.appendChild(clearBtn);
    recentPop.appendChild(head);

    // 有缺失文件时：附「清除缺失」快捷按钮（红色文字）
    const missingCount = recentFiles.filter((e) => missingPaths[e.path]).length;
    if (missingCount) {
      const delMissing = document.createElement("button");
      delMissing.type = "button";
      delMissing.className = "recent-del-missing";
      delMissing.textContent = `清除缺失的 ${missingCount} 条`;
      delMissing.title = "移除标记为〔文件不存在〕的历史条目";
      delMissing.addEventListener("click", (e) => {
        e.stopPropagation();
        recentFiles.forEach((en) => {
          if (missingPaths[en.path]) removeHistory(en.path);
        });
      });
      recentPop.appendChild(delMissing);
    }

    // 列表：默认前 HIST_SHOWN 条
    const list = document.createElement("div");
    list.className = "recent-list";
    const shown = recentFiles.slice(0, HIST_SHOWN);
    const rest = recentFiles.slice(HIST_SHOWN);
    shown.forEach((e) => list.appendChild(buildRecentItem(e)));
    recentPop.appendChild(list);

    // 超过默认展示条数时，追加「更多」展开按钮
    if (rest.length) {
      const moreBtn = document.createElement("button");
      moreBtn.type = "button";
      moreBtn.className = "recent-more";
      moreBtn.textContent = `更多（${rest.length}）`;
      moreBtn.addEventListener("click", () => {
        list.innerHTML = "";
        recentFiles.forEach((e) => list.appendChild(buildRecentItem(e)));
        moreBtn.remove();
      });
      recentPop.appendChild(moreBtn);
    }

    // 源卡打开时：底部附「选择本地文件…」入口，保持系统对话框可用；
    // 顶部加一行提示指引组合入口（组合不在此浮层展示，避免语义冲突）。
    if (recentTargetSrc) {
      const comboHint = document.createElement("div");
      comboHint.className = "recent-combo-hint";
      comboHint.textContent = "提示：点左栏「最近打开」可一键恢复整套源组合";
      recentPop.appendChild(comboHint);

      const pickLocal = document.createElement("button");
      pickLocal.type = "button";
      pickLocal.className = "recent-local";
      pickLocal.textContent = "选择本地文件…";
      pickLocal.addEventListener("click", async (e) => {
        e.stopPropagation();
        const src = recentTargetSrc;   // 先取目标再关闭
        closeRecentPop();
        await pickFileForSource(src);
      });
      recentPop.appendChild(pickLocal);
    }
  }

  // 恢复一个多源组合：清空当前 sources，按其快照逐个重建（只写状态，不加载字段，
  // 避免一次请求多个文件；用户点「查询」时由后端按路径读取）。
  async function restoreCombo(combo) {
    if (!combo || !combo.sources || !combo.sources.length) return;
    sources = [];
    combo.sources.forEach((s) => {
      const src = createSource(s.alias || `t${sources.length + 1}`, s.path, s.sheet || "");
      src.sheets = [];
      src.columns = [];
    });
    activeSourceId = sources.length ? sources[0].id : null;
    renderSources();
    refreshCompletion();   // 新路径进入 SQL 补全候选
    showStatus(uploadStatus, `已恢复组合：${combo.sources.length} 个源`, "ok");
  }

  // 切换浮层展开 / 收起；锚点元素用于定位（默认最近打开按钮）
  function toggleRecentPop(force, anchorEl) {
    if (anchorEl) recentAnchorEl = anchorEl;
    const open = force !== undefined ? force : !recentOpen;
    recentOpen = open;
    if (open) {
      renderRecentPop();
      positionRecentPop();
      recentPop.classList.remove("hidden");
      checkMissingPaths();   // 打开时异步校验历史文件中哪些已被移除
    } else {
      recentPop.classList.add("hidden");
    }
  }

  // 从源卡「选择」按钮打开浮层：目标源 + 锚点 = 该按钮
  function openRecentPopForSource(src, anchorEl) {
    recentTargetSrc = src;
    recentAnchorEl = anchorEl;
    renderRecentPop();
    positionRecentPop();
    recentOpen = true;
    recentPop.classList.remove("hidden");
    checkMissingPaths();   // 打开时异步校验
  }

  // 关闭浮层时清掉目标源（切回「最近打开」）
  function closeRecentPop() {
    recentOpen = false;
    recentTargetSrc = null;
    recentAnchorEl = null;
    recentPop.classList.add("hidden");
  }

  // 把浮层定位到锚点下方（fixed，坐标相对 viewport）
  function positionRecentPop() {
    const anchor = recentAnchorEl || recentBtn;
    if (!anchor) return;
    const r = anchor.getBoundingClientRect();
    const w = 300;                    // 浮层宽度
    let left = r.left;
    const vw = document.documentElement.clientWidth;
    if (left + w > vw - 8) left = Math.max(8, vw - w - 8);
    recentPop.style.left = left + "px";
    recentPop.style.top = (r.bottom + 6) + "px";
  }

  // 外部点击 / Esc 关闭浮层
  function closeRecentPopOnOutside(e) {
    if (recentOpen && !recentPop.contains(e.target) && e.target !== recentAnchorEl && !(recentAnchorEl && recentAnchorEl.contains(e.target))) {
      closeRecentPop();
    }
  }

  function buildRecentItem(entry) {
    const name = basename(entry.path) || entry.path;
    const missing = !!missingPaths[entry.path];
    const row = document.createElement("div");
    row.className = "recent-item" + (missing ? " recent-missing" : "");
    row.title = missing ? `${entry.path}\n（文件已被移动或删除）` : entry.path;

    const nameEl = document.createElement("span");
    nameEl.className = "recent-name";
    nameEl.textContent = name;

    const timeEl = document.createElement("span");
    timeEl.className = "recent-time";
    timeEl.textContent = historyTime(entry.time);

    const del = document.createElement("button");
    del.type = "button";
    del.className = "recent-del";
    del.textContent = "×";
    del.title = "从历史移除";
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      removeHistory(entry.path);
    });

    row.appendChild(nameEl);
    row.appendChild(timeEl);
    row.appendChild(del);

    row.addEventListener("click", () => loadHistoryEntry(entry));
    return row;
  }

  // 点击历史条目：按当前模式复用现有加载链路（单文件 openSingle / 多源 refreshSource）。
  async function loadHistoryEntry(entry) {
    if (!entry || !entry.path) return;
    const targetSrc = recentTargetSrc;   // 先取目标源，再关闭浮层（关闭会清掉它）
    closeRecentPop();
    if (mode === "multi" && targetSrc) {
      // 源卡浮层：赋给打开浮层时指定的那个源（而不是当前激活的源）
      const src = targetSrc;
      const isFirstPick = !src.path;
      src.path = entry.path;
      src.sheet = entry.sheet || "";
      await refreshSource(src, !isFirstPick);
    } else if (mode === "multi") {
      // 全局浮层（多源模式，无指定源）：赋给当前激活源
      const targets = sources.length ? sources : [createSource("t1", "", "")];
      const src = targets.find((s) => s.id === activeSourceId) || targets[0];
      const isFirstPick = !src.path;
      src.path = entry.path;
      src.sheet = entry.sheet || "";
      await refreshSource(src, !isFirstPick);
    } else {
      singlePath = entry.path;
      // 恢复该文件上次用的工作表：若当前下拉有对应项则选中，否则留默认
      const need = entry.sheet || "";
      if (need && Array.from(sheetInput.options).some((o) => o.value === need)) {
        sheetInput.value = need;
      } else {
        sheetInput.value = "";
      }
      await openSingle();
    }
  }

  // 源卡「选择本地文件…」：走系统对话框，赋给指定源
  async function pickFileForSource(src) {
    if (!src) return;
    const d = await request("/api/pick_file", {});
    if (d && d.path) {
      const isFirstPick = !src.path;
      src.path = d.path;
      src.sheet = "";
      await refreshSource(src, !isFirstPick);
    }
  }

  // DOM 元素
  const modeSingleBtn = document.getElementById("modeSingleBtn");
  const modeMultiBtn = document.getElementById("modeMultiBtn");
  const singleModeBox = document.getElementById("singleModeBox");
  const multiModeBox = document.getElementById("multiModeBox");
  const pickBtn = document.getElementById("pickBtn");
  const addSourceBtn = document.getElementById("addSourceBtn");
  const sourceList = document.getElementById("sourceList");
  const fileInfo = document.getElementById("fileInfo");
  const fileNameEl = document.getElementById("fileName");
  const clearSingleBtn = document.getElementById("clearSingleBtn");
  const uploadStatus = document.getElementById("uploadStatus");
  const openStages = document.getElementById("openStages");
  const recentBtn = document.getElementById("recentBtn");
  const recentCount = document.getElementById("recentCount");
  const recentPop = document.getElementById("recentPop");
  const fieldSearch = document.getElementById("fieldSearch");
  const fieldsBox = document.getElementById("fieldsBox");
  const sourceLabel = document.getElementById("sourceLabel");
  const sheetCtrl = document.getElementById("sheetCtrl");
  const sheetInput = document.getElementById("sheetInput");
  const notebook = document.getElementById("notebook");
  const cellOutline = document.getElementById("cellOutline");
  const leftPanel = document.getElementById("leftPanel");
  const splitter = document.getElementById("splitter");
  const sidebarToggle = document.getElementById("sidebarToggle");
  const layoutEl = document.querySelector(".layout");

  // Notebook cells 状态
  let cells = [];        // [{ id, lastResult, dom:{...} }]
  let cellSeq = 0;
  let activeCellId = null;  // 当前聚焦的 cell（插入字段/运行的目标）

  // ---------- 工具 ----------
  function showStatus(el, text, kind) {
    el.textContent = text || "";
    el.className = "status-msg" + (kind ? " " + kind : "");
  }

  function popError(err) {
    return err?.detail || err?.message || String(err || "未知错误");
  }

  async function request(path, body) {
    let res;
    try {
      res = await fetch(API_BASE + path, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
    } catch (e) {
      throw new Error("无法连接后端，请确认服务已启动：" + e.message);
    }
    if (!res.ok) {
      let detail = res.statusText;
      try { detail = (await res.json()).detail || detail; } catch (_) { /* ignore */ }
      const err = new Error(detail);
      err.detail = detail;
      throw err;
    }
    return res.json();
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (m) => (
      { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]
    ));
  }

  // 耗时格式化：ms 级足够；≥1s 显示两位小数的秒
  function fmtMs(ms) {
    if (ms == null || !Number.isFinite(Number(ms))) return "";
    const n = Number(ms);
    return n >= 1000 ? (n / 1000).toFixed(2) + " s" : Math.round(n) + " ms";
  }

  // ---------- 选择文件后的连续阶段展示（每步 + 实时耗时） ----------
  function clearOpenStages() {
    openStages.innerHTML = "";
  }

  // 按 key 增/改一个阶段行；status: pending | active | done | error。
  // active 阶段无需传 ms：自动记录 startMs 并由全局 ticker 每 120ms 刷新实时耗时。
  // 终态（done/error）阶段会附带一个删除按钮，允许用户手动移除该条日志。
  function upsertStage(key, label, status, ms, altLabel) {
    let row = openStages.querySelector(`[data-stage-key="${key}"]`);
    if (!row) {
      row = document.createElement("div");
      row.className = "open-stage";
      row.dataset.stageKey = key;
      openStages.appendChild(row);
    }
    const st = status || "pending";
    row.className = "open-stage " + st;
    if (st === "active" && !row.dataset.startMs) {
      row.dataset.startMs = performance.now();
    }
    row.innerHTML = "";
    const dot = document.createElement("span");
    dot.className = "stage-dot";
    const lab = document.createElement("span");
    lab.className = "stage-label";
    lab.textContent = altLabel || label;
    const time = document.createElement("span");
    time.className = "stage-time";
    if (st === "active") {
      time.textContent = fmtMs(performance.now() - Number(row.dataset.startMs));
    } else {
      time.textContent = fmtMs(ms);
    }
    row.appendChild(dot);
    row.appendChild(lab);
    row.appendChild(time);

    // 终态日志：附加删除按钮，允许用户手动移除该条日志
    if (st === "done" || st === "error") {
      const rm = document.createElement("button");
      rm.type = "button";
      rm.className = "stage-remove";
      rm.title = "删除此条日志";
      rm.textContent = "×";
      rm.addEventListener("click", () => removeStage(key));
      row.appendChild(rm);
    }
  }

  function removeStage(key) {
    const row = openStages.querySelector(`[data-stage-key="${key}"]`);
    if (row) row.remove();
  }

  // 全局 ticker：刷新所有 active 阶段的实时耗时（仅在有 active 行时才有开销）
  let _stageTicker = null;
  function ensureStageTicker() {
    if (_stageTicker) return;
    _stageTicker = setInterval(() => {
      openStages.querySelectorAll(".open-stage.active").forEach((row) => {
        const t = row.querySelector(".stage-time");
        if (t && row.dataset.startMs) {
          t.textContent = fmtMs(performance.now() - Number(row.dataset.startMs));
        }
      });
    }, 120);
  }

  // ---------- SQL 高亮（交给 CodeMirror 的 text/x-sql 模式） ----------
  // 把原关键词与 SUM/COUNT 等函数合并进一张关键词表：二者统一映射到 cm-keyword
  // （CSS 中为“蓝 + 加粗”）。null/true/false 属于原子字面量（cm-atom，蓝色）。
  const SQL_KEYWORD_WORDS = (
    "select from where join inner outer left right full cross on using group by order " +
    "having limit offset union all distinct as and or not in is null like between exists " +
    "case when then else end cast with view table create replace insert update delete drop " +
    "into values set primary key foreign references index unique check default constraint " +
    "asc desc true false if over partition window fetch first next rows only recursive " +
    "materialized describe show summarize pragma " +
    "sum count avg min max group_concat string_agg coalesce nullif extract date_trunc " +
    "now current_date current_timestamp upper lower length trim replace substr substring " +
    "round abs ceil floor sqrt mod concat arg_max arg_min list regexp_matches first_value " +
    "last_value row_number rank dense_rank lag lead ntile percentile_cont percentile_disc"
  );

  // CodeMirror 的 sql 模式要求 keywords 是「对象（set）」，不能是字符串；
  // 这里在应用层构建好，覆盖 text/x-sql 默认关键词表。
  function sqlWordSet(str) {
    const o = {};
    str.split(/\s+/).forEach((w) => { if (w) o[w] = true; });
    return o;
  }

  // mode 对象：name="sql" 指向 sql.js 注册的 mode；keywords 覆盖默认关键词表。
  // identifierQuote 用双引号（DuckDB 约定），让补全在引用含空格的标识符时与 quoteIdent 一致。
  const SQL_MODE_SPEC = {
    name: "sql",
    keywords: sqlWordSet(SQL_KEYWORD_WORDS),
    identifierQuote: '"',
  };

  // ── 全角标点高亮 ──
  // 中文输入法下，全角/半角的括号、逗号混排难以分辨。用 CodeMirror 的
  // addOverlay 叠加一个只认「全角标点/全角空格」的层，标记为 fw-punct
  // token（CSS 中给橙色+淡橙底），纯视觉标记、不改任何内容。
  // 实现要点：
  //   - overlay 对整个行独立跑一遍（看不到 base token），但 stream.lineOracle
  //     是 Context，可用 baseToken(n) 查该位置的 base mode 类型；
  //   - 字符串/注释内部的全角标点是「数据内容」，不与代码结构混淆，跳过不标。
  const FULLWIDTH_PUNCT_RE = /[（），。；：！？、【】《》「」『』“”‘’～·…—\u3000]/;
  const FULLWIDTH_PUNCT_OVERLAY = {
    token(stream) {
      const pos = stream.pos;
      const ch = stream.next();
      if (!ch) return null;
      if (!FULLWIDTH_PUNCT_RE.test(ch)) return null;
      // 跳过字符串/注释内部的全角标点（数据内容，不算“代码标点错误”）
      const ctx = stream.lineOracle;
      if (ctx && typeof ctx.baseToken === "function") {
        const bt = ctx.baseToken(pos);
        if (bt && /(^|\s)(string|comment)(\s|$)/.test(bt.type || "")) return null;
      }
      return "fw-punct";
    },
  };

  // 编辑器高度自适应内容：随内容完整展开，但设上限——超长时编辑器内部滚动，
  // 避免单条 SQL 撑满整屏、把结果挤到屏幕外。上限以内的行为与原自适应一致。
  function autoSizeCell(cell) {
    const cm = cell.dom.cm;
    if (!cm || cell.autoFit === false) return;
    const lineH = cm.defaultTextHeight();      // 整数像素行高（applyCellFont 已取整）
    const pad = 8;                             // .CodeMirror-lines 默认 padding:4px 0
    const contentH = cm.lineCount() * lineH + pad;
    const h = Math.min(SQL_EDITOR_MAX_H, Math.max(lineH * 2 + pad, contentH));
    cm.setSize("100%", h + "px");
    cell.lastAutoH = h;
  }

  // 字段/表名标识符：安全字符裸写，否则加双引号
  function quoteIdent(name) {
    const s = String(name);
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(s) ? s : '"' + s.replace(/"/g, '""') + '"';
  }

  // ---------- SQL 自动补全：表名/别名 → 列名映射 ----------
  // 单文件模式：表名 data；多源模式：别名 = 表名（别名可热改）。
  // 返回 { tables: { TABLE: {columns: [...] or [..], text}, defaultTable: "..." } }
  function completionContext() {
    const tables = {};
    let defaultTable = null;
    if (mode === "single") {
      const cols = (_singleColumns || []).map((c) => c.name);
      tables["data"] = { columns: cols, text: "data" };
      defaultTable = "data";
    } else {
      (sources || []).forEach((s) => {
        if (!s.alias) return;
        const cols = (s.columns || []).map((c) => c.name);
        tables[s.alias] = { columns: cols, text: s.alias };
      });
      // 以当前选中的源作为默认补全表（优先级最高的字段来源）
      const active = sources.find((s) => s.id === activeSourceId);
      if (active && active.alias) defaultTable = active.alias;
    }
    return { tables, defaultTable };
  }

  function refreshCompletion() {
    // 补全上下文由 sqlHintFn 每次触发时惰性读取，此函数为扩展预留的显式刷新入口
  }

  // 每次触发补全时动态读取最新表/列映射，再委托给 sql-hint.js 的 "sql" helper。
  // SQL 关键字补全则由模式自带的 keywords 词表提供（sql-hint 内部自行获取）。
  function sqlHintFn(editor) {
    const ctx = completionContext();
    return CodeMirror.hint.sql(editor, {
      tables: ctx.tables,
      defaultTable: ctx.defaultTable,
    });
  }

  function clearAllCellResults() {
    // 切换模式/清除文件后，所有 cell 的旧结果均失效，清空展示
    cells.forEach((cell) => {
      cell.lastResult = null;
      cell.dom.result.classList.add("hidden");
      cell.dom.status.textContent = "";
      cell.dom.status.className = "cell-status";
    });
  }

  // ---------- 模式切换 ----------
  // 默认 SQL 随模式变化：单文件用 data，多源用第一个源别名并带注释提示。
  function defaultCellSql() {
    if (mode === "multi") {
      const first = sources.find((s) => s.alias) || sources[0];
      const alias = first && first.alias ? first.alias : "t1";
      return `SELECT * FROM ${quoteIdent(alias)} LIMIT 10;`;
    }
    return "SELECT * FROM data LIMIT 10;";
  }

  function setMode(next) {
    if (next === mode) return;
    // 单源 / 多源是两个独立的状态空间（singlePath 与 sources[] 各自记忆），
    // 切换只是显示/隐藏，不互相清空——来回切换时各自的数据源与字段都还在。

    mode = next;
    const isSingle = next === "single";
    modeSingleBtn.classList.toggle("active", isSingle);
    modeMultiBtn.classList.toggle("active", !isSingle);
    singleModeBox.classList.toggle("hidden", !isSingle);
    multiModeBox.classList.toggle("hidden", isSingle);
    sheetCtrl.classList.toggle("hidden", !isSingle);

    clearAllCellResults();
    // 清掉上一次模式的瞬态提示（状态文案 + 选择文件耗时阶段线），避免被误认为当前模式的状态
    showStatus(uploadStatus, "", "");
    clearOpenStages();

    // 切到多源（关联查询场景）只保留一个查询单元：单源的多个 cell 缩减为 1 个，
    // 保留当前聚焦的那个（否则保留第一个）。
    if (next === "multi" && cells.length > 1) {
      const keepId = (activeCellId && getCell(activeCellId)) ? activeCellId : cells[0].id;
      for (let i = cells.length - 1; i >= 0; i--) {
        if (cells[i].id === keepId) continue;
        const c = cells[i];
        const nd = c.dom.root.nextElementSibling;
        if (nd && nd.classList.contains("cell-divider")) nd.remove();
        c.dom.root.remove();
        cells.splice(i, 1);
      }
      activeCellId = keepId;
      refreshCellNumbers();
      refreshOutline();
    }

    // 切换模式后，旧 SQL 引用的表名（单源的 data / 多源的别名）在新模式下可能不存在，
    // 直接执行会报错。因此把每个 cell 的 SQL 统一重置为新模式的默认模板，
    // 结果也已在上方 clearAllCellResults() 清空。
    // （数据源状态 singlePath / sources[] 仍各自保留，互不影响。）
    const newDefault = defaultCellSql();
    cells.forEach((cell) => {
      if (cell.dom.cm && cell.dom.cm.getValue().trim() !== newDefault.trim()) {
        cell.dom.cm.setValue(newDefault);
        autoSizeCell(cell);
      }
    });

    if (isSingle) {
      renderSingle();
    } else {
      renderSources();
    }
    persistSession();
  }

  // ---------- 单文件模式 ----------
  function renderSingle() {
    fileNameEl.textContent = singlePath || "";
    fileNameEl.title = singlePath || "";
    fileInfo.classList.toggle("hidden", !singlePath);
    sourceLabel.textContent = "";   // 单文件无别名，清掉多源残留的「（别名）」
    renderFields(singlePath ? getSingleColumns() : null);
  }

  async function pickFile() {
    showStatus(uploadStatus, "正在等待选择文件…", "");
    try {
      const data = await request("/api/pick_file", {});
      if (!data || data.cancelled || !data.path) {
        showStatus(uploadStatus, "已取消选择", "");
        return;
      }
      singlePath = data.path;
      await openSingle();
    } catch (e) {
      showStatus(uploadStatus, popError(e), "error");
    }
  }

  async function openSingle() {
    if (!singlePath) return;
    const sheet = sheetInput.value.trim() || null;
    // 数据源即将改变（换文件/换工作表），旧 cell 结果随之失效
    clearAllCellResults();
    clearOpenStages();
    ensureStageTicker();
    upsertStage("open", "读取字段与工作表", "active");
    try {
      const data = await request("/api/open", { path: singlePath, sheet });
      singlePath = data.path;
      // 打开成功后记入「最近打开」（后续点击可一键重新加载）
      recordHistory(singlePath, sheet);
      // 缓存字段以支持搜索过滤（挂在单例对象上）
      _singleColumns = data.columns || [];
      renderSingleSheets(data.sheets, sheet);
      renderFields(_singleColumns);
      // 连续展示：字段/工作表读取完成 → 打点耗时；随后进入后台预处理阶段。
      const t = data.timings || {};
      upsertStage("open", "读取字段与工作表", "done", t.total_ms);
      refreshCompletion();
      // 选文件后后台预构建缓存：轮询直到就绪，期间持续展示预处理耗时。
      pollPrebuild(singlePath, sheet, _singleColumns.length, t.total_ms);
    } catch (e) {
      upsertStage("open", "读取字段与工作表", "error");
      showStatus(uploadStatus, popError(e), "error");
      // 文件被移动/删除：该历史条目已失效，自动从最近打开移除
      if (singlePath && /不存在/.test(String(popError(e)))) {
        removeHistory(singlePath);
        showStatus(uploadStatus, "文件已被移动或删除，已从「最近打开」移除", "error");
      }
    }
  }

  // 轮询后台预构建状态，就绪后恢复「就绪」提示。
  // openMs：/api/open 已花费的毫秒（用于阶段展示首行时间换算）。
  async function pollPrebuild(path, sheet, fieldCount, openMs) {
    upsertStage("build", "后台预处理数据", "active");
    const started = performance.now();
    // 自适应退避：初始 300ms 起步，每次退避增长，上限 2s；
    // 阶段展示的实时耗时由全局 ticker 独立刷新，不依赖轮询节奏。
    let delay = 300;
    for (let i = 0; i < 240; i++) {   // 最多约 5 分钟
      try {
        const s = await request("/api/prebuild_status", { path, sheet: sheet || null });
        if (s.status === "done") {
          upsertStage("build", "后台预处理数据", "done", s.elapsed_ms);
          showStatus(uploadStatus, `已就绪，共 ${fieldCount} 个字段`, "ok");
          return;
        }
        if (s.status === "error") {
          upsertStage("build", "后台预处理数据", "done", s.elapsed_ms, "后台预处理（兼容模式）");
          showStatus(uploadStatus, "已就绪（兼容模式）", "ok");
          return;
        }
        if (s.status === "none") {
          // CSV 等无需预构建的类型：没有后台任务，直接结束本阶段
          removeStage("build");
          showStatus(uploadStatus, `已就绪，共 ${fieldCount} 个字段`, "ok");
          return;
        }
        if (s.status === "running") {
          upsertStage("build", "后台预处理数据", "active");
        }
      } catch (_) { /* 忽略轮询失败，稍后再试 */ }
      await new Promise((r) => setTimeout(r, delay));
      delay = Math.min(2000, Math.round(delay * 1.2));
    }
    // 超时仍未就绪：标记完成但提示兜底，避免永久转圈
    upsertStage("build", "后台预处理数据", "done", performance.now() - started);
    showStatus(uploadStatus, `已就绪，共 ${fieldCount} 个字段`, "ok");
  }

  let _singleColumns = [];
  function getSingleColumns() { return _singleColumns; }

  function renderSingleSheets(sheets, selected) {
    sheetInput.innerHTML = "";
    const def = document.createElement("option");
    def.value = "";
    def.textContent = "默认（第一张）";
    sheetInput.appendChild(def);
    (sheets || []).forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sheetInput.appendChild(opt);
    });
    sheetInput.value = selected || "";
  }

  function clearSingle() {
    singlePath = null;
    _singleColumns = [];
    // 重置工作表下拉
    sheetInput.innerHTML = '<option value="">默认（第一张）</option>';
    clearAllCellResults();
    showStatus(uploadStatus, "已去除文件，请重新选择", "");
    renderSingle();
  }

  // ---------- 多源模式 ----------
  function createSource(alias, path, sheet) {
    const src = {
      id: ++sourceSeq,
      alias: alias || "",
      path: path || "",
      sheet: sheet || "",
      sheets: [],
      columns: [],
    };
    sources.push(src);
    activeSourceId = src.id;
    return src;
  }

  function removeSource(id) {
    sources = sources.filter((s) => s.id !== id);
    if (activeSourceId === id) {
      activeSourceId = sources.length ? sources[0].id : null;
    }
    renderSources();
  }

  function renderSources() {
    sourceList.innerHTML = "";
    if (!sources.length) {
      sourceList.innerHTML = '<p class="placeholder">点击「添加数据源」开始</p>';
      renderFields(null);
      sourceLabel.textContent = "";
      return;
    }

    sources.forEach((src, idx) => {
      const row = document.createElement("div");
      row.className = "source-item" + (src.id === activeSourceId ? " active" : "");

      // 卡头：序号徽标 + 别名输入 + 删除
      const head = document.createElement("div");
      head.className = "source-item-head";

      const badge = document.createElement("span");
      badge.className = "src-index";
      badge.textContent = String(idx + 1);
      badge.title = "数据源顺序（即 SQL 关联顺序）";

      const aliasInput = document.createElement("input");
      aliasInput.type = "text";
      aliasInput.className = "alias-input";
      aliasInput.placeholder = "别名";
      aliasInput.value = src.alias;
      aliasInput.title = "SQL 中的表名（字母/数字/下划线）";
      aliasInput.addEventListener("input", () => {
        src.alias = aliasInput.value.trim();
        // 若正在编辑的是当前选中的源，同步更新字段区标题，但不重绘（避免打断输入）
        if (src.id === activeSourceId) {
          sourceLabel.textContent = src.alias ? `（${src.alias}）` : "";
        }
      });

      const del = document.createElement("button");
      del.type = "button";
      del.className = "src-del";
      del.innerHTML = ICON_TRASH;   // 图标化，省掉「删除」文字宽度
      del.title = "移除此数据源";
      del.addEventListener("click", (e) => {
        e.stopPropagation();   // 不触发行选中
        removeSource(src.id);
      });

      // 选择/更换文件按钮：图标化（文件夹），放在卡头行（序号 + 别名 + 选择 + 删除），
      // 省去「选择/更换」文字宽度。点击弹出「最近打开」浮层选历史文件（赋给本源），
      // 或走系统对话框。tooltip 随状态区分「选择」「更换」。
      const pickOne = document.createElement("button");
      pickOne.type = "button";
      pickOne.className = "src-pick";
      pickOne.innerHTML = ICON_FOLDER;
      pickOne.title = src.path ? "更换此数据源的文件（含最近打开）" : "为该数据源选择文件（含最近打开）";
      pickOne.addEventListener("click", (e) => {
        e.stopPropagation();   // 不触发行选中
        // 无历史文件时直接走系统对话框；有则弹出浮层
        if (recentFiles.length) {
          openRecentPopForSource(src, pickOne);
        } else {
          pickFileForSource(src);
        }
      });

      head.appendChild(badge);
      head.appendChild(aliasInput);
      head.appendChild(pickOne);
      head.appendChild(del);

      // 文件行：文件名（单行省略）+ 字段数徽标（选择/更换已在卡头行）
      const pathRow = document.createElement("div");
      pathRow.className = "src-path-row";

      const pathDisplay = document.createElement("div");
      pathDisplay.className = "src-path";
      pathDisplay.textContent = src.path ? basename(src.path) : "（未选择文件）";
      pathDisplay.title = src.path || "";

      pathRow.appendChild(pathDisplay);
      if (src.columns && src.columns.length) {
        const colCount = document.createElement("span");
        colCount.className = "src-colcount";
        colCount.textContent = `${src.columns.length} 字段`;
        pathRow.appendChild(colCount);
      }

      // 工作表下拉：仅多 sheet 时展示，单 sheet / 未选文件时隐藏
      let sheetRow = null;
      let sheetSel = null;
      if (src.sheets && src.sheets.length > 1) {
        sheetRow = document.createElement("div");
        sheetRow.className = "src-sheet-row";
        sheetSel = document.createElement("select");
        sheetSel.className = "src-sheet";
        const defOpt = document.createElement("option");
        defOpt.value = "";
        defOpt.textContent = "默认（第一张）";
        sheetSel.appendChild(defOpt);
        src.sheets.forEach((n) => {
          const o = document.createElement("option");
          o.value = n;
          o.textContent = n;
          sheetSel.appendChild(o);
        });
        sheetSel.value = src.sheet || "";
        sheetSel.addEventListener("change", async () => {
          src.sheet = sheetSel.value;
          await refreshSource(src);
        });
        sheetRow.appendChild(sheetSel);
      }

      // 组装：head / pathRow（含 pickOne）/ sheetRow（可选）
      row.appendChild(head);
      row.appendChild(pathRow);
      if (sheetRow) row.appendChild(sheetRow);

      // 选中该源显示字段（点击控件不触发行切换，由控件自身 stopPropagation 或这里兜底）
      row.addEventListener("click", (e) => {
        if (e.target.closest("button, input, select")) return;
        activeSourceId = src.id;
        renderSources();
      });

      sourceList.appendChild(row);
    });

    const active = sources.find((s) => s.id === activeSourceId) || sources[0];
    activeSourceId = active.id;
    sourceLabel.textContent = active.alias ? `（${active.alias}）` : "";
    renderFields(active.columns);
  }

  async function refreshSource(src, clearResults = true) {
    if (!src.path) return;
    // 换文件/换 sheet 时数据变了，旧结果应失效；给新源首次选文件时其他源未变，
    // 保留已有查询结果（clearResults=false）。
    if (clearResults) clearAllCellResults();
    clearOpenStages();
    ensureStageTicker();
    upsertStage("open", "读取字段与工作表", "active");
    try {
      const data = await request("/api/open", { path: src.path, sheet: src.sheet || null });
      src.path = data.path;
      // 打开成功后记入「最近打开」
      recordHistory(src.path, src.sheet || "");
      src.sheets = data.sheets || [];
      src.columns = data.columns || [];
      renderSources();
      refreshCompletion();  // 新字段/别名进入补全候选
      const t = data.timings || {};
      upsertStage("open", "读取字段与工作表", "done", t.total_ms);
      // 选完源后后台预构建缓存：轮询提示
      pollPrebuild(src.path, src.sheet || null, src.columns.length, t.total_ms);
    } catch (e) {
      upsertStage("open", "读取字段与工作表", "error");
      showStatus(uploadStatus, popError(e), "error");
      // 文件被移动/删除：该历史条目已失效，自动从最近打开移除
      if (src.path && /不存在/.test(String(popError(e)))) {
        removeHistory(src.path);
        showStatus(uploadStatus, "文件已被移动或删除，已从「最近打开」移除", "error");
      }
    }
  }

  function basename(p) {
    const parts = String(p).split(/[\\/]/);
    return parts[parts.length - 1];
  }

  // ---------- 字段渲染（含搜索过滤） ----------
  function renderFields(columns) {
    const kw = fieldSearch.value.trim().toLowerCase();
    const list = (columns || []).filter((c) =>
      !kw || String(c.name).toLowerCase().includes(kw)
    );

    fieldsBox.innerHTML = "";
    if (!columns || !columns.length) {
      fieldsBox.innerHTML = '<p class="placeholder">读取字段后，这里自动列出列名与类型</p>';
      return;
    }
    if (!list.length) {
      fieldsBox.innerHTML = '<p class="placeholder">无匹配字段</p>';
      return;
    }
    list.forEach((c) => {
      const chip = document.createElement("span");
      chip.className = "field-chip";
      chip.title = c.type;
      chip.innerHTML = `${escapeHtml(c.name)}<span class="ftype">${escapeHtml(c.type)}</span>`;
      chip.addEventListener("click", () => insertField(c.name));
      fieldsBox.appendChild(chip);
    });
  }

  function insertField(name) {
    let token;
    if (mode === "multi") {
      const active = sources.find((s) => s.id === activeSourceId);
      const alias = active?.alias || "";
      token = alias
        ? `${quoteIdent(alias)}.${quoteIdent(name)}`
        : quoteIdent(name);
    } else {
      token = quoteIdent(name);
    }

    // 插入到当前聚焦的 cell；没有聚焦则用最后一个 cell
    let cell = getCell(activeCellId);
    if (!cell) {
      cell = cells[cells.length - 1];
      if (!cell) return;
      activeCellId = cell.id;
    }
    const cm = cell.dom.cm;
    // 在光标处替换插入（替换当前选区），并把光标落到插入文本之后
    cm.replaceSelection(token, "around");
    cm.focus();
  }

  // ---------- Notebook cells ----------
  // 编辑器字号缩放（Ctrl + 滚轮 / Ctrl+加号 / Ctrl+减号 / Ctrl+0 复位）
  // CodeMirror 把光标/文字/行号/滚动统一在同一套布局里，缩放只需改根元素字号 +
  // 让 CM 重新测量（refresh），不再有 textarea vs pre 两层取整不一致的问题。
  const SQL_FONT_BASE = 13.5;
  const SQL_LINE_HEIGHT = 22;   // 与 CSS 的 line-height: 22px 一一对应，缩放时同步缩放
  const SQL_FONT_MIN = 0.7;
  const SQL_FONT_MAX = 1.9;
  const SQL_FONT_STEP = 0.1;

  function applyCellFont(cell) {
    const cm = cell.dom.cm;
    if (!cm) return;
    const scale = cell.fontScale || 1;
    const px = (SQL_FONT_BASE * scale).toFixed(2) + "px";
    // 行高四舍五入到整数像素，避免小数行高造成行盒步长取整漂移
    const lh = Math.round(SQL_LINE_HEIGHT * scale) + "px";
    const root = cm.getWrapperElement();
    root.style.fontSize = px;
    root.style.lineHeight = lh;
    cm.refresh();
  }

  function zoomCell(cell, dir) {
    const cur = cell.fontScale || 1;
    let scale = dir === 0 ? 1 : cur + dir * SQL_FONT_STEP;
    scale = Math.max(SQL_FONT_MIN, Math.min(SQL_FONT_MAX, scale));
    cell.fontScale = scale;
    applyCellFont(cell);
    autoSizeCell(cell);       // 字号变化后重新适配高度（行高变了）
  }

  // cell 头部图标（折叠 chevron / 删除 trash），与整体黑白灰极简风一致
  const ICON_CHEVRON_DOWN = '<svg viewBox="0 0 12 8" width="11" height="8" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1 1.5 6 6.5 11 1.5"/></svg>';
  const ICON_CHEVRON_RIGHT = '<svg viewBox="0 0 8 12" width="8" height="11" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M1.5 1 6.5 6 1.5 11"/></svg>';
  const ICON_TRASH = '<svg viewBox="0 0 14 14" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M2.5 4h9M5.5 4V2.5h3V4M4 4l.5 7h5L10 4M6 6v3.5M8 6v3.5"/></svg>';
  const ICON_FOLDER = '<svg viewBox="0 0 14 14" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M1.5 4.5A1.5 1.5 0 0 1 3 3h2.2l1.4 1.8h4.4A1.5 1.5 0 0 1 12.5 6.3V10a1.5 1.5 0 0 1-1.5 1.5H3A1.5 1.5 0 0 1 1.5 10z"/></svg>';
  const ICON_PLUS = '<svg viewBox="0 0 14 14" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M7 2.5v9M2.5 7h9"/></svg>';

  // 设置「当前聚焦 cell」高亮：所有 cell 移除 focused，仅目标 cell 加上
  function setFocusedCell(id) {
    cells.forEach((c) => {
      c.dom.root.classList.toggle("focused", c.id === id && !c.dom.cellBody.classList.contains("collapsed"));
    });
  }

  // ---------- cell 分隔按钮（「+ 查询」下沉为 cell 之间的分隔） ----------
  // notebook 内顺序约定：[divider, cell, divider, cell, ...]。每个 cell 前都有一个
  // 分隔按钮，点击它即在此位置插入新 cell。
  function makeDivider() {
    const d = document.createElement("div");
    d.className = "cell-divider";
    const lineL = document.createElement("span");
    lineL.className = "cell-divider-line";
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "cell-divider-btn";
    btn.title = "在此插入查询";
    btn.innerHTML = ICON_PLUS;
    btn.addEventListener("click", () => insertCellAt(d, ""));
    const lineR = document.createElement("span");
    lineR.className = "cell-divider-line";
    d.appendChild(lineL);
    d.appendChild(btn);
    d.appendChild(lineR);
    return d;
  }

  // 数 divider 之前的 cell 数量（此前已有 k 个 cell → 新 cell 数组索引 = k）
  function countCellsBefore(node) {
    let n = 0;
    let el = node.previousElementSibling;
    while (el) {
      if (el.classList.contains("cell")) n += 1;
      el = el.previousElementSibling;
    }
    return n;
  }

  // 在指定 divider 处插入一个新 cell（divider 前插入 [新divider, 新cell]）
  function insertCellAt(divider, sqlText) {
    return createCell(sqlText, false, undefined, divider);
  }

  // ---------- cell 序号徽标 + 大纲导航 ----------
  function refreshCellNumbers() {
    cells.forEach((c, i) => {
      if (c.dom.indexBadge) c.dom.indexBadge.textContent = String(i + 1);
    });
  }

  // cell 数量 >1 时显示右侧大纲；点序号跳转并聚焦对应 cell
  function refreshOutline() {
    const n = cells.length;
    if (n <= 1) {
      cellOutline.innerHTML = "";
      cellOutline.classList.add("hidden");
      return;
    }
    cellOutline.classList.remove("hidden");
    cellOutline.innerHTML = "";
    cells.forEach((c, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "outline-chip" + (c.id === activeCellId ? " active" : "");
      b.textContent = String(i + 1);
      b.title = `跳转到查询 ${i + 1}`;
      b.addEventListener("click", () => focusCell(c.id));
      cellOutline.appendChild(b);
    });
  }

  function focusCell(id) {
    const cell = getCell(id);
    if (!cell) return;
    expandCell(id);
    activeCellId = id;
    setFocusedCell(id);
    cell.dom.cm.focus();
    cell.dom.root.scrollIntoView({ block: "start", behavior: "smooth" });
    refreshOutline();
  }

  function createCell(sqlText, silent, initialLimit, anchorDivider) {
    const cell = {
      id: ++cellSeq,
      lastResult: null,
      fontScale: 1,
      autoFit: true,      // 高度自适应内容；用户手动拖拽调整后置 false
      lastAutoH: null,
      limit: initialLimit || DEFAULT_LIMIT,   // per-cell 行数限制
      dom: {},
    };
    // 注意：cell 的挂载与 cells 数组的插入统一在 mountCell 中完成，这里只构建。

    const root = document.createElement("div");
    root.className = "cell";
    root.dataset.cellId = cell.id;

    // 查询进度条（indeterminate，运行时显示）
    const progress = document.createElement("div");
    progress.className = "cell-progress hidden";
    const progressBar = document.createElement("div");
    progressBar.className = "cell-progress-bar";
    progress.appendChild(progressBar);

    // 头部：运行 + 状态 + 限制行数 + 折叠 + 删除（左右分组）
    const head = document.createElement("div");
    head.className = "cell-head";

    // 左组：序号徽标 + 运行 + 状态
    const headLeft = document.createElement("div");
    headLeft.className = "cell-head-left";

    const indexBadge = document.createElement("span");
    indexBadge.className = "cell-index";
    indexBadge.textContent = String(cells.length);   // 显示时暂用 cells 长度，挂载后在 refreshCellNumbers 里校正
    indexBadge.title = "查询编号";

    const runBtn = document.createElement("button");
    runBtn.type = "button";
    runBtn.className = "cell-run";
    runBtn.textContent = "▶ 运行";
    runBtn.addEventListener("click", () => runCell(cell.id));

    const status = document.createElement("span");
    status.className = "cell-status";

    headLeft.appendChild(indexBadge);
    headLeft.appendChild(runBtn);
    headLeft.appendChild(status);

    // 右组：限制行数 + 折叠 + 删除
    const headRight = document.createElement("div");
    headRight.className = "cell-head-right";

    // per-cell 行数限制输入框（独立于其他 cell，替换原全局「限制行数」）
    const limitInput = document.createElement("input");
    limitInput.type = "number";
    limitInput.min = "1";
    limitInput.max = String(MAX_RESULT_ROWS);
    limitInput.step = "1";
    limitInput.value = String(cell.limit);
    limitInput.className = "cell-limit";
    limitInput.title = `本查询最多返回行数（1~${MAX_RESULT_ROWS}，默认 ${DEFAULT_LIMIT}）`;
    limitInput.addEventListener("change", () => {
      const v = parseInt(limitInput.value, 10);
      cell.limit = (!Number.isFinite(v) || v <= 0) ? DEFAULT_LIMIT : Math.min(v, MAX_RESULT_ROWS);
      limitInput.value = String(cell.limit);
      persistSession();
    });

    // 折叠/展开按钮（仅收起 editor+result，头部保留）——图标化
    const toggleBtn = document.createElement("button");
    toggleBtn.type = "button";
    toggleBtn.className = "cell-icon cell-toggle";
    toggleBtn.title = "折叠/展开";
    toggleBtn.innerHTML = ICON_CHEVRON_DOWN;
    toggleBtn.addEventListener("click", () => toggleCell(cell.id));

    // 删除按钮——图标化（trash），hover 变红
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "cell-icon cell-del";
    delBtn.title = "删除此查询";
    delBtn.innerHTML = ICON_TRASH;
    delBtn.addEventListener("click", () => removeCell(cell.id));

    headRight.appendChild(limitInput);
    headRight.appendChild(toggleBtn);
    headRight.appendChild(delBtn);

    head.appendChild(headLeft);
    head.appendChild(headRight);

    // SQL 编辑器：CodeMirror（光标/文字/行号/高亮/滚动统一同坐标系）
    const editor = document.createElement("div");
    editor.className = "cell-editor";
    editor.dataset.cellId = cell.id;

    const cm = CodeMirror(editor, {
      mode: SQL_MODE_SPEC,
      theme: "excelsql",
      value: sqlText || defaultCellSql(),
      lineNumbers: true,
      lineWrapping: false,      // 长行横向滚动（行号列固定，永不覆盖文字）
      indentUnit: 4,
      tabSize: 4,
      indentWithTabs: false,    // Tab 插入空格（4 个），不插制表符
      viewportMargin: Infinity, // 全部行纳入渲染与高度度量，避免高分滚动下重排错位
      hintOptions: {
        hint: sqlHintFn,
        completeSingle: false,  // 单候选不自动顶替，补全只做主动/触发式
      },
      extraKeys: {
        "Ctrl-Enter": () => runCell(cell.id),
        "Cmd-Enter": () => runCell(cell.id),
        "Ctrl-/": "toggleComment",   // 注释/取消注释（行 -- / 块 /* */）
        "Cmd-/": "toggleComment",
        "Ctrl-=": () => zoomCell(cell, 1),
        "Ctrl-+": () => zoomCell(cell, 1),
        "Ctrl--": () => zoomCell(cell, -1),
        "Ctrl-0": () => zoomCell(cell, 0),
        "Ctrl-Space": "autocomplete",
      },
    });

    // 输入自动补全：敲字母/点号后，光标处于长度≥2 的标识符内时轻量触发（防抖）
    let hintTimer = null;
    cm.on("inputRead", (cmInstance, change) => {
      if (!change || !change.text || !change.text.length) return;
      if (!/[.\w]/.test(change.text[0])) return;
      clearTimeout(hintTimer);
      hintTimer = setTimeout(() => {
        if (cmInstance.state.completionActive) return;
        const cur = cmInstance.getCursor();
        const token = cmInstance.getTokenAt(cur);
        const word = token.string || "";
        if (/^[\w.]{2,}$/.test(word)) {
          CodeMirror.commands.autocomplete(cmInstance);
        }
      }, 120);
    });

    // 全角标点高亮层：所有 cell 统一挂同一个 overlay，纯视觉、
    // 不改内容。CodeMirror 每次 addOverlay 会重排一次，量级可忽略。
    cm.addOverlay(FULLWIDTH_PUNCT_OVERLAY);

    // 聚焦时标记为当前活动 cell（字段插入 / 运行的目标），并高亮头部 + 大纲
    cm.on("focus", () => { activeCellId = cell.id; setFocusedCell(cell.id); refreshOutline(); });
    // 内容变化后重新适配编辑器高度
    cm.on("change", () => autoSizeCell(cell));
    // 内容变化后防抖持久化，避免每次击键都写 localStorage
    cm.on("change", () => {
      clearTimeout(cm._persistTimer);
      cm._persistTimer = setTimeout(persistSession, 400);
    });
    // Ctrl + 滚轮缩放编辑器字号（阻止页面整体缩放）
    editor.addEventListener("wheel", (e) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        zoomCell(cell, e.deltaY < 0 ? 1 : -1);
      }
    }, { passive: false });

    // 结果区
    const result = document.createElement("div");
    result.className = "cell-result hidden";

    const resultHead = document.createElement("div");
    resultHead.className = "cell-result-head";
    const resultTitle = document.createElement("div");
    resultTitle.className = "step-title";
    const resultH2 = document.createElement("h2");
    resultH2.textContent = "查询结果";
    resultTitle.appendChild(resultH2);
    const rowCount = document.createElement("span");
    rowCount.className = "cell-row-count";
    const downloadBtn = document.createElement("button");
    downloadBtn.type = "button";
    downloadBtn.className = "cell-download";
    downloadBtn.textContent = "下载";
    downloadBtn.title = "导出为 Excel 或 CSV";
    downloadBtn.addEventListener("click", () => exportResult(cell.id));
    resultHead.appendChild(resultTitle);
    resultHead.appendChild(rowCount);
    resultHead.appendChild(downloadBtn);

    const resultWrap = document.createElement("div");
    resultWrap.className = "cell-result-wrap";
    const placeholder = document.createElement("p");
    placeholder.className = "placeholder";
    placeholder.textContent = "结果将显示在这里";
    resultWrap.appendChild(placeholder);

    result.appendChild(resultHead);
    result.appendChild(resultWrap);

    // 可折叠主体：editor + result 包一层，折叠时整体隐藏
    const cellBody = document.createElement("div");
    cellBody.className = "cell-body";
    cellBody.appendChild(editor);
    cellBody.appendChild(result);

    root.appendChild(progress);
    root.appendChild(head);
    root.appendChild(cellBody);

    cell.dom = { root, head, headLeft, headRight, indexBadge, runBtn, status, limitInput, toggleBtn, delBtn, cellBody, editor, cm, progress, progressBar, result, resultWrap, resultHead, rowCount, downloadBtn, placeholder };

    // 挂载：结构固定为 [cell, divider, cell, divider, …]——分隔按钮跟在每个 cell 之后
    // （cell 之间 + 末尾各一个），顶部第一个 cell 上方不放分隔按钮。
    if (anchorDivider && anchorDivider.parentNode === notebook) {
      // 点击某个分隔按钮 → 在它「之后」插入 [新 cell, 新分隔按钮]
      const idx = countCellsBefore(anchorDivider);   // 该 divider 之前的 cell 数 = 新 cell 索引
      cells.splice(idx, 0, cell);
      const next = anchorDivider.nextSibling;         // 后面的 cell 或 null
      notebook.insertBefore(root, next);              // 新 cell 插到 anchor 之后
      notebook.insertBefore(makeDivider(), next);     // 新分隔按钮紧跟新 cell 之后
    } else {
      // 初始 / 追加：cell 之后再挂一个分隔按钮
      cells.push(cell);
      notebook.appendChild(root);
      notebook.appendChild(makeDivider());
    }

    // 新创建的查询 cell 自动聚焦并滚入视野，便于立即输入 / 执行（恢复会话时跳过）
    activeCellId = cell.id;
    applyCellFont(cell);
    autoSizeCell(cell);
    refreshCellNumbers();
    refreshOutline();
    if (!silent) {
      cm.focus();
      root.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
    if (!silent) persistSession();
    return cell;
  }

  // 折叠/展开 cell 主体（editor + result）；展开后刷新 CodeMirror 防止布局错位
  function toggleCell(id) {
    const cell = getCell(id);
    if (!cell) return;
    const { cellBody, toggleBtn } = cell.dom;
    const collapsed = cellBody.classList.toggle("collapsed");
    toggleBtn.innerHTML = collapsed ? ICON_CHEVRON_RIGHT : ICON_CHEVRON_DOWN;
    if (!collapsed) {
      cell.dom.cm.refresh();
      autoSizeCell(cell);
    }
    // 折叠即退出聚焦高亮；展开则保持当前聚焦高亮
    setFocusedCell(activeCellId);
    persistSession();
  }

  function expandCell(id) {
    const cell = getCell(id);
    if (!cell) return;
    cell.dom.cellBody.classList.remove("collapsed");
    cell.dom.toggleBtn.innerHTML = ICON_CHEVRON_DOWN;
    cell.dom.cm.refresh();
    autoSizeCell(cell);
    setFocusedCell(activeCellId);
  }

  function removeCell(id) {
    const idx = cells.findIndex((c) => c.id === id);
    if (idx < 0) return;
    const cell = cells[idx];

    // 结构为 [cell, divider, cell, divider, …]：每个 cell 后面紧跟一个分隔按钮，
    // 删除 cell 时连带移除它后面的分隔按钮（若末尾无后置 divider 则不移除）。
    const nextDiv = cell.dom.root.nextElementSibling;
    if (nextDiv && nextDiv.classList.contains("cell-divider")) {
      nextDiv.remove();
    }

    cell.dom.root.remove();
    cells.splice(idx, 1);
    // 若没有 cell 了，补一个空 cell
    if (!cells.length) createCell("");
    else if (activeCellId === id) {
      activeCellId = cells[cells.length - 1].id;
      cells[cells.length - 1].dom.cm.focus();
    }
    refreshCellNumbers();
    refreshOutline();
    persistSession();
  }

  function getCell(id) {
    return cells.find((c) => c.id === id);
  }

  // ---------- 数据源 payload ----------
  function buildSourcesPayload() {
    if (mode === "single") {
      return [{ alias: "data", path: singlePath, sheet: sheetInput.value.trim() || null }];
    }
    return sources.map((s) => ({
      alias: s.alias,
      path: s.path,
      sheet: s.sheet || null,
    }));
  }

  // ---------- 进度条（一条连续、单调递增的填充条） ----------
  // 准备阶段（读文件/转码/嗅探，无真实进度）缓慢爬升到这个上限，
  // 之后由真实百分比接管，保证视觉连续、绝不回退。
  const PREP_CAP = 30;
  const PREP_DURATION = 1600; // 爬到 PREP_CAP 大约用时（ms）

  function startPrepare(cell) {
    const { progressBar, progress } = cell.dom;
    progress.classList.remove("hidden");
    progressBar.classList.remove("indeterminate");
    cell._phase = "prepare";
    cell._displayPct = 0;
    progressBar.style.width = "0%";

    const t0 = performance.now();
    cell._prepTimer = setInterval(() => {
      if (cell._phase !== "prepare") return; // 已进入 run，交给真实进度
      const t = Math.min(1, (performance.now() - t0) / PREP_DURATION);
      const eased = 1 - Math.pow(1 - t, 3);  // ease-out：先快后慢逼近
      const pct = PREP_CAP * eased;
      cell._displayPct = Math.max(cell._displayPct, pct);
      progressBar.style.width = cell._displayPct + "%";
    }, 40);
  }

  function stopPrepare(cell) {
    if (cell._prepTimer) {
      clearInterval(cell._prepTimer);
      cell._prepTimer = null;
    }
  }

  // 真实百分比：单调递增，绝不回退（DuckDB 进度可能从较小值起跳）
  function setRealProgress(cell, pct) {
    const { progressBar } = cell.dom;
    const v = Math.max(0, Math.min(100, Number(pct) || 0));
    if (cell._phase !== "run") {
      cell._phase = "run";
      stopPrepare(cell);
    }
    cell._displayPct = Math.max(cell._displayPct || 0, v);
    progressBar.classList.remove("indeterminate");
    progressBar.style.width = cell._displayPct + "%";
  }

  // ---------- 运行某个 cell ----------
  async function runCell(id) {
    const cell = getCell(id);
    if (!cell) return;
    const { cm, status, runBtn, progress } = cell.dom;

    const payloadSources = buildSourcesPayload();
    if (!payloadSources.length || payloadSources.some((s) => !s.path)) {
      status.textContent = "请先选择数据文件（或多源下补齐所有文件）";
      status.className = "cell-status error";
      return;
    }
    if (mode === "multi" && payloadSources.some((s) => !s.alias)) {
      status.textContent = "多源模式下每个源都需要填写别名";
      status.className = "cell-status error";
      return;
    }

    // 有选区时只执行选中的 SQL；否则执行整个编辑器内容
    const selection = cm.getSelection();
    const sql = (selection ? selection : cm.getValue()).trim();

    if (!sql) {
      status.textContent = "请输入 SQL（或先选中要执行的 SQL）";
      status.className = "cell-status error";
      return;
    }

    status.textContent = "准备数据…";
    status.className = "cell-status";
    runBtn.disabled = true;
    startPrepare(cell);         // 开始爬升准备进度条

    try {
      // per-cell 行数限制：运行时快照，避免运行期间被其他操作改动干扰
      const limit = cell.limit ?? DEFAULT_LIMIT;
      // 1) 提交任务，拿到 task_id
      const submit = await request("/api/query", {
        sources: payloadSources,
        sql,
        limit: Number.isFinite(limit) && limit > 0 ? limit : DEFAULT_LIMIT,
      });
      const taskId = submit.task_id;

      // 2) 轮询进度，直到 done / error（自适应退避：准备阶段慢轮询，运行阶段适度加快）
      let data = null;
      let pollDelay = 80;
      for (;;) {
        const p = await request("/api/progress", { task_id: taskId });
        if (p.phase === "run") {
          setRealProgress(cell, p.progress);
          if (status.textContent !== "查询中…") {
            status.textContent = "查询中…";
          }
        }
        if (p.status === "done") {
          data = p.result;
          // 多源查询成功：把当前源组合记入「最近组合」快照（供下次一键恢复）
          if (mode === "multi") recordCombo(sources);
          break;
        }
        if (p.status === "error") {
          throw new Error(p.error || "查询执行出错");
        }
        await new Promise((r) => setTimeout(r, pollDelay));
        // 线性退避到 500ms 上限之后不再增长，避免进度条更新显得迟滞
        pollDelay = Math.min(500, pollDelay + 40);
      }

      cell.lastResult = data;
      expandCell(cell.id);     // 先展开（解除 body 的 display:none），虚拟滚动才能测到容器高度
      renderCellResult(cell, data);
      // 结果完成后滚动右侧卡片（唯一纵向滚动容器）让结果区顶部对齐可视区，
      // 之后用户上滑即可回看 SQL。用 scrollBy 定位，避免滚到错误容器。
      setTimeout(() => {
        const scroller = notebook.parentElement;   // 即右侧 .card（滚动容器）
        const scrollerTop = scroller.getBoundingClientRect().top;
        const resTop = cell.dom.result.getBoundingClientRect().top;
        scroller.scrollBy({ top: resTop - scrollerTop, behavior: "smooth" });
      }, 60);
      // 优先展示「全流程总耗时」（含读文件/构建缓存），与用户体感一致；
      // 可选附准备段+SQL 段细分。elapsed_ms 仅为 SQL 执行段，单独回显会显得偏快。
      const fmt = (ms) => ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${Math.round(ms)} ms`;
      let timeText = "";
      if (typeof data.total_ms === "number") {
        timeText = `，耗时 ${fmt(data.total_ms)}`;
        if (typeof data.prepare_ms === "number" && data.prepare_ms > 0) {
          timeText += `（准备 ${fmt(data.prepare_ms)} + 查询 ${fmt(data.elapsed_ms ?? 0)}）`;
        }
      } else if (typeof data.elapsed_ms === "number") {
        timeText = `，耗时 ${fmt(data.elapsed_ms)}`;
      }
      const cacheText = data.cache === "built" ? "，已生成缓存"
        : data.cache === "hit" ? "，已用缓存"
        : "";
      status.textContent = `成功，返回 ${data.row_count} 行${timeText}${cacheText}`;
      status.className = "cell-status ok";
      setRealProgress(cell, 100);
    } catch (e) {
      status.textContent = popError(e);
      status.className = "cell-status error";
    } finally {
      runBtn.disabled = false;
      stopPrepare(cell);
      // 完成/出错后再显示一小段时间后隐藏进度条
      setTimeout(() => {
        if (progress) progress.classList.add("hidden");
      }, 350);
    }
  }

  // 结果表虚拟滚动：行数超过阈值时只渲染可视区的行，避免 1 万行 DOM 卡顿。
  // 实现：单表结构保持 sticky 表头不变，tbody 顶部/底部各放一行等高 spacer
  // （内嵌一个 colSpan=全部列的 td），把总滚动高度撑到「总行数 × 行高」；
  // 滚动时只重建可视窗口内的实际数据行，因此 DOM 节点数恒定为「可视行数」。
  const VSCROLL_ROW_PX = 32;          // 与 .vscroll td 的固定行高一致
  const VSCROLL_BUFFER = 6;           // 上下各多渲染几行，减少滚动白边
  const VSCROLL_THRESHOLD = 30;       // 超过此行数才启用虚拟滚动

  function renderCellResult(cell, data) {
    const cols = data.columns || [];
    const rows = data.rows || [];
    const { result, resultWrap, rowCount } = cell.dom;

    // 清空旧结果，重建表格
    resultWrap.innerHTML = "";
    // 先解除隐藏（display:none 时 clientHeight=0，虚拟滚动的 paint() 会算不出
    // 可视行数，导致首屏只渲染前几行 + 大 spacer = 一半空白），再渲染表格。
    result.classList.remove("hidden");

    const virtual = rows.length > VSCROLL_THRESHOLD;
    resultWrap.classList.toggle("vscroll", virtual);
    if (virtual) {
      renderVirtualTable(cell, cols, rows);
    } else {
      resultWrap.appendChild(buildFullTable(cols, rows, 0, rows.length));
    }

    const effectiveLimit = data.limit;
    const truncated = typeof effectiveLimit === "number" && data.row_count >= effectiveLimit;
    rowCount.textContent = `${rows.length} 行${truncated ? `（已达 ${effectiveLimit} 行上限，可能截断）` : ""}`;
  }

  // 把「数字或可转数字的字符串」转成 number，转不了返回 NaN。
  // 兼容混合类型文件全文本读出的数字字符串（"123.45"、"1,234"——千分位逗号
  // 与原文件里的空格会被剥离）；"123abc"、"0x10" 这类不视为数字。
  const DECIMAL_RE = /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/;
  function toNumber(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
    if (typeof v !== "string") return NaN;
    const s = v.trim().replace(/,/g, "");
    if (s === "") return NaN;
    return DECIMAL_RE.test(s) ? Number(s) : NaN;
  }

  // 计算各列合计：某列至少有一个可转数字的值时返回其 SUM（非数字单元格
  // 忽略不计，与 Excel SUM 行为一致，容忍一列里混个别文本/空值）；
  // 一列完全没有数字时返回 null。只统计当前已返回的行。
  function computeColumnTotals(cols, rows) {
    if (!rows.length) return null;
    const sums = new Array(cols.length).fill(null);
    for (let ci = 0; ci < cols.length; ci++) {
      let sum = 0;
      let any = false;
      for (let i = 0; i < rows.length; i++) {
        const n = toNumber(rows[i][cols[ci]]);
        if (Number.isNaN(n)) continue;   // 空值/文本：跳过，不影响求和
        sum += n;
        any = true;
      }
      if (any) sums[ci] = sum;
    }
    return sums;
  }

  // 合计值格式化：整数原样输出，浮点去掉误差尾巴（0.30000000000000004 → 0.3）
  function formatTotal(v) {
    if (typeof v !== "number" || !Number.isFinite(v)) return "—";
    if (Number.isInteger(v)) return String(v);
    return String(parseFloat(v.toFixed(10)));
  }

  // 构建完整 table（表头 + from..to 之间数据行 + 底部合计行）；
  // to<=from 时只建空 tbody（虚拟滚动模式由 paint() 填充）。
  function buildFullTable(cols, rows, from, to) {
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const trHead = document.createElement("tr");
    cols.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c;
      trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let i = from; i < to; i++) {
      tbody.appendChild(buildRow(cols, rows[i]));
    }
    table.appendChild(tbody);

    // 合计行：纯数字列显示 SUM，首列单元格放「合计」标签（首列也是数字列时
    // 显示「合计 <sum>」；其余非数字列显示「—」）。无数据时不输出合计行。
    const totals = computeColumnTotals(cols, rows);
    if (totals) {
      const tfoot = document.createElement("tfoot");
      const tr = document.createElement("tr");
      tr.className = "totals-row";
      cols.forEach((c, ci) => {
        const td = document.createElement("td");
        if (ci === 0) {
          td.textContent =
            totals[0] !== null ? `合计 ${formatTotal(totals[0])}` : "合计";
          td.className = "totals-label";
        } else {
          td.textContent = totals[ci] !== null ? formatTotal(totals[ci]) : "—";
        }
        tr.appendChild(td);
      });
      tfoot.appendChild(tr);
      table.appendChild(tfoot);
    }
    return table;
  }

  function buildRow(cols, r) {
    const tr = document.createElement("tr");
    cols.forEach((c) => {
      const td = document.createElement("td");
      const v = r[c];
      td.textContent = v === null || v === undefined ? "" : String(v);
      tr.appendChild(td);
    });
    return tr;
  }

  // 在 tbody 里放一个整行高为 px 的空行，用于撑出滚动高度
  function buildSpacerRow(cols, px) {
    const tr = document.createElement("tr");
    tr.className = "vscroll-spacer-row";
    const td = document.createElement("td");
    td.colSpan = cols.length;
    td.style.height = px + "px";
    td.style.padding = "0";
    td.style.border = "none";
    tr.appendChild(td);
    return tr;
  }

  function renderVirtualTable(cell, cols, rows) {
    const total = rows.length;
    const wrap = cell.dom.resultWrap;

    // tfoot 只建一次（含合计行），滚动重建 tbody 时不动它
    const table = buildFullTable(cols, rows, 0, 0);  // 表头 + 空 tbody
    wrap.appendChild(table);
    const tbody = table.querySelector("tbody");

    function paint() {
      const scrollTop = wrap.scrollTop;
      const first = Math.max(0, Math.floor(scrollTop / VSCROLL_ROW_PX) - VSCROLL_BUFFER);
      const last = Math.min(total, Math.ceil((scrollTop + wrap.clientHeight) / VSCROLL_ROW_PX) + VSCROLL_BUFFER);

      // 重建 tbody：顶部 spacer + 数据行 + 底部 spacer
      const frag = document.createDocumentFragment();
      if (first > 0) frag.appendChild(buildSpacerRow(cols, first * VSCROLL_ROW_PX));
      for (let i = first; i < last; i++) frag.appendChild(buildRow(cols, rows[i]));
      if (last < total) frag.appendChild(buildSpacerRow(cols, (total - last) * VSCROLL_ROW_PX));
      tbody.innerHTML = "";
      tbody.appendChild(frag);
    }
    if (wrap._vscrollHandler) wrap.removeEventListener("scroll", wrap._vscrollHandler);
    wrap._vscrollHandler = () => {
      if (!wrap._paintRaf) {
        wrap._paintRaf = requestAnimationFrame(() => {
          wrap._paintRaf = null;
          paint();
        });
      }
    };
    wrap.addEventListener("scroll", wrap._vscrollHandler);
    paint();
  }

  async function exportResult(id) {
    const cell = getCell(id);
    if (!cell || !cell.lastResult || !cell.lastResult.columns?.length) {
      if (cell) {
        cell.dom.status.textContent = "没有可导出的结果";
        cell.dom.status.className = "cell-status error";
      }
      return;
    }
    const cols = cell.lastResult.columns;
    const rows = cell.lastResult.rows || [];
    const filename = "query_result_" + new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");

    const cellStatus = cell.dom.status;
    const prevText = cellStatus.textContent;
    const prevClass = cellStatus.className;
    cellStatus.textContent = "正在导出…";
    cellStatus.className = "cell-status";

    try {
      const res = await request("/api/export", {
        columns: cols,
        rows,
        filename,
      });
      if (!res || res.cancelled) {
        cellStatus.textContent = prevText;
        cellStatus.className = prevClass;
        return;
      }
      cellStatus.textContent = `已导出 ${res.rows ?? rows.length} 行 → ${res.path}`;
      cellStatus.className = "cell-status ok";
    } catch (e) {
      cellStatus.textContent = popError(e);
      cellStatus.className = "cell-status error";
    }
  }

  // ---------- 事件绑定 ----------
  modeSingleBtn.addEventListener("click", () => setMode("single"));
  modeMultiBtn.addEventListener("click", () => setMode("multi"));
  addSourceBtn.addEventListener("click", () => {
    const n = sources.length + 1;
    createSource(`t${n}`, "", "");
    renderSources();
  });
  pickBtn.addEventListener("click", pickFile);
  clearSingleBtn.addEventListener("click", clearSingle);

  // 最近打开浮层：按钮切换，点击外部 / Esc 关闭
  recentBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    if (recentOpen && recentAnchorEl === recentBtn) {
      closeRecentPop();
    } else {
      recentTargetSrc = null;          // 全局按钮打开：清掉源卡目标
      toggleRecentPop(true, recentBtn);
    }
  });
  document.addEventListener("click", closeRecentPopOnOutside);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && recentOpen) closeRecentPop();
  });
  fieldSearch.addEventListener("input", () => {
    if (mode === "single") {
      renderFields(getSingleColumns());
    } else {
      const active = sources.find((s) => s.id === activeSourceId);
      renderFields(active ? active.columns : []);
    }
  });
  sheetInput.addEventListener("change", () => {
    if (mode === "single" && singlePath) openSingle();
  });

  // ---------- 大纲浮轨：滚动时显示，停止后自动隐藏 ----------
  // 滚动容器是右侧 notebook 所在的卡片（唯一的纵向滚动容器）。
  (function initOutlineReveal() {
    const scroller = notebook.parentElement;   // .panel.grow .card:last-child
    let hideTimer = null;
    let hovering = false;

    function scheduleHide() {
      clearTimeout(hideTimer);
      hideTimer = setTimeout(() => {
        if (!hovering) cellOutline.classList.remove("visible");
      }, 800);
    }

    function revealOutline() {
      if (cellOutline.classList.contains("hidden")) return;   // cell<=1 时无大纲
      cellOutline.classList.add("visible");
      scheduleHide();
    }

    if (scroller) {
      scroller.addEventListener("scroll", revealOutline, { passive: true });
    }
    // 点击大纲时保持显示（用户在操作导航，不应立刻淡出）
    cellOutline.addEventListener("click", revealOutline);

    // 鼠标悬停在大纲上时保持显示、不自动隐藏；移出后恢复定时隐藏
    cellOutline.addEventListener("mouseenter", () => {
      if (cellOutline.classList.contains("hidden")) return;
      hovering = true;
      clearTimeout(hideTimer);
      cellOutline.classList.add("visible");
    });
    cellOutline.addEventListener("mouseleave", () => {
      hovering = false;
      scheduleHide();
    });
  })();

  // ---------- 左栏可拖拽分栏 ----------
  // 拖动 .splitter 调整左栏宽度（180~420px），双击复位到默认 210px。
  const LAYOUT_GUTTER = 10;      // 与 CSS .splitter 的 flex-basis 一致
  const LEFT_MIN = 180;
  const LEFT_MAX = 420;
  const LEFT_DEFAULT = 210;

  function setLeftWidth(w) {
    leftPanel.style.width = Math.round(Math.max(LEFT_MIN, Math.min(LEFT_MAX, w))) + "px";
  }

  splitter.addEventListener("pointerdown", (e) => {
    e.preventDefault();
    splitter.setPointerCapture(e.pointerId);
    splitter.classList.add("dragging");
    document.body.style.userSelect = "none";
    document.body.style.cursor = "col-resize";

    const onMove = (ev) => {
      // 用左栏右边缘到指针的距离 + 半个分栏器，换算成左栏新宽度
      const rect = leftPanel.getBoundingClientRect();
      const newW = ev.clientX - rect.left + LAYOUT_GUTTER / 2;
      setLeftWidth(newW);
    };
    const onUp = () => {
      splitter.classList.remove("dragging");
      document.body.style.userSelect = "";
      document.body.style.cursor = "";
      splitter.removeEventListener("pointermove", onMove);
      splitter.removeEventListener("pointerup", onUp);
      splitter.removeEventListener("pointercancel", onUp);
      // 拖拽改变宽度后刷新 CodeMirror 布局（编辑器随容器宽度变化）
      cells.forEach((c) => c.dom.cm && c.dom.cm.refresh());
    };

    splitter.addEventListener("pointermove", onMove);
    splitter.addEventListener("pointerup", onUp);
    splitter.addEventListener("pointercancel", onUp);
  });

  splitter.addEventListener("dblclick", () => setLeftWidth(LEFT_DEFAULT));

  // ---------- 窄窗口：侧栏抽屉 ----------
  function setDrawer(open) {
    layoutEl.classList.toggle("drawer-open", open);
  }
  sidebarToggle.addEventListener("click", () => {
    setDrawer(!layoutEl.classList.contains("drawer-open"));
  });
  // 点击抽屉遮罩（::before）即关闭：通过判断点击目标是否为遮罩本体
  layoutEl.addEventListener("click", (e) => {
    if (layoutEl.classList.contains("drawer-open") && e.target === layoutEl) {
      setDrawer(false);
    }
  });

  // 初始状态
  _singleColumns = [];
  renderSingle();
  renderRecent();   // 渲染「最近打开」历史列表（无历史时整区隐藏）
  // 会话恢复：优先用上次保存的 SQL / 模式 / 折叠状态；无则建默认 cell
  const saved = restoreSession();
  if (saved && Array.isArray(saved.cells)) {
    const targetMode = saved.mode === "multi" ? "multi" : "single";
    if (targetMode !== mode) setMode(targetMode);   // setMode 会做完整的 UI 切换
    const cellSpecs = saved.cells;
    if (cellSpecs.length) {
      cellSpecs.forEach((cs) => {
        const c = createCell(cs.sql || "", true, cs.limit);
        if (cs.collapsed && c.dom.cellBody) {
          c.dom.cellBody.classList.add("collapsed");
          c.dom.toggleBtn.innerHTML = ICON_CHEVRON_RIGHT;
        }
      });
    } else {
      createCell("SELECT * FROM data LIMIT 10;", true);
    }
    persistSession();  // 固化本次恢复后的最终状态（覆盖 setMode 中途写入的空 cells）
  } else {
    // 初始建立一个查询 cell
    createCell("SELECT * FROM data LIMIT 10;");
  }

  // 等宽字体加载完成后重新校准高度（避免按回退字体度量测量）
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(() => cells.forEach(autoSizeCell));
  }
})();