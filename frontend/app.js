// Excel SQL 查询台 —— 前端逻辑（单文件 / 多源关联 两种模式）
(() => {
  "use strict";

  const API_BASE = ""; // 同源部署；如前后端分离可改为 http://127.0.0.1:8000

  // ---------- 后端 SQLite 存储（替代 localStorage） ----------
  // 所有持久化数据统一走后端 SQLite（%LOCALAPPDATA%\ExcelSqlConsole\app.sqlite），
  // 摆脱 WebView2 profile/localStorage 的绑定（清缓存/换 profile 不丢、可备份）。
  // storeSet 静默写（不阻塞界面）；storeGet 批量读。
  function storeSet(kvs) {
    fetch(API_BASE + "/api/store/set", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ kvs }),   // kvs: {key: JSON字符串 | null}，null 表示删除
    }).catch(() => { /* 后端暂不可用时静默，下次写入重试 */ });
  }
  async function storeGet(keys) {
    try {
      const res = await request("/api/store/get", { keys });
      return (res && res.values) || {};
    } catch (_) {
      return {};
    }
  }

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

  // ---------- 会话持久化（后端 SQLite） ----------
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
      storeSet({ [STORE_KEY]: JSON.stringify(payload) });
    } catch (_) { /* 序列化失败时静默降级 */ }
  }
  function restoreSession(raw) {
    // raw 为已读到的 JSON 字符串（由 bootstrap 统一传入）；失败返回 null
    try {
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
  let recentFiles = [];       // 由 bootstrap 从后端 SQLite 载入
  let combos = [];            // 同上
  let histSeq = 0;            // 历史条目唯一 id 自增（支持「另存为」同文件多条命名条目）
  let comboSeq = 0;           // 组合唯一 id 自增（用于「另存为」产生多条同指纹条目）
  let recentOpen = false;      // 浮层当前是否展开
  let recentAnchorEl = null;   // 当前浮层的锚点元素（最近打开按钮 / 源卡「选择」按钮）
  let recentTargetSrc = null;  // 浮层选中后的目标源（多源模式源卡打开时；null=单文件）
  let missingPaths = {};       // path -> true（文件已被移除，用于标红）
  let checkingMissing = false; // 防止并发检查

  function loadCombos(raw) {
    try {
      const arr = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(arr)) return [];
      let maxId = 0;
      const out = [];
      arr.forEach((c) => {
        if (!c || !Array.isArray(c.sources) || !c.sources.length) return;
        if (!c.id) c.id = ++comboSeq;          // 旧版组合没有 id，补发
        if (c.id > maxId) maxId = c.id;
        out.push(c);
      });
      if (maxId > comboSeq) comboSeq = maxId;  // 与既有 id 不冲突
      return out;
    } catch (_) { return []; }
  }

  // 记录一次多源组合：查询成功后调用，同指纹条目顶到最前，超限截断。
  // name 可选：给组合起名字（含 name 的保存决定名称；已有名字不覆盖）。
  // 返回 { fp } —— fp 即该组合的来源指纹（供对话框定位「当前组合」）。
  function recordCombo(sourcesArr, sql, name) {
    if (!sourcesArr || !sourcesArr.length) return { fp: "" };
    const snap = sourcesArr
      .filter((s) => s && (s.path || (s.kind === "doris" && s.conn && s.conn.host)))
      .map((s) => {
        if (s.kind === "doris") {
          return {
            kind: "doris",
            alias: s.alias || "",
            conn: s.conn || null,
            db: s.db || "",
            table: s.table || "",
            path: "",
            sheet: "",
          };
        }
        return { path: s.path, sheet: s.sheet || "", alias: s.alias || "" };
      });
    if (!snap.length) return { fp: "" };
    // 组合指纹：路径（doris 用 别名@host:port，连接级源可能与 db/table 无关）序列
    // （稳定排序，忽略顺序差异）
    const keyOf = (s) => s.kind === "doris"
      ? `doris:${s.alias}@${s.conn ? s.conn.host + ":" + s.conn.port : ""}`
      : s.path;
    const fp = snap.map(keyOf).sort().join("|");
    const existing = combos.find((c) => c.fp === fp);
    if (existing) {
      // 指纹重复：条目顶到最前，刷新时间戳与 SQL（保留已有名字，除非本次显式带新名字且无旧名）
      combos = combos.filter((c) => c.id !== existing.id);
      if (!existing.name && name) existing.name = name;
      existing.time = Date.now();
      existing.sql = sql || "";
      combos.unshift(existing);
      storeSet({ [COMBOS_KEY]: JSON.stringify(combos) });
      return { fp };
    }
    combos.unshift({
      id: ++comboSeq,
      fp,
      name: name || "",
      sources: snap,
      time: Date.now(),
      sql: sql || "",
    });
    if (combos.length > COMBO_MAX) combos.length = COMBO_MAX;
    storeSet({ [COMBOS_KEY]: JSON.stringify(combos) });
    return { fp };
  }

  // 另存为：无条件产生一条新组合（新 id、可与原组合同指纹共存），不影响既有条目。
  // 返回新条目的 { fp, id }。
  function recordComboAs(sourcesArr, sql, name) {
    if (!sourcesArr || !sourcesArr.length) return { fp: "", id: 0 };
    const snap = sourcesArr
      .filter((s) => s && (s.path || (s.kind === "doris" && s.conn && s.conn.host)))
      .map((s) => {
        if (s.kind === "doris") {
          return { kind: "doris", alias: s.alias || "", conn: s.conn || null, db: s.db || "", table: s.table || "", path: "", sheet: "" };
        }
        return { path: s.path, sheet: s.sheet || "", alias: s.alias || "" };
      });
    if (!snap.length) return { fp: "", id: 0 };
    const keyOf = (s) => s.kind === "doris"
      ? `doris:${s.alias}@${s.conn ? s.conn.host + ":" + s.conn.port : ""}`
      : s.path;
    const fp = snap.map(keyOf).sort().join("|");
    const entry = { id: ++comboSeq, fp, name: name || "", sources: snap, time: Date.now(), sql: sql || "" };
    combos.unshift(entry);
    if (combos.length > COMBO_MAX) combos.length = COMBO_MAX;
    storeSet({ [COMBOS_KEY]: JSON.stringify(combos) });
    return { fp, id: entry.id };
  }

  // 删除一个已保存的组合（按唯一 id），持久化并重绘当前弹框。
  function removeCombo(id) {
    combos = combos.filter((c) => c.id !== id);
    storeSet({ [COMBOS_KEY]: JSON.stringify(combos) });
    if (recentOpen) renderRecentPop();
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

  function loadHistory(raw) {
    try {
      const arr = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(arr)) return [];
      let maxId = 0;
      const out = [];
      arr.forEach((e) => {
        if (!e || typeof e.path !== "string") return;
        if (!e.id) e.id = ++histSeq;              // 旧版历史没有 id，补发
        if (e.id > maxId) maxId = e.id;
        if (typeof e.name !== "string") e.name = ""; // 旧版无 name 字段
        out.push(e);
      });
      if (maxId > histSeq) histSeq = maxId;       // 与既有 id 不冲突
      return out;
    } catch (_) { return []; }
  }

  // 记录一次打开：同路径的最前一条条目顶到最前并刷新，超限截断，再重绘。
  // （语义与组合「按指纹去重」一致：自动记录不会新建重复条目，命名条目保留名称）
  // id 可选：精确更新指定条目（从历史列表点开旧命名条目时用）。
  function recordHistory(path, sheet, name, id) {
    if (!path) return;
    const now = Date.now();
    const existing = recentFiles.find((e) => (id ? e.id === id : e.path === path));
    if (existing) {
      // 已有条目：顶到最前，刷新时间与工作表（保留已有名字，除非显式给出）
      recentFiles = recentFiles.filter((e) => e.id !== existing.id);
      if (name !== undefined && typeof name === "string") existing.name = name;
      existing.sheet = sheet || "";
      existing.time = now;
      recentFiles.unshift(existing);
    } else {
      // 无对应条目：新建一条
      recentFiles.unshift({ id: ++histSeq, path, sheet: sheet || "", name: name || "", time: now });
    }
    if (recentFiles.length > HIST_MAX) recentFiles.length = HIST_MAX;
    storeSet({ [HIST_KEY]: JSON.stringify(recentFiles) });
    renderRecent();
  }

  // 查询成功后：把本次执行的 SQL 记到该文件的历史条目（供下次打开时询问是否回填）。
  function recordHistorySql(path, sql) {
    if (!path || !sql) return;
    const e = recentFiles.find((x) => x.path === path);
    if (!e) return;
    e.sql = sql;
    storeSet({ [HIST_KEY]: JSON.stringify(recentFiles) });
  }

  // 从历史移除一条（文件被删/手动删除），按唯一 id
  function removeHistory(idOrPath) {
    if (typeof idOrPath === "number") {
      recentFiles = recentFiles.filter((e) => e.id !== idOrPath);
    } else {
      // 兼容旧调用（按 path 删除所有该路径条目）
      recentFiles = recentFiles.filter((e) => e.path !== idOrPath);
    }
    storeSet({ [HIST_KEY]: JSON.stringify(recentFiles) });
    renderRecent();
  }

  function clearHistory() {
    recentFiles = [];
    storeSet({ [HIST_KEY]: null });
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

  // 渲染「最近打开」：图标按钮显示条数角标；有历史才显示按钮，浮层内容 → recentPop。
  function renderRecent() {
    const n = recentFiles.length;
    recentBtn.classList.toggle("hidden", !n);
    recentCount.textContent = n ? String(n) : "";
    if (!n && recentOpen) closeRecentPop();
    else if (recentOpen) renderRecentPop();
  }

  // 「主动保存」按钮显隐：常驻显示（无论有无内容都可见）。
  function renderSaveBtn() {
    if (saveBtn) saveBtn.classList.remove("hidden");
  }

  // 主动保存当前工作状态：
  //   多源 → 弹「保存组合」对话框（改名称 / 保存当前组合 / 另存为）；
  //   单文件 → 弹「保存文件」对话框（改名称 / 保存当前文件 / 另存为），与组合一致。
  function saveCurrentState() {
    if (mode === "multi") {
      if (!sources.some((s) => s.path || (s.kind === "doris" && s.conn))) {
        showConfirm("当前没有可保存的数据源组合。", null);
        return;
      }
    } else if (!singlePath) {
      showConfirm("请先选择一个文件。", null);
      return;
    }
    openSaveComboPop();
  }

  // 取当前活动 cell（无活动时取最后一个 cell）里的 SQL 文本。
  function getActiveSql() {
    const cell = getCell(activeCellId) || cells[cells.length - 1];
    return cell && cell.dom.cm ? cell.dom.cm.getValue() : "";
  }

  // ────────── 保存对话框（保存当前 / 另存为 / 重命名；多源=组合，单源=文件）──────────
  // 计算多源指纹（无副作用，仅用于定位既有组合）
  function comboFpOf(sourcesArr) {
    if (!sourcesArr || !sourcesArr.length) return "";
    const snap = sourcesArr
      .filter((s) => s && (s.path || (s.kind === "doris" && s.conn && s.conn.host)));
    if (!snap.length) return "";
    const keyOf = (s) => s.kind === "doris"
      ? `doris:${s.alias}@${s.conn ? s.conn.host + ":" + s.conn.port : ""}`
      : s.path;
    return snap.map(keyOf).sort().join("|");
  }

  let saveDlgMode = "multi";   // "multi" | "single"：当前保存对话框作用的目标类型

  // 打开保存对话框：按当前模式定位既有条目，预填名称。
  function openSaveComboPop() {
    if (recentOpen) closeRecentPop();   // 两个模态同层，先收起「最近打开」
    saveDlgMode = mode === "multi" ? "multi" : "single";
    let existing = null;
    if (saveDlgMode === "multi") {
      const fp = comboFpOf(sources);
      existing = combos.find((c) => c.fp === fp);
    } else {
      // 单源：定位该路径最近的一条历史条目（与组合按指纹定位语义一致），预填其名称
      existing = singlePath ? recentFiles.find((e) => e.path === singlePath) : null;
    }
    const name = existing ? existing.name || "" : "";
    const isFile = saveDlgMode === "single";
    const titleEl = document.getElementById("saveComboTitle");
    if (titleEl) titleEl.textContent = isFile ? "保存文件" : "保存组合";
    const nameLabel = document.getElementById("saveComboNameLabel");
    if (nameLabel) nameLabel.textContent = isFile ? "文件名称" : "组合名称";
    const curBtn = document.getElementById("saveComboCurrent");
    if (curBtn) curBtn.textContent = isFile ? "保存当前文件" : "保存当前组合";
    if (saveComboNameInput) saveComboNameInput.value = name;
    if (saveComboTip) {
      saveComboTip.textContent = existing
        ? (isFile
            ? "「保存当前文件」将更新它的时间与 SQL；也可以在下方修改名称"
            : "「保存当前组合」将更新它的时间与 SQL；也可以在下方修改组合名称")
        : (isFile
            ? "为这个文件起个名字，方便在「最近打开」里快速找回"
            : "为这个组合起个名字，方便之后快速找回");
    }
    saveComboPop.classList.remove("hidden");
    if (saveComboNameInput) saveComboNameInput.focus();
  }

  function closeSaveComboPop() {
    saveComboPop.classList.add("hidden");
  }

  // 多源「保存当前组合」：已有条目 → 顶到最前并刷新时间/SQL/名称；没有 → 新建。
  function saveComboCurrentMulti(name, sql) {
    const fp = comboFpOf(sources);
    const existing = combos.find((c) => c.fp === fp);
    if (existing) {
      // 原子更新：重命名 + 刷新时间/SQL（不要先改旧引用，会丢名称）
      const updated = { ...existing, name, time: Date.now(), sql: sql || "" };
      combos = [updated, ...combos.filter((c) => c.id !== existing.id)];
      if (combos.length > COMBO_MAX) combos.length = COMBO_MAX;
      storeSet({ [COMBOS_KEY]: JSON.stringify(combos) });
      if (recentOpen) renderRecentPop();
    } else {
      recordCombo(sources, sql, name || null);
    }
    closeSaveComboPop();
    showStatus(uploadStatus, "已保存组合", "ok");
  }

  // 单源「保存当前文件」：更新该路径最近的历史条目（顶到最前），没有则新建。
  function saveComboCurrentSingle(name, sql) {
    if (!singlePath) { closeSaveComboPop(); return; }
    const sheet = sheetInput.value.trim() || "";
    const existing = recentFiles.find((e) => e.path === singlePath);
    if (existing) {
      const updated = { ...existing, name, sheet, time: Date.now(), sql: sql || "" };
      recentFiles = [updated, ...recentFiles.filter((e) => e.id !== existing.id)];
      if (recentFiles.length > HIST_MAX) recentFiles.length = HIST_MAX;
      storeSet({ [HIST_KEY]: JSON.stringify(recentFiles) });
    } else {
      recentFiles.unshift({
        id: ++histSeq, path: singlePath, sheet, name, time: Date.now(), sql: sql || "",
      });
      if (recentFiles.length > HIST_MAX) recentFiles.length = HIST_MAX;
      storeSet({ [HIST_KEY]: JSON.stringify(recentFiles) });
    }
    renderRecent();
    closeSaveComboPop();
    showStatus(uploadStatus, "已保存文件", "ok");
  }

  // 「保存当前」：按模式分流。
  function saveComboCurrent() {
    const name = saveComboNameInput.value.trim();
    const sql = getActiveSql();
    if (saveDlgMode === "multi") {
      saveComboCurrentMulti(name, sql);
    } else {
      saveComboCurrentSingle(name, sql);
    }
  }

  // 「另存为」：无条件产生一条新条目（新 id，可与原条目并存），不影响既有条目。
  function saveComboAsCopy() {
    const name = saveComboNameInput.value.trim() || "未命名";
    if (saveDlgMode === "multi") {
      const res = recordComboAs(sources, getActiveSql(), name);
      closeSaveComboPop();
      showStatus(uploadStatus, res.id ? `已另存为「${name}」` : "已保存组合", "ok");
      return;
    }
    if (!singlePath) { closeSaveComboPop(); return; }
    recentFiles.unshift({
      id: ++histSeq, path: singlePath,
      sheet: sheetInput.value.trim() || "",
      name, time: Date.now(), sql: getActiveSql() || "",
    });
    if (recentFiles.length > HIST_MAX) recentFiles.length = HIST_MAX;
    storeSet({ [HIST_KEY]: JSON.stringify(recentFiles) });
    renderRecent();
    closeSaveComboPop();
    showStatus(uploadStatus, `已另存为「${name}」`, "ok");
  }

  // 重建弹框内容（组合区 + 列表 + 更多），整体由 JS 构建到 recentPopBody 容器。
  function renderRecentPop() {
    const box = recentPopBody;
    box.innerHTML = "";
    // 弹框标题：源卡打开时称「选择历史文件」，全局打开称「最近打开」
    const titleEl = recentPop.querySelector(".recent-title");
    if (titleEl) titleEl.textContent = recentTargetSrc ? "选择历史文件" : "最近打开";

    // 「最近组合」只在全局「最近打开」弹框展示（多源模式）：
    // 源卡文件夹图标弹框是「选择历史文件」语义（给单个源换文件），
    // 混入会清空全部源的「恢复组合」会误操作，故排除。
    const isMultiPopup = !recentTargetSrc && mode === "multi";

    // 多源模式下、全局打开时：顶部展示「最近使用过的组合」区
    if (isMultiPopup) {
      if (combos.length) {
        const comboHead = document.createElement("div");
        comboHead.className = "recent-combo-head";
        comboHead.textContent = "最近组合";
        box.appendChild(comboHead);

        const comboList = document.createElement("div");
        comboList.className = "recent-combo-list";
        combos.forEach((combo) => {
          const item = document.createElement("div");
          item.className = "recent-combo-item";
          item.title = combo.sources.map((s) =>
            s.kind === "doris" ? `doris: ${s.db}.${s.table}` : s.path
          ).join("\n");

          // 摘要：组合名 + 源个数，若没有名字则用源摘要
          const names = combo.sources.map((s) =>
            s.kind === "doris" ? `🔗 ${s.db}.${s.table}` : basename(s.path)
          );
          const label = names.length <= 3 ? names.join(" · ") : `${names[0]} … +${names.length - 1} 个`;
          const nameEl = document.createElement("span");
          nameEl.className = "recent-combo-name";
          // 有名字 → 显示名字（加粗），副标题显示源摘要；无名字 → 直接显示源摘要
          if (combo.name) {
            nameEl.innerHTML = `<b>${escapeHtml(combo.name)}</b> <i class="recent-combo-label">${escapeHtml(label)}</i>`;
          } else {
            nameEl.textContent = label;
          }

          const timeEl = document.createElement("span");
          timeEl.className = "recent-time";
          timeEl.textContent = historyTime(combo.time);

          // 组合删除按钮（按唯一 id）
          const delBtn = document.createElement("button");
          delBtn.type = "button";
          delBtn.className = "recent-del";
          delBtn.textContent = "×";
          delBtn.title = "删除该组合";
          delBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            removeCombo(combo.id);
          });

          item.appendChild(nameEl);
          item.appendChild(timeEl);
          item.appendChild(delBtn);

          // 点击恢复组合
          item.addEventListener("click", (e) => {
            e.stopPropagation();
            closeRecentPop();
            restoreCombo(combo);
          });
          comboList.appendChild(item);
        });
        box.appendChild(comboList);
      } else {
        // 空态引导：让用户知道何时会出现组合
        const hint = document.createElement("div");
        hint.className = "recent-combo-hint";
        hint.textContent = "多源执行一次查询后，这里会自动记住你的源组合";
        box.appendChild(hint);
      }
    }

    // 列表区标题 + 清空按钮
    const head = document.createElement("div");
    head.className = "recent-head";
    const subTitle = document.createElement("span");
    subTitle.className = "recent-subtitle";
    subTitle.textContent = recentTargetSrc ? "历史文件" : "最近文件";
    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "recent-clear";
    clearBtn.textContent = "清空";
    clearBtn.title = "清空全部历史记录";
    clearBtn.addEventListener("click", (e) => {
      e.stopPropagation();   // 不触发外部点击关闭
      clearHistory();
    });
    head.appendChild(subTitle);
    head.appendChild(clearBtn);
    box.appendChild(head);

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
      box.appendChild(delMissing);
    }

    // 列表：默认前 HIST_SHOWN 条
    const list = document.createElement("div");
    list.className = "recent-list";
    const shown = recentFiles.slice(0, HIST_SHOWN);
    const rest = recentFiles.slice(HIST_SHOWN);
    shown.forEach((e) => list.appendChild(buildRecentItem(e)));
    box.appendChild(list);

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
      box.appendChild(moreBtn);
    }

    // 源卡打开时：底部附「选择本地文件…」入口，保持系统对话框可用。
    if (recentTargetSrc) {
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
      box.appendChild(pickLocal);
    }

    // 全局弹框底部：查询缓存管理（显示占用 + 清空入口）
    if (!recentTargetSrc) {
      const cacheRow = document.createElement("div");
      cacheRow.className = "recent-cache-row";
      const cacheInfo = document.createElement("span");
      cacheInfo.className = "recent-cache-info";
      cacheInfo.textContent = "查询缓存…";
      const cacheClear = document.createElement("button");
      cacheClear.type = "button";
      cacheClear.className = "recent-cache-clear";
      cacheClear.textContent = "清除缓存";
      cacheClear.title = "删除已生成的 parquet 查询缓存（下次查询自动重建）";
      cacheClear.addEventListener("click", (e) => {
        e.stopPropagation();
        showConfirm("确定清空全部查询缓存吗？\n下次查询时会自动重新生成。", async () => {
          try {
            const r = await request("/api/cache/clear", {});
            showStatus(uploadStatus, `已清除 ${r.removed} 个缓存文件`, "ok");
            refreshCacheInfo(cacheInfo);
          } catch (err) {
            showStatus(uploadStatus, popError(err), "error");
          }
        });
      });
      cacheRow.appendChild(cacheInfo);
      cacheRow.appendChild(cacheClear);
      box.appendChild(cacheRow);
      refreshCacheInfo(cacheInfo);   // 异步拉取缓存占用
    }
  }

  // 刷新缓存占用信息到指定元素
  async function refreshCacheInfo(el) {
    if (!el) return;
    try {
      const r = await request("/api/cache/info", {});
      el.textContent = (r && r.count)
        ? `查询缓存：${r.count} 个文件 · ${fmtBytes(r.size_bytes)}`
        : "查询缓存：无";
      el.title = r && r.dir ? `缓存目录：${r.dir}` : "";
    } catch (_) {
      el.textContent = "查询缓存：—";
    }
  }

  // 字节数人性化：B / KB / MB / GB
  function fmtBytes(n) {
    if (!Number.isFinite(n) || n <= 0) return "0 B";
    const units = ["B", "KB", "MB", "GB"];
    let v = n;
    let i = 0;
    while (v >= 1024 && i < units.length - 1) { v /= 1024; i++; }
    return (v >= 10 || i === 0 ? v.toFixed(0) : v.toFixed(1)) + " " + units[i];
  }

  // 恢复一个多源组合：清空当前 sources，按其快照逐个重建（只写状态，不加载字段，
  // 避免一次请求多个文件；用户点「查询」时由后端读取）。doris 源按 kind/conn/db/table 重建。
  async function restoreCombo(combo) {
    if (!combo || !combo.sources || !combo.sources.length) return;
    sources = [];
    combo.sources.forEach((s) => {
      if (s.kind === "doris") {
        const src = createSource(s.alias || `t${sources.length + 1}`, "", "", {
          kind: "doris",
          conn: s.conn || null,
          db: s.db || "",
          table: s.table || "",
        });
        src.sheets = [];
        src.columns = [];
        return;
      }
      const src = createSource(s.alias || `t${sources.length + 1}`, s.path, s.sheet || "");
      src.sheets = [];
      src.columns = [];
    });
    activeSourceId = sources.length ? sources[0].id : null;
    renderSources();
    refreshCompletion();   // 新路径进入 SQL 补全候选
    showStatus(uploadStatus, `已恢复组合：${combo.sources.length} 个源`, "ok");
    // 组合带上次 SQL 时，询问是否回填到当前 cell
    maybeApplySavedSql(combo.sql || "", `${combo.sources.length} 个源的组合`);
    // 并行预热：doris 源拉库表清单/字段，file 源拉字段（补全候选与字段区即时可用）。
    sources.forEach((s) => {
      if (s.kind === "doris") {
        describeDorisSource(s);
      } else if (s.path) {
        loadFileFields(s);
      }
    });
  }

  // 按快照 path/sheet 异步加载一个 file 源的字段与工作表（恢复组合/历史用）。
  async function loadFileFields(src) {
    if (!src || src.kind === "doris" || !src.path) return;
    try {
      const data = await request("/api/open", { path: src.path, sheet: src.sheet || null });
      src.path = data.path;
      src.sheets = data.sheets || [];
      src.columns = data.columns || [];
      if (src.id === activeSourceId) {
        renderSources();
        refreshCompletion();
      }
    } catch (_) { /* 文件可能已被移动/删除；查询时后端会再报 */ }
  }

  // 切换弹框展开 / 收起（居中模态，无需锚点定位）
  function toggleRecentPop(force, anchorEl) {
    if (anchorEl) recentAnchorEl = anchorEl;
    const open = force !== undefined ? force : !recentOpen;
    recentOpen = open;
    if (open) {
      renderRecentPop();
      recentPop.classList.remove("hidden");
      checkMissingPaths();   // 打开时异步校验历史文件中哪些已被移除
    } else {
      recentPop.classList.add("hidden");
    }
  }

  // 从源卡「选择」按钮打开弹框：目标源 + 锚点 = 该按钮
  function openRecentPopForSource(src, anchorEl) {
    recentTargetSrc = src;
    recentAnchorEl = anchorEl;
    renderRecentPop();
    recentOpen = true;
    recentPop.classList.remove("hidden");
    checkMissingPaths();   // 打开时异步校验
  }

  // 关闭弹框时清掉目标源（切回「最近打开」）
  function closeRecentPop() {
    recentOpen = false;
    recentTargetSrc = null;
    recentAnchorEl = null;
    recentPop.classList.add("hidden");
  }

  // 点遮罩关闭（点击弹框外部背景）
  // 注意：必须在 mousedown 时判断 target —— click 的 target 是按下/松开两点的
  // 公共最深祖先：在输入框内从后往前拖选（按在 input、松开在遮罩）时，
  // click target 会变成遮罩本体，导致「选字却关框」。mousedown 永远指向精确按下点。
  function closeRecentPopOnOutside(e) {
    if (recentOpen && e.target === recentPop) {
      closeRecentPop();
    }
  }

  function buildRecentItem(entry) {
    const base = basename(entry.path) || entry.path;
    const missing = !!missingPaths[entry.path];
    const row = document.createElement("div");
    row.className = "recent-item" + (missing ? " recent-missing" : "");
    row.title = missing ? `${entry.path}\n（文件已被移动或删除）` : entry.path;

    const nameEl = document.createElement("span");
    nameEl.className = "recent-name";
    // 有名称 → 显示名称（加粗）+ 灰色文件名小字；无名称 → 直接显示文件名
    if (entry.name) {
      nameEl.innerHTML = `<b>${escapeHtml(entry.name)}</b> <i class="recent-combo-label">${escapeHtml(base)}</i>`;
    } else {
      nameEl.textContent = base;
    }

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
      removeHistory(entry.id);
    });

    row.appendChild(nameEl);
    row.appendChild(timeEl);
    row.appendChild(del);

    row.addEventListener("click", () => loadHistoryEntry(entry));
    return row;
  }

  // 把一段 SQL 回填到当前聚焦的 cell（无则最后一个/第一个）；用于恢复历史/组合时按需加载上次 SQL。
  function applySqlToActiveCell(sql) {
    if (!sql) return;
    let cell = getCell(activeCellId);
    if (!cell) cell = cells[cells.length - 1];
    if (!cell) return;
    cell.dom.cm.setValue(sql);
    autoSizeCell(cell);
    expandCell(cell.id);
    persistSession();
  }

  // 若条目带上次 SQL，用自定义确认弹框询问用户是否回填；确认则写回当前 cell。
  function maybeApplySavedSql(sql, what) {
    if (!sql) return;
    showConfirm(`是否加载上次查询的 SQL？\n${what}`, () => applySqlToActiveCell(sql));
  }

  // ---------- 通用确认弹框（替代 window.confirm）----------
  // 原生 confirm 在 WebView2 里会显示页面 URL 且为系统深色样式，与界面割裂。
  let confirmOkCallback = null;
  function showConfirm(message, onOk) {
    confirmMsg.textContent = message || "";
    confirmOkCallback = onOk || null;
    confirmPop.classList.remove("hidden");
  }
  function hideConfirm() {
    confirmPop.classList.add("hidden");
    confirmOkCallback = null;
  }

  // 点击历史条目：按当前模式复用现有加载链路（单文件 openSingle / 多源 refreshSource）。
  async function loadHistoryEntry(entry) {
    if (!entry || !entry.path) return;
    const targetSrc = recentTargetSrc;   // 先取目标源，再关闭浮层（关闭会清掉它）
    const savedSql = entry.sql || "";
    closeRecentPop();
    if (mode === "multi" && targetSrc) {
      // 源卡浮层：赋给打开浮层时指定的那个源（而不是当前激活的源）
      const src = targetSrc;
      const isFirstPick = !src.path;
      src.path = entry.path;
      src.sheet = entry.sheet || "";
      await refreshSource(src, !isFirstPick);
      maybeApplySavedSql(savedSql, basename(entry.path));
    } else if (mode === "multi") {
      // 全局浮层（多源模式，无指定源）：赋给当前激活源
      const targets = sources.length ? sources : [createSource("t1", "", "")];
      const src = targets.find((s) => s.id === activeSourceId) || targets[0];
      const isFirstPick = !src.path;
      src.path = entry.path;
      src.sheet = entry.sheet || "";
      await refreshSource(src, !isFirstPick);
      maybeApplySavedSql(savedSql, basename(entry.path));
    } else {
      singlePath = entry.path;
      // 恢复该文件上次用的工作表：若当前下拉有对应项则选中，否则留默认
      const need = entry.sheet || "";
      if (need && Array.from(sheetInput.options).some((o) => o.value === need)) {
        sheetInput.value = need;
      } else {
        sheetInput.value = "";
      }
      // 精确刷新到被点开的那一条（history 条目可能同名多条），再走打开链路
      recordHistory(singlePath, need, entry.name, entry.id);
      await openSingle();
      maybeApplySavedSql(savedSql, entry.name || basename(entry.path));
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

  // ---------- 数据库连接（Doris）管理 ----------
  // 已保存的连接配置存后端 SQLite（本机桌面应用；密码可选不保存）。
  // 每条形如 { id, name, host, port, user, password, db }
  const DBCONN_KEY = "excel_sql_dbconns_v1";
  const DBCONN_MAX = 10;
  let dbConns = [];   // 由 bootstrap 从后端 SQLite 载入
  let dbPendingConn = null;   // 已探测成功、等待选表添加的连接配置
  let dbPendingTables = [];   // 探测返回的表清单 [{db, table}] 拍平
  let dbEditTarget = null;    // 编辑模式：正在切换库/表的 doris 源（null = 新建模式）

  function loadDbConns(raw) {
    try {
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr.filter((c) => c && c.host) : [];
    } catch (_) { return []; }
  }
  function saveDbConns() {
    storeSet({ [DBCONN_KEY]: JSON.stringify(dbConns) });
  }
  function upsertDbConn(conn) {
    // 按 host+port+user+db 去重：更新已有或新增
    const key = `${conn.host}:${conn.port}:${conn.user}:${conn.db}`;
    const idx = dbConns.findIndex(
      (c) => `${c.host}:${c.port}:${c.user}:${c.db}` === key
    );
    if (idx >= 0) {
      dbConns[idx] = { ...dbConns[idx], ...conn };
      return dbConns[idx];
    }
    conn.id = conn.id || ("dbc_" + Date.now().toString(36));
    dbConns.unshift(conn);
    if (dbConns.length > DBCONN_MAX) dbConns.length = DBCONN_MAX;
    saveDbConns();
    return conn;
  }

  function renderDbConnSaved() {
    if (!dbConnSaved) return;
    dbConnSaved.innerHTML = "";
    const def = document.createElement("option");
    def.value = "";
    def.textContent = dbConns.length ? "（选择历史连接自动填充）" : "（暂无已保存的连接）";
    dbConnSaved.appendChild(def);
    dbConns.forEach((c, i) => {
      const o = document.createElement("option");
      o.value = String(i);
      const label = c.name || `${c.user}@${c.host}:${c.port}`;
      o.textContent = `${label}${c.db ? ` · ${c.db}` : ""}`;
      dbConnSaved.appendChild(o);
    });
    dbConnSaved.value = "";
    // 有历史连接才显示选择器；否则隐藏（首用无历史）
    if (dbConnSavedField) dbConnSavedField.classList.toggle("hidden", !dbConns.length);
  }

  function openDbConnPop() {
    renderDbConnSaved();
    // 默认清空表单；只有用户从「已保存的连接」下拉选中后才会填充。
    fillDbConnForm(null, false);
    dbConnStatus.textContent = "";
    dbConnStatus.className = "db-conn-status";
    dbConnStep2.classList.add("hidden");
    dbConnPop.classList.remove("hidden");
  }

  // 用一条连接配置填充表单；remember 控制「记住连接」勾选与密码是否回填。
  function fillDbConnForm(conn, remember) {
    dbConnName.value = conn ? (conn.name || "") : "";
    dbConnHost.value = conn ? (conn.host || "") : "";
    dbConnPort.value = conn ? (conn.port || 9030) : 9030;
    dbConnUser.value = conn ? (conn.user || "") : "";
    dbConnPass.value = conn && conn.password ? conn.password : "";
    dbConnRemember.checked = remember;
  }
  function closeDbConnPop() {
    dbConnPop.classList.add("hidden");
    dbPendingConn = null;
    dbPendingTables = [];
  }

  function currentDbConnForm() {
    return {
      name: dbConnName.value.trim(),
      host: dbConnHost.value.trim(),
      port: parseInt(dbConnPort.value, 10) || 9030,
      user: dbConnUser.value.trim(),
      password: dbConnPass.value || "",
      db: "",
    };
  }

  async function testDbConn() {
    const conn = currentDbConnForm();
    if (!conn.host) {
      dbConnStatus.textContent = "请填写主机";
      dbConnStatus.className = "db-conn-status error";
      return;
    }
    dbConnStatus.textContent = "正在连接…";
    dbConnStatus.className = "db-conn-status";
    try {
      const res = await request("/api/db_probe", { conn });
      dbPendingConn = conn;
      dbPendingTables = [];
      // 拍平 {db: [table...]} → [{db, table}]，供下拉选择
      const tables = (res.tables && typeof res.tables === "object") ? res.tables : {};
      Object.keys(tables).sort().forEach((db) => {
        (tables[db] || []).forEach((t) => dbPendingTables.push({ db, table: t }));
      });
      dbConnStatus.textContent = `连接成功（${res.server}），共 ${dbPendingTables.length} 张表`;
      dbConnStatus.className = "db-conn-status ok";
      renderDbTableOptions();
      dbConnStep2.classList.remove("hidden");
    } catch (e) {
      dbPendingConn = null;
      dbConnStep2.classList.add("hidden");
      dbConnStatus.textContent = popError(e);
      dbConnStatus.className = "db-conn-status error";
    }
  }

  function renderDbTableOptions() {
    const prev = dbTableSelect.value;
    dbTableSelect.innerHTML = "";
    // 首项：不绑定具体表（添加连接级源，之后在字段区选库表）
    const noneOpt = document.createElement("option");
    noneOpt.value = "";
    noneOpt.textContent = "（不选表，添加连接）";
    dbTableSelect.appendChild(noneOpt);
    if (!dbPendingTables.length) {
      return;
    }
    const groupBy = {};
    dbPendingTables.forEach((t) => {
      (groupBy[t.db] = groupBy[t.db] || []).push(t.table);
    });
    Object.keys(groupBy).sort().forEach((db) => {
      const og = document.createElement("optgroup");
      og.label = db;
      groupBy[db].forEach((tbl) => {
        const o = document.createElement("option");
        o.value = `${db}\u0001${tbl}`;   // 用 \u0001 分隔 db|table，避免表名含点歧义
        o.textContent = tbl;
        og.appendChild(o);
      });
      dbTableSelect.appendChild(og);
    });
    if (prev) dbTableSelect.value = prev;
  }

  function addDbSource() {
    if (!dbPendingConn) return;
    const val = dbTableSelect.value;
    // 连接级：允许不选表直接添加连接（db/table 留空，稍后在字段区选）
    let db = "";
    let table = "";
    if (val) {
      const parts = val.split("\u0001");
      db = parts[0] || "";
      table = parts[1] || "";
    }
    const conn = { ...dbPendingConn, db };
    // 记住连接（勾选时保存密码；不勾则保存时密码置空）
    if (dbConnRemember.checked) {
      upsertDbConn(conn);
    } else {
      upsertDbConn({ ...conn, password: "" });
    }
    const n = sources.length + 1;
    const src = createSource(`d${n}`, "", "", {
      kind: "doris",
      conn,
      db,
      table,
    });
    closeDbConnPop();
    renderSources();
    refreshCompletion();
    // 异步拉库表清单（连接已建连，无默认表时字段区出库表下拉）
    describeDorisSource(src);
    showStatus(uploadStatus, table ? `已添加数据库源：${db}.${table}` : `已添加数据库连接：${conn.host}:${conn.port}`, "ok");
  }

  async function describeDorisSource(src) {
    if (!src || src.kind !== "doris") return;
    // 连接级源：优先拉库表清单（供字段区切换）；有默认表则顺带拉字段。
    if (!src.dbTables) {
      try {
        const res = await request("/api/db_probe", { conn: src.conn || null });
        src.dbTables = {
          dbs: res.dbs || [],
          tables: res.tables || {},
        };
      } catch (e) {
        showStatus(uploadStatus, popError(e), "error");
        return;
      }
    }
    if (src.table) {
      await refreshDorisFields(src, false);
    } else {
      // 无默认表：只更新下拉，字段区保持空态引导
      if (src.id === activeSourceId) renderDorisPicker(src);
      if (src.id === activeSourceId) renderFields([]);
    }
  }

  // 按 src 当前 db/table 拉字段；成功则回填 columns 并刷新视图。
  async function refreshDorisFields(src, renderPicker = true) {
    if (!src || src.kind !== "doris" || !src.db || !src.table) return;
    try {
      const res = await request("/api/describe", {
        sources: [{
          alias: src.alias,
          kind: "doris",
          path: "",
          sheet: null,
          conn: src.conn || null,
          db: src.db,
          table: src.table,
        }],
        sql: "",   // /api/describe 用 QueryRequest 模型，sql 必填（此处占位，后端 describe 分支不读它）
      });
      const info = (res.sources && res.sources[0]) || {};
      src.columns = info.columns || [];
      if (src.id === activeSourceId) {
        if (renderPicker) renderDorisPicker(src);
        renderFields(src.columns);
      }
    } catch (e) {
      // 字段加载失败不打断（查询时后端会再报）
      showStatus(uploadStatus, popError(e), "error");
    }
  }

  // DOM 元素
  const modeSingleBtn = document.getElementById("modeSingleBtn");
  const modeMultiBtn = document.getElementById("modeMultiBtn");
  const singleModeBox = document.getElementById("singleModeBox");
  const multiModeBox = document.getElementById("multiModeBox");
  const pickBtn = document.getElementById("pickBtn");
  const addSourceBtn = document.getElementById("addSourceBtn");
  const addDbBtn = document.getElementById("addDbBtn");
  const dbConnPop = document.getElementById("dbConnPop");
  const dbConnClose = document.getElementById("dbConnClose");
  const dbConnName = document.getElementById("dbConnName");
  const dbConnSaved = document.getElementById("dbConnSaved");
  const dbConnSavedField = document.getElementById("dbConnSavedField");
  const dbConnHost = document.getElementById("dbConnHost");
  const dbConnPort = document.getElementById("dbConnPort");
  const dbConnUser = document.getElementById("dbConnUser");
  const dbConnPass = document.getElementById("dbConnPass");
  const dbConnRemember = document.getElementById("dbConnRemember");
  const dbConnStatus = document.getElementById("dbConnStatus");
  const dbConnStep2 = document.getElementById("dbConnStep2");
  const dbTableSelect = document.getElementById("dbTableSelect");
  const dbConnAdd = document.getElementById("dbConnAdd");
  const dorisPicker = document.getElementById("dorisPicker");
  const dorisDbSelect = document.getElementById("dorisDbSelect");
  const dorisTableSelect = document.getElementById("dorisTableSelect");
  const sourceList = document.getElementById("sourceList");
  const fileInfo = document.getElementById("fileInfo");
  const fileNameEl = document.getElementById("fileName");
  const clearSingleBtn = document.getElementById("clearSingleBtn");
  const uploadStatus = document.getElementById("uploadStatus");
  const openStages = document.getElementById("openStages");
  const recentBtn = document.getElementById("recentBtn");
  const saveBtn = document.getElementById("saveBtn");
  const recentCount = document.getElementById("recentCount");
  const recentPop = document.getElementById("recentPop");
  const recentPopBody = document.getElementById("recentPopBody");
  const recentPopClose = document.getElementById("recentPopClose");
  const fieldSearch = document.getElementById("fieldSearch");
  const fieldsBox = document.getElementById("fieldsBox");
  const sourceLabel = document.getElementById("sourceLabel");
  const confirmPop = document.getElementById("confirmPop");
  const confirmMsg = document.getElementById("confirmMsg");
  const confirmOk = document.getElementById("confirmOk");
  const confirmCancel = document.getElementById("confirmCancel");
  const saveComboPop = document.getElementById("saveComboPop");
  const saveComboNameInput = document.getElementById("saveComboNameInput");
  const saveComboTip = document.getElementById("saveComboTip");
  const saveComboClose = document.getElementById("saveComboClose");
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
    // FastAPI 422 校验错误的 detail 是 [{loc, msg, type}] 数组，直接 String() 会
    // 变成 "[object Object]"，这里拍平成可读文本（取 msg，缺省回退 loc 路径）。
    const d = err?.detail;
    if (Array.isArray(d)) {
      const parts = d.map((x) => x?.msg || (x?.loc ? x.loc.join(".") : "") || String(x))
        .filter(Boolean);
      return parts.length ? parts.join("；") : "请求参数错误";
    }
    return d || err?.message || String(err || "未知错误");
  }

  // 展示错误信息到 cell 状态（主文本 + title 悬浮存原始详情）
  function showCellError(cell, err) {
    if (!cell || !cell.dom) return;
    const text = popError(err);
    cell.dom.status.textContent = text;
    cell.dom.status.className = "cell-status error";
    // 完整信息（可能是多行翻译）放进 title，悬浮可见
    cell.dom.status.title = text;
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
  // 「最近打开」图标按钮：随模式移动位置——单文件跟在「选择文件」行，多源跟在「+ 文件/+ Doris」行。
  function relocateRecentBtn() {
    if (!recentBtn) return;
    // 保存按钮与最近打开按钮始终并排，一起移动
    const btns = [saveBtn, recentBtn].filter(Boolean);
    if (mode === "single") {
      const line = singleModeBox.querySelector(".pick-line");
      if (line) btns.forEach((b) => line.appendChild(b));
    } else {
      const row = multiModeBox.querySelector(".src-add-row");
      if (row) btns.forEach((b) => row.appendChild(b));
    }
  }

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
    relocateRecentBtn();   // 「最近打开」图标按钮随模式移动到对应行
    persistSession();
  }

  // ---------- 单文件模式 ----------
  function renderSingle() {
    fileNameEl.textContent = singlePath || "";
    fileNameEl.title = singlePath || "";
    fileInfo.classList.toggle("hidden", !singlePath);
    sourceLabel.textContent = "";   // 单文件无别名，清掉多源残留的「（别名）」
    renderDorisPicker(null);        // 隐藏多源 Doris 源的库表下拉，避免切模式残留
    renderFields(singlePath ? getSingleColumns() : null);
    renderSaveBtn();
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
  function createSource(alias, path, sheet, opts) {
    const o = opts || {};
    const src = {
      id: ++sourceSeq,
      alias: alias || "",
      kind: o.kind || "file",          // "file" | "doris"
      path: path || "",
      sheet: sheet || "",
      conn: o.conn || null,            // doris：连接配置 {host,port,user,password,db,name,id}
      db: o.db || "",                  // doris：当前选中/默认数据库名
      table: o.table || "",            // doris：当前选中/默认表名
      dbTables: null,                  // doris：{dbs:[...], tables:{db:[table...]}} 库表清单（连接级）
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
      // doris 源没有本地文件：不渲染选择按钮（更换走「连接数据库」）。
      const pickOne = src.kind !== "doris" ? (() => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "src-pick";
        b.innerHTML = ICON_FOLDER;
        b.title = src.path ? "更换此数据源的文件（含最近打开）" : "为该数据源选择文件（含最近打开）";
        b.addEventListener("click", (e) => {
          e.stopPropagation();   // 不触发行选中
          // 无历史文件时直接走系统对话框；有则弹出浮层
          if (recentFiles.length) {
            openRecentPopForSource(src, b);
          } else {
            pickFileForSource(src);
          }
        });
        return b;
      })() : null;

      head.appendChild(badge);
      head.appendChild(aliasInput);
      if (pickOne) head.appendChild(pickOne);
      head.appendChild(del);

      // 文件行：文件名（单行省略）+ 字段数徽标（选择/更换已在卡头行）
      // doris 源：显示 🔗 db.table，无 sheet
      const pathRow = document.createElement("div");
      pathRow.className = "src-path-row";

      const pathDisplay = document.createElement("div");
      pathDisplay.className = "src-path";
      if (src.kind === "doris") {
        const hostLabel = src.conn ? `${src.conn.host}:${src.conn.port}` : "";
        pathDisplay.textContent = src.db && src.table
          ? `🔗 ${src.db}.${src.table}`
          : (hostLabel ? `🔌 ${hostLabel}` : "（数据库连接）");
        pathDisplay.title = src.conn
          ? `${hostLabel}${src.db && src.table ? ` / ${src.db}.${src.table}` : ""}（SQL 里可写 别名.库.表）`
          : "";
      } else {
        pathDisplay.textContent = src.path ? basename(src.path) : "（未选择文件）";
        pathDisplay.title = src.path || "";
      }

      pathRow.appendChild(pathDisplay);
      if (src.columns && src.columns.length) {
        const colCount = document.createElement("span");
        colCount.className = "src-colcount";
        colCount.textContent = `${src.columns.length} 字段`;
        pathRow.appendChild(colCount);
      }

      // 工作表下拉：仅多 sheet 时展示且是 file 源，单 sheet / 未选文件时隐藏
      let sheetRow = null;
      let sheetSel = null;
      if (src.kind !== "doris" && src.sheets && src.sheets.length > 1) {
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
    renderDorisPicker(active);
    renderFields(active.columns);
    renderSaveBtn();
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

  // ---------- Doris 连接级源：字段区库/表下拉 ----------
  // 仅当选中源是 doris 且已拉到库表清单（src.dbTables）时显示；否则隐藏。
  function renderDorisPicker(src) {
    const isDoris = src && src.kind === "doris";
    if (!isDoris || !src.dbTables) {
      dorisPicker.classList.add("hidden");
      return;
    }
    dorisPicker.classList.remove("hidden");
    const dbs = src.dbTables.dbs || [];
    const tables = src.dbTables.tables || {};

    // 库下拉
    dorisDbSelect.innerHTML = "";
    if (!dbs.length) {
      const o = document.createElement("option");
      o.value = "";
      o.textContent = "（无可用库）";
      dorisDbSelect.appendChild(o);
    } else {
      dbs.forEach((db) => {
        const o = document.createElement("option");
        o.value = db;
        o.textContent = db;
        dorisDbSelect.appendChild(o);
      });
    }
    dorisDbSelect.value = src.db && dbs.includes(src.db) ? src.db : (dbs[0] || "");

    // 表下拉：随当前库联动
    const fillTables = () => {
      const curDb = dorisDbSelect.value;
      dorisTableSelect.innerHTML = "";
      const list = tables[curDb] || [];
      if (!list.length) {
        const o = document.createElement("option");
        o.value = "";
        o.textContent = "（无可用表）";
        dorisTableSelect.appendChild(o);
        return;
      }
      list.forEach((t) => {
        const o = document.createElement("option");
        o.value = t;
        o.textContent = t;
        dorisTableSelect.appendChild(o);
      });
      // 切换库后默认选第一张表（若原本的表不在此库）
      const keep = src.table && list.includes(src.table) && src.db === curDb;
      dorisTableSelect.value = keep ? src.table : (list[0] || "");
    };
    fillTables();

    // 事件绑定（每次重渲染都重建 select，需重新挂；不重复累加因为都是新建元素）
    dorisDbSelect.onchange = async () => {
      src.db = dorisDbSelect.value;
      src.table = "";   // 换库先清表，由 fillTables 同步
      fillTables();
      src.table = dorisTableSelect.value;
      await refreshDorisFields(src, false);
    };
    dorisTableSelect.onchange = async () => {
      src.table = dorisTableSelect.value;
      await refreshDorisFields(src, false);
    };
  }

  // ---------- 字段渲染（含搜索过滤） ----------
  function renderFields(columns) {
    const kw = fieldSearch.value.trim().toLowerCase();
    const list = (columns || []).filter((c) =>
      !kw || String(c.name).toLowerCase().includes(kw)
    );

    fieldsBox.innerHTML = "";
    if (!columns || !columns.length) {
      const active = sources && sources.find((s) => s.id === activeSourceId);
      if (active && active.kind === "doris" && !active.table) {
        fieldsBox.innerHTML = '<p class="placeholder">在上面选择库和表，字段会列在这里；查询时 SQL 可写 <code>别名.库.表</code></p>';
      } else {
        fieldsBox.innerHTML = '<p class="placeholder">读取字段后，这里自动列出列名与类型</p>';
      }
      return;
    }
    if (!list.length) {
      fieldsBox.innerHTML = '<p class="placeholder">无匹配字段</p>';
      return;
    }
    list.forEach((c) => {
      const chip = document.createElement("span");
      chip.className = "field-chip";
      const tipParts = [c.type];
      if (c.is_key) tipParts.push("键列");
      if (c.comment) tipParts.push(c.comment);
      chip.title = tipParts.filter(Boolean).join(" · ");

      // 键列标记（Doris 源才有 is_key；file 源无此字段按非键列处理）
      if (c.is_key) {
        const keyMark = document.createElement("span");
        keyMark.className = "fkey";
        keyMark.textContent = "🔑";
        chip.appendChild(keyMark);
      }
      const nameSpan = document.createElement("span");
      nameSpan.className = "fname";
      nameSpan.textContent = c.name;
      chip.appendChild(nameSpan);

      const typeSpan = document.createElement("span");
      typeSpan.className = "ftype";
      typeSpan.textContent = c.type;
      chip.appendChild(typeSpan);

      // Doris 字段注释（有则作副标题挂到 chip 底部，无注释维持名+类型一行式）
      if (c.comment) {
        const cmt = document.createElement("span");
        cmt.className = "fcomment";
        cmt.textContent = c.comment;
        chip.appendChild(cmt);
      }

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
  const ICON_DB = '<svg viewBox="0 0 14 14" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.3"><ellipse cx="7" cy="3.5" rx="5" ry="2"/><path d="M2 3.5v7c0 1.1 2.24 2 5 2s5-.9 5-2v-7"/><path d="M2 7c0 1.1 2.24 2 5 2s5-.9 5-2"/></svg>';

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
    runBtn.title = "运行查询（Ctrl+Enter）";
    // 运行中点击 = 取消当前查询；空闲点击 = 运行
    runBtn.addEventListener("click", () => {
      if (cell._running) cancelCellQuery(cell.id);
      else runCell(cell.id);
    });

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
    // 内容变化后防抖持久化，避免每次击键都写后端 SQLite
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
    return sources.map((s) => {
      if (s.kind === "doris") {
        return {
          alias: s.alias,
          kind: "doris",
          path: "",            // 占位，doris 源用 conn/db/table
          sheet: null,
          conn: s.conn || null,
          db: s.db || "",
          table: s.table || "",
        };
      }
      return {
        alias: s.alias,
        path: s.path,
        sheet: s.sheet || null,
      };
    });
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
    if (cell._running) return;   // 运行中：忽略重复触发（按钮/Ctrl+Enter）

    const payloadSources = buildSourcesPayload();
    if (!payloadSources.length) {
      status.textContent = "请先添加数据源";
      status.className = "cell-status error";
      return;
    }
    // 数据源合法性：file 需有路径；doris 需有连接配置（库表可在 SQL 显式指定，不强求）
    const badFile = payloadSources.some((s) => s.kind !== "doris" && !s.path);
    const badDoris = payloadSources.some((s) => s.kind === "doris" && !s.conn);
    if (badFile || badDoris) {
      status.textContent = "请先为每个数据源选择文件 / 配置数据库连接";
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

    cell._running = true;
    cell._taskId = null;
    // 运行中按钮变为「取消」（原黑色运行按钮换成描边样式）
    runBtn.textContent = "✕ 取消";
    runBtn.title = "取消当前查询";
    runBtn.classList.add("cell-run-cancel");
    runBtn.disabled = false;

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
      cell._taskId = taskId;

      // 2) 轮询进度，直到 done / error（自适应退避：准备阶段慢轮询，运行阶段适度加快）
      let data = null;
      let pollDelay = 80;
      let timedOut = false;
      let cancelled = false;
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
          // 查询成功：把 SQL 快照记到源组合 / 单文件历史，供下次恢复时提示是否回填。
          if (mode === "multi") {
            recordCombo(sources, sql);
          } else if (singlePath) {
            recordHistorySql(singlePath, sql);
          }
          break;
        }
        if (p.status === "error") {
          throw new Error(p.error || "查询执行出错");
        }
        if (p.status === "cancelled") {
          cancelled = true;
          break;
        }
        if (p.status === "timeout") {
          timedOut = true;
          break;
        }
        await new Promise((r) => setTimeout(r, pollDelay));
        // 线性退避到 500ms 上限之后不再增长，避免进度条更新显得迟滞
        pollDelay = Math.min(500, pollDelay + 40);
      }

      if (cancelled) {
        status.textContent = "已取消查询";
        status.className = "cell-status";
        return;
      }
      if (timedOut) {
        status.textContent = "查询超过时限，已自动中止";
        status.className = "cell-status error";
        return;
      }

      cell.lastResult = data;
      cell._lastTaskId = taskId;   // 供「按 task_id 导出」直读后端缓存（大结果集不回传前端）
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
      showCellError(cell, e);
    } finally {
      // 恢复「运行」按钮状态
      cell._running = false;
      cell._taskId = null;
      runBtn.classList.remove("cell-run-cancel");
      runBtn.textContent = "▶ 运行";
      runBtn.title = "运行查询（Ctrl+Enter）";
      runBtn.disabled = false;
      stopPrepare(cell);
      // 完成/出错/取消后再显示一小段时间后隐藏进度条
      setTimeout(() => {
        if (progress) progress.classList.add("hidden");
      }, 350);
    }
  }

  // 取消一个正在运行的 cell 查询：通知后端置 cancelled（后台线程会 interrupt() 终止）
  async function cancelCellQuery(id) {
    const cell = getCell(id);
    if (!cell || !cell._running || !cell._taskId) return;
    const taskId = cell._taskId;
    cell._taskId = null;   // 防重复点击
    try {
      await request("/api/cancel", { task_id: taskId });
    } catch (_) { /* 后端已结束/不可达：无需处理，轮询会收尾 */ }
    // 轮询循环下一次 /api/progress 会拿到 cancelled 状态并收尾；若请求失败
    // （任务已被惰性清理），这里兜底把 UI 复位
    setTimeout(() => {
      if (cell._running && !cell._taskId) {
        cell._running = false;
        cell.dom.runBtn.classList.remove("cell-run-cancel");
        cell.dom.runBtn.textContent = "▶ 运行";
        cell.dom.runBtn.title = "运行查询（Ctrl+Enter）";
        cell.dom.runBtn.disabled = false;
        cell.dom.status.textContent = "已取消查询";
        cell.dom.status.className = "cell-status";
      }
    }, 1500);
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
    const columnTypes = data.column_types || null;
    const { result, resultWrap, rowCount } = cell.dom;

    // 清空旧结果，重建表格
    resultWrap.innerHTML = "";
    // 先解除隐藏（display:none 时 clientHeight=0，虚拟滚动的 paint() 会算不出
    // 可视行数，导致首屏只渲染前几行 + 大 spacer = 一半空白），再渲染表格。
    result.classList.remove("hidden");

    const virtual = rows.length > VSCROLL_THRESHOLD;
    resultWrap.classList.toggle("vscroll", virtual);
    if (virtual) {
      renderVirtualTable(cell, cols, rows, columnTypes);
    } else {
      resultWrap.appendChild(buildFullTable(cols, rows, 0, rows.length, columnTypes));
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

  // 计算各列合计：用「列名黑名单 → 类型白名单 → 数字占比启发式」三层判定，
  // 避免把状态、日期、编号等列误当成数值求和。
  //
  // 判定顺序（任一命中即生效）：
  //   1. 列名匹配黑名单关键词（状态/日期/时间/类型/编号/序号/id 等）→ 不求和
  //   2. 类型明确是数值（int/decimal/float/double…）→ 求和
  //   3. 类型是日期时间 → 不求和
  //   4. 类型缺失或文本 → 回退「非空值中数字占比 ≥ 60%」启发式（兼容全文本读出的金额列）
  const TOTAL_NUM_RATIO = 0.6;
  // 列名黑名单关键词（小写匹配；命中即视为语义列，不求和）
  const TOTAL_NAME_BLACKLIST = /state|status|type|kind|code|no\b|id\b|date|time|year|month|week|flag|bool|是否|状态|类型|编号|序号|日期|时间|年|月|周|编码|等级|级别|标志/;
  // 数值类型白名单（DuckDB 类型名；DECIMAL(p,s) 与 INT 系的别名都含在内）
  const NUMERIC_TYPE_RE = /^(decimal|dec|numeric|tinyint|smallint|integer|int|bigint|hugeint|utinyint|usmallint|uinteger|ubigint|uhugeint|float|real|double)/;
  // 日期时间类型（硬排除）
  const TEMPORAL_TYPE_RE = /^(date|datetime|timestamp|time|interval)/;
  // 布尔类型（硬排除：0/1 全是数字，落到启发式会被误求和）
  const BOOLEAN_TYPE_RE = /^bool/;

  function isNumericColumn(col, type) {
    // 1) 列名黑名单
    if (TOTAL_NAME_BLACKLIST.test(String(col).toLowerCase())) return false;
    if (!type) return null;   // 无类型信息 → 交给启发式
    const t = String(type).toLowerCase().trim();
    // 2) 日期时间 / 布尔 → 明确不求和
    if (TEMPORAL_TYPE_RE.test(t) || BOOLEAN_TYPE_RE.test(t)) return false;
    // 3) 数值类型白名单 → 求和
    if (NUMERIC_TYPE_RE.test(t)) return true;
    // 4) 其余（varchar/text 等）→ 启发式
    return null;
  }

  function computeColumnTotals(cols, rows, columnTypes) {
    if (!rows.length) return null;
    const sums = new Array(cols.length).fill(null);
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      const type = columnTypes ? columnTypes[ci] : null;
      const numericKind = isNumericColumn(col, type);

      // 类型明确非数值（黑名单/日期时间/文本且非启发式通道）→ 跳过
      if (numericKind === false) continue;

      let sum = 0;
      let any = false;
      let nonEmpty = 0;
      let numeric = 0;
      for (let i = 0; i < rows.length; i++) {
        const v = rows[i][col];
        if (v === null || v === undefined || v === "") continue;
        nonEmpty++;
        const n = toNumber(v);
        if (Number.isNaN(n)) continue;
        sum += n;
        any = true;
        numeric++;
      }
      if (!any) continue;
      // 类型明确是数值 → 直接求和；无类型信息/文本类型 → 启发式（数字占比达标才求和）
      if (numericKind === true || (numericKind === null && numeric / nonEmpty >= TOTAL_NUM_RATIO)) {
        sums[ci] = sum;
      }
    }
    return sums;
  }

  // 合计值格式化：整数原样输出，浮点去掉误差尾巴（0.30000000000000004 → 0.3）
  function formatTotal(v) {
    if (typeof v !== "number" || !Number.isFinite(v)) return "—";
    if (Number.isInteger(v)) return String(v);
    return String(parseFloat(v.toFixed(10)));
  }

  // 构建完整 table（行号列 + 表头 + from..to 之间数据行 + 底部合计行）；
  // to<=from 时只建空 tbody（虚拟滚动模式由 paint() 填充）。
  function buildFullTable(cols, rows, from, to, columnTypes) {
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const trHead = document.createElement("tr");
    // 行号列：固定表头「#」，不参与列数据
    const thIdx = document.createElement("th");
    thIdx.className = "row-index-head";
    thIdx.textContent = "#";
    trHead.appendChild(thIdx);
    cols.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c;
      trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let i = from; i < to; i++) {
      tbody.appendChild(buildRow(cols, rows[i], i));
    }
    table.appendChild(tbody);

    // 合计行：「合计」标签放在最左的序号列那一格；数据列统一按
    // 「数字列显示 SUM、非数字列显示 —」。无数据时不输出合计行。
    const totals = computeColumnTotals(cols, rows, columnTypes);
    if (totals) {
      const tfoot = document.createElement("tfoot");
      const tr = document.createElement("tr");
      tr.className = "totals-row";
      // 序号列位置：放「合计」标签（与表头「#」对齐）
      const thIdxF = document.createElement("td");
      thIdxF.className = "row-index-cell";
      thIdxF.textContent = "合计";
      tr.appendChild(thIdxF);
      cols.forEach((c, ci) => {
        const td = document.createElement("td");
        td.textContent = totals[ci] !== null ? formatTotal(totals[ci]) : "—";
        tr.appendChild(td);
      });
      tfoot.appendChild(tr);
      table.appendChild(tfoot);
    }
    return table;
  }

  // 行号列单元格：灰底、右对齐、等宽、窄列
  function buildRowIndexCell(i) {
    const td = document.createElement("td");
    td.className = "row-index-cell";
    td.textContent = String(i + 1);
    return td;
  }

  function buildRow(cols, r, i) {
    const tr = document.createElement("tr");
    if (typeof i === "number") tr.appendChild(buildRowIndexCell(i));
    cols.forEach((c) => {
      const td = document.createElement("td");
      const v = r[c];
      td.textContent = v === null || v === undefined ? "" : String(v);
      tr.appendChild(td);
    });
    return tr;
  }

  // 在 tbody 里放一个整行高为 px 的空行，用于撑出滚动高度（含行号占位）
  function buildSpacerRow(cols, px) {
    const tr = document.createElement("tr");
    tr.className = "vscroll-spacer-row";
    const td = document.createElement("td");
    td.colSpan = cols.length + 1;   // +1 行号列
    td.style.height = px + "px";
    td.style.padding = "0";
    td.style.border = "none";
    tr.appendChild(td);
    return tr;
  }

  function renderVirtualTable(cell, cols, rows, columnTypes) {
    const total = rows.length;
    const wrap = cell.dom.resultWrap;

    // tfoot 只建一次（含合计行），滚动重建 tbody 时不动它
    const table = buildFullTable(cols, rows, 0, 0, columnTypes);  // 表头 + 空 tbody
    wrap.appendChild(table);
    const tbody = table.querySelector("tbody");

    function paint() {
      const scrollTop = wrap.scrollTop;
      const first = Math.max(0, Math.floor(scrollTop / VSCROLL_ROW_PX) - VSCROLL_BUFFER);
      const last = Math.min(total, Math.ceil((scrollTop + wrap.clientHeight) / VSCROLL_ROW_PX) + VSCROLL_BUFFER);

      // 重建 tbody：顶部 spacer + 数据行 + 底部 spacer
      const frag = document.createDocumentFragment();
      if (first > 0) frag.appendChild(buildSpacerRow(cols, first * VSCROLL_ROW_PX));
      for (let i = first; i < last; i++) frag.appendChild(buildRow(cols, rows[i], i));
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

    const filename = "query_result_" + new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");

    const cellStatus = cell.dom.status;
    const prevText = cellStatus.textContent;
    const prevClass = cellStatus.className;
    cellStatus.textContent = "正在导出…";
    cellStatus.className = "cell-status";

    try {
      // 优先走 task_id 直读后端结果缓存：大结果集不经过前端 JSON 回传，避免双重传输。
      // 仅当任务已过期时回退为前端把 rows 原样传回（兼容旧路径）。
      const taskId = cell._lastTaskId;
      let res;
      if (taskId) {
        try {
          res = await request("/api/export", { task_id: taskId, filename });
        } catch (e) {
          // 404「结果已过期」→ 回退走前端传参；其他错误原样抛
          if (String(e?.detail || e) !== "查询结果已过期，请重新查询后再导出") throw e;
          res = null;
        }
      }
      if (!res) {
        res = await request("/api/export", {
          columns: cell.lastResult.columns,
          rows: cell.lastResult.rows || [],
          filename,
        });
      }
      if (!res || res.cancelled) {
        cellStatus.textContent = prevText;
        cellStatus.className = prevClass;
        return;
      }
      cellStatus.textContent = `已导出 ${res.rows ?? (cell.lastResult.rows || []).length} 行 → ${res.path}`;
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

  // 数据库连接浮层
  addDbBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    openDbConnPop();
  });
  dbConnClose.addEventListener("click", closeDbConnPop);
  dbConnPop.addEventListener("mousedown", (e) => {
    if (e.target === dbConnPop) closeDbConnPop();   // 点遮罩关闭（mousedown 判断，杜绝拖选误关）
  });
  dbConnHost.addEventListener("keydown", (e) => { if (e.key === "Enter") testDbConn(); });
  dbConnPort.addEventListener("keydown", (e) => { if (e.key === "Enter") testDbConn(); });
  dbConnUser.addEventListener("keydown", (e) => { if (e.key === "Enter") testDbConn(); });
  dbConnPass.addEventListener("keydown", (e) => { if (e.key === "Enter") testDbConn(); });
  dbConnName.addEventListener("keydown", (e) => { if (e.key === "Enter") testDbConn(); });
  document.getElementById("dbConnTest").addEventListener("click", testDbConn);
  dbConnAdd.addEventListener("click", addDbSource);
  // 选中已保存连接 → 自动填充表单（密码随「是否保存密码」决定能否回填）
  dbConnSaved.addEventListener("change", () => {
    const i = parseInt(dbConnSaved.value, 10);
    if (!Number.isInteger(i) || !dbConns[i]) return;
    fillDbConnForm(dbConns[i], true);
    dbConnStatus.textContent = "";
    dbConnStatus.className = "db-conn-status";
  });

  // 通用确认弹框：确定 → 执行回调并关闭；取消 → 仅关闭
  confirmOk.addEventListener("click", () => {
    const cb = confirmOkCallback;
    hideConfirm();
    if (cb) cb();
  });
  confirmCancel.addEventListener("click", hideConfirm);
  confirmPop.addEventListener("mousedown", (e) => {
    if (e.target === confirmPop) hideConfirm();   // 点遮罩取消（mousedown 判断，杜绝拖选误关）
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !confirmPop.classList.contains("hidden")) hideConfirm();
  });

  // 最近打开弹框：图标按钮切换，点击遮罩 / 关闭按钮 / Esc 关闭
  recentBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    if (recentOpen && recentAnchorEl === recentBtn) {
      closeRecentPop();
    } else {
      recentTargetSrc = null;          // 全局按钮打开：清掉源卡目标
      toggleRecentPop(true, recentBtn);
    }
  });
  // 主动保存：多源 → 打开「保存组合」对话框（改名/保存当前/另存为）；单文件 → 直接保存
  saveBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    saveCurrentState();
  });

  // 保存组合对话框：关闭按钮 / 遮罩 / Esc；回车触发「保存当前组合」
  saveComboClose.addEventListener("click", closeSaveComboPop);
  saveComboPop.addEventListener("mousedown", (e) => {
    if (e.target === saveComboPop) closeSaveComboPop();
  });
  saveComboNameInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); saveComboCurrent(); }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !saveComboPop.classList.contains("hidden")) closeSaveComboPop();
  });
  document.getElementById("saveComboCurrent").addEventListener("click", saveComboCurrent);
  document.getElementById("saveComboAs").addEventListener("click", saveComboAsCopy);
  recentPopClose.addEventListener("click", closeRecentPop);
  document.addEventListener("mousedown", closeRecentPopOnOutside);
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
  relocateRecentBtn();   // 初始把「最近打开」图标按钮放进单文件「选择文件」行

  // ---------- 启动：从后端 SQLite 拉取持久化数据 ----------
  // 历史/组合/连接/会话统一存后端 SQLite；这里一次性批量读取并初始化内存态。
  (async function bootstrap() {
    const values = await storeGet([HIST_KEY, COMBOS_KEY, DBCONN_KEY, STORE_KEY]);
    recentFiles = loadHistory(values[HIST_KEY]);
    combos = loadCombos(values[COMBOS_KEY]);
    dbConns = loadDbConns(values[DBCONN_KEY]);

    renderRecent();   // 渲染「最近打开」历史列表（无历史时整区隐藏）

    // 会话恢复：优先用上次保存的 SQL / 模式 / 折叠状态；无则建默认 cell
    const saved = restoreSession(values[STORE_KEY]);
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
  })();

  // 等宽字体加载完成后重新校准高度（避免按回退字体度量测量）
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(() => cells.forEach(autoSizeCell));
  }
})();