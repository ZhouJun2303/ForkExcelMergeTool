const state = {
  boot: null,
  screen: "settings",
  back: null,
  nav: "general",
  settings: null,
  integrations: null,
  fixtures: null,
  merge: null,
  compare: null,
  q: "",
  offset: 0,
  pick: 0,
  aiOpen: true,
  compareSheet: undefined,
  compareQ: "",
  compareOffset: 0,
  comparePick: null,
  showAllCells: false,
  detail: null,
  fields: { local: "", base: "", remote: "", merged: "", left: "", right: "", backup: "" },
  updateText: "",
  note: "",
  busy: false,
};

let focusRestore = null;
let searchTimer = 0;

const app = document.getElementById("app");
app.addEventListener("click", onClick);
app.addEventListener("change", onChange);
app.addEventListener("input", onInput);

if (window.chrome && window.chrome.webview) {
  window.chrome.webview.addEventListener("message", (event) => {
    const data = event.data;
    if (data && data.type === "folder") {
      state.fields.backup = data.path || "";
      saveSettings({ backupRootDir: state.fields.backup }).catch(showError);
    }
  });
}

boot().catch(showError);

async function boot() {
  const [bootBody, settings] = await Promise.all([api("/api/boot"), api("/api/settings")]);
  state.boot = bootBody;
  state.settings = settings;
  state.screen = bootBody.screen || "settings";
  state.fields.backup = settings.backupRootDir || "";
  document.body.classList.toggle("dark", !!settings.darkTheme);
  if (bootBody.mergeId) {
    state.merge = await api(mergeUrl(bootBody.mergeId));
  }
  if (bootBody.compareId) {
    state.compare = await api(compareUrl(bootBody.compareId));
    state.compareSheet = state.compare.activeSheet;
    if (settings.autoOpenCompare && state.compare.total > 0) {
      await exportCompare(true);
    }
  }
  if (state.screen === "settings") {
    state.fixtures = await api("/api/fixtures");
  }
  render();
}

let sizedScreen = "";

function render() {
  if (sizedScreen !== state.screen) {
    sizedScreen = state.screen;
    hostMessage("size:" + state.screen);
  }
  app.innerHTML = screens[state.screen] ? screens[state.screen]() : screens.settings();
  restoreFocus();
}

const screens = {
  settings: renderSettings,
  merge: renderMerge,
  compare: renderCompare,
  choose: renderChoose,
  backup: renderBackup,
};

function renderSettings() {
  const page = {
    general: renderGeneral,
    backup: renderBackupPage,
    fork: renderFork,
    git: renderGit,
    update: renderUpdate,
  }[state.nav] || renderGeneral;
  const error = state.boot && state.boot.error ? `<div class="banner">${esc(state.boot.error)}</div>` : "";
  return `<div class="settings">
    <nav class="nav">
      <h2>设置</h2>
      ${navButton("general", "常规")}
      ${navButton("backup", "备份")}
      ${navButton("fork", "Fork")}
      ${navButton("git", "Git")}
      ${navButton("update", "更新")}
    </nav>
    <div class="page"><div class="block">${error}${page()}</div></div>
  </div>`;
}

function navButton(id, label) {
  return `<button class="${state.nav === id ? "on" : ""}" data-act="nav" data-nav="${id}">${label}</button>`;
}

function renderGeneral() {
  const feature = state.settings.startupFeature;
  const back = state.back ? `<p><button data-act="back">返回</button></p>` : "";
  return `${back}
    <h1>默认运行模式</h1>
    <p class="muted">从 Fork、命令行或 Git 传入文件时的默认行为。</p>
    ${radio("快速备份模式", "backup_only", "只备份输入文件，不生成合并或对比结果。", feature)}
    ${radio("合并对比模式", "merge_diff", "打开完整三向合并或二向对比窗口。", feature)}
    ${radio("每次询问", "ask_each_time", "每次先选择备份还是合并/对比。适合大表。", feature)}
    <hr>
    <label class="check"><input type="checkbox" data-act="dark" ${state.settings.darkTheme ? "checked" : ""}> 深色主题</label>
    <hr>
    <h1>打开工作簿</h1>
    <p class="muted">浏览器和打包后的程序用同一页面。这里可以直接打开测试文件。</p>
    <div class="chips">${fixtureChips()}</div>
    <p>本地</p><input data-field="local" value="${esc(state.fields.local)}">
    <p>基准</p><input data-field="base" value="${esc(state.fields.base)}">
    <p>线上</p><input data-field="remote" value="${esc(state.fields.remote)}">
    <p>输出</p><input data-field="merged" value="${esc(state.fields.merged)}" placeholder="可留空，生成时写入临时目录">
    <p class="row"><button class="primary" data-act="open-merge">打开合并</button></p>
    <p>对比线上</p><input data-field="left" value="${esc(state.fields.left)}">
    <p>对比本地</p><input data-field="right" value="${esc(state.fields.right)}">
    <p class="row"><button class="primary" data-act="open-compare">打开对比</button></p>
    ${state.note ? `<p class="muted">${esc(state.note)}</p>` : ""}`;
}

function radio(title, value, hint, current) {
  return `<label class="radio"><input type="radio" name="startup" data-act="startup" data-value="${value}" ${current === value ? "checked" : ""}> ${esc(title)}</label>
    <div class="muted" style="margin-left:24px">${esc(hint)}</div>`;
}

function fixtureChips() {
  const data = state.fixtures;
  if (!data) return "";
  const merges = (data.merges || []).map((item) =>
    `<button data-act="fixture-merge" data-id="${esc(item.id)}">${esc(item.title)}</button>`).join("");
  const compares = (data.compares || []).map((item) =>
    `<button data-act="fixture-compare" data-id="${esc(item.id)}">${esc(item.title)}</button>`).join("");
  return merges + compares;
}

function renderBackupPage() {
  return `<h1>备份根目录</h1>
    <p class="muted">留空则使用目标文件同目录下的 MergeExcelBackup。</p>
    <input id="backup-root" data-field="backup" value="${esc(state.fields.backup)}" placeholder="留空则使用目标文件旁的 MergeExcelBackup">
    <p class="row">
      <button data-act="pick-backup">选择目录</button>
      <button class="primary" data-act="save-backup">保存</button>
      <button data-act="open-backup">打开目录</button>
    </p>`;
}

function renderFork() {
  const info = state.integrations;
  if (!info) return `<h1>Fork 一键注入</h1><p class="muted">正在读取...</p>`;
  return `<h1>Fork 一键注入</h1>
    <p class="muted">安装后，Fork 的 Merge Tool 和 External Diff Tool 会指向本工具。请先关闭 Fork。</p>
    <p>${esc(info.fork.detail)}</p>
    <p class="row"><button class="primary" data-act="fork-install">安装注入</button><button data-act="fork-uninstall">移除注入</button></p>
    ${copyRow("工具路径", info.toolPath)}
    ${copyRow("合并参数", info.mergeArgs)}
    ${copyRow("对比参数", info.diffArgs)}`;
}

function renderGit() {
  const info = state.integrations;
  if (!info) return `<h1>全局 Git 注入</h1><p class="muted">正在读取...</p>`;
  return `<h1>全局 Git 注入</h1>
    <p class="muted">安装后，任意 Git 工具遇到常见 Excel 后缀冲突都会调用本工具。</p>
    <p>${esc(info.git.detail)}</p>
    <p class="row"><button class="primary" data-act="git-install">安装注入</button><button data-act="git-uninstall">移除注入</button></p>
    ${copyRow("driver 值", info.driver)}
    ${copyRow("driver 参数", info.driverArgs)}`;
}

function copyRow(label, value) {
  return `<p style="margin:16px 0 4px;font-weight:600">${esc(label)}</p>
    <div class="copy"><input readonly value="${esc(value)}"><button data-act="copy" data-text="${esc(value)}">复制</button></div>`;
}

function renderUpdate() {
  const version = (state.settings && state.settings.version) || "";
  return `<h1>程序更新</h1>
    <p class="muted">检查 GitHub Releases。不会自动替换，必须确认后才更新。</p>
    <p>${esc(state.updateText || ("当前版本 v" + version))}</p>
    <button class="primary" data-act="check-update" ${state.busy ? "disabled" : ""}>检查更新</button>`;
}

function renderMerge() {
  const data = state.merge;
  if (!data) return `<div class="screen pad"><p>正在打开合并...</p></div>`;
  const options = data.options || {};
  const wizard = data.wizard || {};
  const item = (data.items || [])[state.pick] || null;
  return `<div class="screen">
    <div class="pad head">
      <div class="text">
        <h1>Excel 三向合并</h1>
        <div class="muted" style="margin-top:4px">${esc(data.localCommit || "")}</div>
        <div class="muted">${esc(data.remoteCommit || "")}</div>
        <div class="tiny muted ellipsis" title="${esc(data.targetPath || "")}">${esc(data.targetPath || "")}</div>
      </div>
      <div class="row">
        <button data-act="toggle-ai">${state.aiOpen ? "收起 AI 合并" : "AI 合并"}</button>
        <button data-act="open-settings">设置中心</button>
      </div>
    </div>
    <div class="options wrap">
      <span>基准侧</span>
      <select data-act="base-side" style="width:140px">
        <option value="local" ${options.baseSide !== "remote" ? "selected" : ""}>本地 Local</option>
        <option value="remote" ${options.baseSide === "remote" ? "selected" : ""}>线上 Remote</option>
      </select>
      ${flag("跳过新增行", "skipNewRows", options.skipNewRows)}
      ${flag("跳过新增列", "skipNewColumns", options.skipNewColumns)}
      ${flag("删除缺失行", "deleteMissingRows", options.deleteMissingRows)}
      ${flag("删除缺失列", "deleteMissingColumns", options.deleteMissingColumns)}
      ${flag("追加新 Sheet", "addNewSheets", options.addNewSheets)}
      ${flag("删除缺失 Sheet", "deleteMissingSheets", options.deleteMissingSheets)}
      ${flag("处理冲突", "resolveConflicts", options.resolveConflicts)}
    </div>
    <div class="work">
      <div class="main">
        <div class="sheetbar">
          <div class="grow">
            <div class="row"><span>当前 Sheet：</span><strong class="ellipsis">${esc(wizard.current || "")}</strong><span class="muted">${esc(wizard.progress || "")}</span></div>
          </div>
          <input id="merge-q" type="search" placeholder="搜索键或工作表" value="${esc(state.q)}" style="width:220px">
          <button data-act="sheet-prev" ${wizard.canPrevious ? "" : "disabled"}>上一个 Sheet</button>
          <button data-act="sheet-next" ${wizard.canNext ? "" : "disabled"}>下一个 Sheet</button>
          <button data-act="sheet-finish" ${wizard.canFinish ? "" : "disabled"}>确认本 Sheet</button>
        </div>
        <div class="summary muted">${esc(data.summaryText || "")}</div>
        <div class="summary muted">${esc(wizard.sheetStatus || "")}</div>
        <div class="table-wrap"><table class="grid">
          <thead><tr><th style="width:96px">类型</th><th style="width:140px">工作表</th><th>键</th><th>处理</th><th style="width:156px">选择</th></tr></thead>
          <tbody>${(data.items || []).map((row, index) => `<tr class="${index === state.pick ? "selected" : ""}" data-act="pick" data-index="${index}">
            <td><span class="tag tag-${esc(row.tag)}">${esc(row.typeLabel)}</span></td>
            <td class="ellipsis">${esc(row.sheet)}</td>
            <td class="ellipsis">${esc(row.key)}</td>
            <td class="ellipsis">${esc(row.action)}</td>
            <td>${row.canChoose ? `<button class="choice ${row.choice === "local" ? "on" : ""}" data-act="choice" data-index="${index}" data-choice="local">本地</button>
              <button class="choice ${row.choice === "remote" ? "on" : ""}" data-act="choice" data-index="${index}" data-choice="remote">线上</button>` : ""}</td>
          </tr>`).join("")}</tbody>
        </table></div>
        ${pager(data)}
        ${item ? `<div class="panes">
          ${pane("BASE", item.baseValues)}
          ${pane("LOCAL", item.localValues)}
          ${pane("REMOTE", item.remoteValues)}
        </div>` : `<div class="panes"><div class="hint">选择上方的一行，查看 BASE、LOCAL、REMOTE。</div></div>`}
        <div class="footer">
          <div class="grow ellipsis">${esc(data.status || "")}</div>
          <button data-act="cancel">取消</button>
          <button data-act="generate" ${state.busy ? "disabled" : ""}>生成合并结果</button>
          <button class="primary" data-act="confirm" ${data.canConfirm && !state.busy ? "" : "disabled"}>确认并解决冲突</button>
        </div>
      </div>
      ${state.aiOpen ? renderAi(data.ai || {}) : ""}
    </div>
  </div>`;
}

function flag(label, key, checked) {
  return `<label class="check"><input type="checkbox" data-act="option" data-key="${key}" ${checked ? "checked" : ""}> ${label}</label>`;
}

function pane(title, lines) {
  const body = (lines || []).map((line) => `<p>${esc(line)}</p>`).join("") || `<p class="muted">无</p>`;
  return `<section class="pane"><h2>${title}</h2>${body}</section>`;
}

function pager(data) {
  const total = data.itemTotal || 0;
  const limit = data.limit || 500;
  if (total <= limit) return "";
  const page = Math.floor((data.offset || 0) / limit) + 1;
  const pages = Math.max(1, Math.ceil(total / limit));
  return `<div class="pager"><button data-act="page" data-dir="-1" ${page <= 1 ? "disabled" : ""}>上一页</button><span>${page} / ${pages}</span><button data-act="page" data-dir="1" ${page >= pages ? "disabled" : ""}>下一页</button></div>`;
}

function renderAi(ai) {
  const files = (ai.files || []).map((file) => `<div class="file">
      <label><input type="checkbox" data-act="ai-check" data-path="${esc(file.fullPath)}" ${file.isChecked ? "checked" : ""}><span>
        <strong>${esc(file.relativePath)}</strong>
        <div class="tiny muted">${esc(file.fullPath)}</div>
        ${file.extraPathText ? `<pre class="paths">${esc(file.extraPathText)}</pre>` : ""}
        <div class="tiny muted">${esc(file.statusLabel)}  ${esc(file.sidesText)}</div>
      </span></label>
    </div>`).join("");
  return `<aside class="ai">
    <header>
      <h1 style="font-size:18px">AI 合并</h1>
      <p class="muted">勾选当前冲突仓库里要交给 AI 的 Excel。每条下面是文件路径。</p>
      <p class="muted">${esc(ai.repoText || "")}</p>
      <p>${esc(ai.summaryText || "")}</p>
    </header>
    <div class="list">${files}</div>
    <footer>
      <p class="tiny muted">复制后可交给 AI。窗口保持打开，可以继续合并。</p>
      <p class="tiny muted">${esc(ai.statusText || "")}</p>
      <div class="wrap">
        <button data-act="ai-refresh">刷新</button>
        <button data-act="ai-all" data-value="true">全选</button>
        <button data-act="ai-all" data-value="false">全不选</button>
        <button data-act="ai-prompt" data-detail="false">复制路径给 AI</button>
        <button data-act="ai-prompt" data-detail="true">复制冲突给 AI</button>
      </div>
    </footer>
  </aside>`;
}

function renderCompare() {
  const data = state.compare;
  if (!data) return `<div class="screen pad"><p>正在打开对比...</p></div>`;
  const filter = state.settings.diffFilter || {};
  const sheets = (data.sheets || []).map((sheet) =>
    `<button class="sheet ${sheet.current ? "on" : ""}" data-act="compare-sheet" data-sheet="${sheet.sheet == null ? "*" : esc(sheet.sheet)}">
      <span class="ellipsis">${esc(sheet.name)}</span><span class="tiny muted">${esc(sheet.badge)}</span></button>`).join("");
  const rows = (data.rows || []).map((row) => `<tr class="${state.comparePick === row.index ? "selected" : ""}" data-act="compare-pick" data-index="${row.index}">
      <td><span class="tag ${statusClass(row.status)}">${esc(row.status)}</span></td>
      <td class="ellipsis">${esc(row.key)}</td>
      <td class="ellipsis">${esc(row.location || "")}</td>
      <td class="ellipsis">${esc(row.preview || "")}</td>
    </tr>`).join("");
  return `<div class="screen">
    <div class="pad head">
      <div class="text">
        <h1>本地 对比 线上</h1>
        <p class="muted">在左侧选择工作表，逐张查看差异。重复的主键会按出现顺序逐条对比。</p>
      </div>
      <div class="row">
        <button data-act="swap">互换基准</button>
        <button data-act="open-settings">设置中心</button>
      </div>
    </div>
    <div class="cards">
      <div class="card"><div class="tiny muted">本地</div><strong class="ellipsis">${esc(data.localLabel || "")}</strong><div class="tiny muted ellipsis" title="${esc(data.localPath || "")}">${esc(data.localPath || "")}</div></div>
      <div class="card"><div class="tiny muted">线上</div><strong class="ellipsis">${esc(data.remoteLabel || "")}</strong><div class="tiny muted ellipsis" title="${esc(data.remotePath || "")}">${esc(data.remotePath || "")}</div></div>
    </div>
    <div class="work">
      <aside class="sheets">
        <h2>工作表</h2>
        <div class="list">${sheets}</div>
        ${data.skippedHint ? `<p class="tiny muted" style="margin:8px">${esc(data.skippedHint)}</p>` : ""}
      </aside>
      <div class="main">
        <div class="filters">
          <div class="checks">
            ${diffCheck("新增行", filter)}
            ${diffCheck("删除行", filter)}
            ${diffCheck("新增列", filter)}
            ${diffCheck("删除列", filter)}
            ${diffCheck("修改", filter)}
          </div>
          <input id="compare-q" type="search" placeholder="搜索键、工作表或单元格" value="${esc(state.compareQ)}" style="width:220px">
        </div>
        <div class="table-wrap">${data.showEmpty ? `<div class="empty"><h2>${esc(data.emptyTitle || "")}</h2><p class="muted">${esc(data.emptyMessage || "")}</p></div>` : `<table class="grid diff">
          <thead><tr><th style="width:96px">状态</th><th>键</th><th style="width:180px">位置</th><th>差异</th></tr></thead>
          <tbody>${rows}</tbody></table>`}</div>
        ${comparePager(data)}
        <div class="pane" style="flex:1;margin-top:8px">${renderDetail()}</div>
        <div class="footer">
          <div class="grow ellipsis">${esc(data.status || "")}</div>
          <label class="check"><input type="checkbox" data-act="auto-open" ${state.settings.autoOpenCompare ? "checked" : ""}> 导出后打开</label>
          <button class="primary" data-act="export">导出对比工作簿</button>
          <button data-act="cancel">关闭</button>
        </div>
      </div>
    </div>
  </div>`;
}

function diffCheck(name, filter) {
  const on = filter[name] !== false;
  return `<label class="check"><input type="checkbox" data-act="diff-filter" data-name="${esc(name)}" ${on ? "checked" : ""}> ${esc(name)}</label>`;
}

function statusClass(status) {
  if (status === "新增行" || status === "新增列") return "tag-status-add";
  if (status === "删除行" || status === "删除列") return "tag-status-del";
  return "tag-status-mod";
}

function comparePager(data) {
  const total = data.itemTotal || 0;
  const limit = data.limit || 500;
  if (total <= limit) return "";
  const page = Math.floor((data.offset || 0) / limit) + 1;
  const pages = Math.max(1, Math.ceil(total / limit));
  return `<div class="pager"><button data-act="compare-page" data-dir="-1" ${page <= 1 ? "disabled" : ""}>上一页</button><span>${page} / ${pages}</span><button data-act="compare-page" data-dir="1" ${page >= pages ? "disabled" : ""}>下一页</button></div>`;
}

function renderDetail() {
  const detail = state.detail;
  if (!detail) return `<div class="hint">选择上方的一行，按列查看本地和线上的差异。</div>`;
  const cells = (detail.cells || []).map((cell) => `<div class="cell"><b>${esc(cell.header)}</b></div><div class="cell">${esc(cell.right)}</div><div class="cell">${esc(cell.left)}</div>`).join("");
  return `<div class="detail-head">
      <div><strong>${esc(detail.status)}  ${esc(detail.key)}</strong><div class="tiny muted">${esc(detail.location || "")}</div></div>
      <label class="check"><input type="checkbox" data-act="all-cells" ${state.showAllCells ? "checked" : ""}> 显示全部列</label>
    </div>
    <div class="cols"><b class="muted">列</b><b>本地</b><b>线上</b>${cells}</div>`;
}

function renderChoose() {
  const choice = (state.boot && state.boot.choice) || {};
  const rows = (choice.files || []).map((file) => `<tr>
    <td>${esc(file.role)}</td><td class="ellipsis">${esc(file.name)}</td><td>${esc(file.size)}</td><td class="ellipsis" title="${esc(file.path)}">${esc(file.path)}</td></tr>`).join("");
  return `<div class="screen">
    <div class="pad">
      <h1>${esc(choice.title || "")}</h1>
      <p class="muted">${esc(choice.notice || "")}</p>
      ${choice.warn ? `<p class="warn">${esc(choice.warn)}</p>` : ""}
    </div>
    <div class="main" style="flex:1">
      <div class="table-wrap"><table class="grid"><thead><tr><th style="width:120px">角色</th><th style="width:180px">文件</th><th style="width:80px">大小</th><th>路径</th></tr></thead><tbody>${rows}</tbody></table></div>
      <div class="footer between">
        <button data-act="choice-cancel">取消</button>
        <div class="row"><button data-act="boot-choice" data-choice="backup_only">快速备份</button><button class="primary" data-act="boot-choice" data-choice="merge_diff" ${choice.canMerge ? "" : "disabled"}>进入合并/对比</button></div>
      </div>
    </div>
  </div>`;
}

function renderBackup() {
  const backup = (state.boot && state.boot.backup) || {};
  const rows = (backup.files || []).map((file) => `<tr><td>${esc(file.role)}</td><td class="ellipsis" title="${esc(file.path)}">${esc(file.path)}</td></tr>`).join("");
  return `<div class="screen">
    <div class="pad">
      <h1>${esc(backup.headline || "备份")}</h1>
      <p class="muted">${esc(backup.detail || "")}</p>
    </div>
    <div class="main" style="flex:1">
      <div class="table-wrap"><table class="grid"><thead><tr><th style="width:140px">类型</th><th>备份文件</th></tr></thead><tbody>${rows}</tbody></table></div>
      <div class="footer" style="justify-content:flex-end">
        <button class="primary" data-act="open-backup-dir" ${backup.ok ? "" : "disabled"}>打开目录</button>
        <button data-act="cancel">关闭</button>
      </div>
    </div>
  </div>`;
}

async function onClick(event) {
  const node = event.target.closest("[data-act]");
  if (!node || node.disabled) return;
  const act = node.dataset.act;
  try {
    if (act === "nav") return openNav(node.dataset.nav);
    if (act === "back") return back();
    if (act === "copy") return copyText(node.dataset.text || "");
    if (act === "pick-backup") {
      if (!(window.chrome && window.chrome.webview)) {
        const input = document.getElementById("backup-root");
        if (input) input.focus();
        return;
      }
      return hostMessage("pick-folder");
    }
    if (act === "save-backup") return saveBackup();
    if (act === "open-backup") return openBackupRoot();
    if (act === "open-backup-dir") return api("/api/open", { method: "POST", body: JSON.stringify({ path: state.boot.backup.dir }) });
    if (act === "fork-install") return refreshIntegrations("/api/integrations/fork/install");
    if (act === "fork-uninstall") return refreshIntegrations("/api/integrations/fork/uninstall");
    if (act === "git-install") return refreshIntegrations("/api/integrations/git/install");
    if (act === "git-uninstall") return refreshIntegrations("/api/integrations/git/uninstall");
    if (act === "check-update") return checkUpdate();
    if (act === "fixture-merge") return fillFixture("merges", node.dataset.id);
    if (act === "fixture-compare") return fillFixture("compares", node.dataset.id);
    if (act === "open-merge") return openMerge(state.fields.local, state.fields.base, state.fields.remote, state.fields.merged, state.fields.mode || null);
    if (act === "open-compare") return openCompare(state.fields.left, state.fields.right);
    if (act === "open-settings") return openSettings();
    if (act === "toggle-ai") { state.aiOpen = !state.aiOpen; render(); return; }
    if (act === "sheet-prev") return postMerge("/sheets/previous");
    if (act === "sheet-next") return postMerge("/sheets/next");
    if (act === "sheet-finish") return postMerge("/sheets/complete");
    if (act === "pick") { state.pick = Number(node.dataset.index); render(); return; }
    if (act === "choice") return chooseRow(Number(node.dataset.index), node.dataset.choice);
    if (act === "page") return turnPage(Number(node.dataset.dir));
    if (act === "generate") return generate();
    if (act === "confirm") return confirmMerge();
    if (act === "cancel") return cancel();
    if (act === "ai-refresh") return postMerge("/ai/refresh");
    if (act === "ai-all") return postMergeBody("/ai/checks", { all: node.dataset.value === "true" });
    if (act === "ai-prompt") return copyPrompt(node.dataset.detail === "true");
    if (act === "swap") return postCompare("/swap");
    if (act === "compare-sheet") return selectCompareSheet(node.dataset.sheet);
    if (act === "compare-pick") return loadDetail(Number(node.dataset.index));
    if (act === "compare-page") return turnComparePage(Number(node.dataset.dir));
    if (act === "export") return exportCompare(false);
    if (act === "boot-choice") return chooseBoot(node.dataset.choice);
    if (act === "choice-cancel") return cancelBoot();
  } catch (error) {
    showError(error);
  }
}

async function onChange(event) {
  const node = event.target;
  try {
    if (node.dataset.act === "startup") return saveSettings({ startupFeature: node.dataset.value });
    if (node.dataset.act === "dark") {
      document.body.classList.toggle("dark", node.checked);
      return saveSettings({ darkTheme: node.checked });
    }
    if (node.dataset.act === "option") return saveOption(node.dataset.key, node.checked);
    if (node.dataset.act === "base-side") return saveOption("baseSide", node.value);
    if (node.dataset.act === "ai-check") return postMergeBody("/ai/checks", { path: node.dataset.path, checked: node.checked });
    if (node.dataset.act === "diff-filter") return saveDiffFilter(node.dataset.name, node.checked);
    if (node.dataset.act === "auto-open") return saveSettings({ autoOpenCompare: node.checked });
    if (node.dataset.act === "all-cells") {
      state.showAllCells = node.checked;
      if (state.comparePick != null) return loadDetail(state.comparePick);
    }
  } catch (error) {
    showError(error);
  }
}

function onInput(event) {
  const node = event.target;
  if (node.dataset.field) {
    state.fields[node.dataset.field] = node.value;
    if (node.dataset.field === "local" || node.dataset.field === "base" || node.dataset.field === "remote") {
      state.fields.mode = null;
    }
    return;
  }
  if (node.id === "merge-q") {
    state.q = node.value;
    focusRestore = { id: "merge-q", start: node.selectionStart, end: node.selectionEnd };
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => reloadMerge().catch(showError), 120);
  }
  if (node.id === "compare-q") {
    state.compareQ = node.value;
    focusRestore = { id: "compare-q", start: node.selectionStart, end: node.selectionEnd };
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => reloadCompare().catch(showError), 120);
  }
}

async function openNav(nav) {
  state.nav = nav;
  if ((nav === "fork" || nav === "git") && !state.integrations) {
    render();
    state.integrations = await api("/api/integrations");
  }
  render();
}

function openSettings() {
  state.back = state.screen;
  state.screen = "settings";
  state.nav = "general";
  if (!state.fixtures) api("/api/fixtures").then((data) => { state.fixtures = data; if (state.screen === "settings") render(); });
  render();
}

function back() {
  state.screen = state.back || "settings";
  state.back = null;
  render();
}

async function saveSettings(patch) {
  state.settings = await api("/api/settings", { method: "PUT", body: JSON.stringify(patch) });
  document.body.classList.toggle("dark", !!state.settings.darkTheme);
  if (state.screen === "settings" || state.screen === "compare") render();
}

async function saveBackup() {
  await saveSettings({ backupRootDir: state.fields.backup || "" });
  alert("备份根目录已保存。");
}

async function openBackupRoot() {
  await api("/api/open", { method: "POST", body: JSON.stringify({ path: state.fields.backup || "" }) });
}

async function refreshIntegrations(path) {
  state.integrations = await api(path, { method: "POST" });
  render();
}

async function checkUpdate() {
  state.busy = true;
  state.updateText = "正在检查...";
  render();
  try {
    const info = await api("/api/update/check");
    state.updateText = info.message;
    if (!info.found) {
      alert("最新 Release 未找到 ExcelMergeFork.exe。");
      return;
    }
    if (!info.isNewer) {
      alert("当前已是最新版本 v" + info.current + "。");
      return;
    }
    const text = "发现新版本 v" + info.version + "，当前版本 v" + info.current + "。\n\n"
      + "点击确定后会下载新版 exe；下载完成后需要关闭当前窗口，工具会在退出后替换文件。\n"
      + "如果正在合并冲突，建议先完成或取消当前合并后再更新。";
    if (!confirm(text)) return;
    state.updateText = "下载中...";
    render();
    const applied = await api("/api/update/apply", { method: "POST", body: "{}" });
    alert(applied.message);
    if (applied.close) hostMessage("close");
  } catch (error) {
    state.updateText = "检查失败";
    throw error;
  } finally {
    state.busy = false;
    render();
  }
}

function fillFixture(kind, id) {
  const item = ((state.fixtures && state.fixtures[kind]) || []).find((row) => row.id === id);
  if (!item) return;
  if (kind === "merges") {
    state.fields.local = item.local;
    state.fields.base = item.base;
    state.fields.remote = item.remote;
    state.fields.mode = item.mode;
  } else {
    state.fields.left = item.left;
    state.fields.right = item.right;
  }
  render();
}

async function openMerge(local, base, remote, merged, mode) {
  state.merge = await api("/api/merge/sessions?limit=500", {
    method: "POST",
    body: JSON.stringify({ local, base, remote, merged: merged || null, mode }),
  });
  state.q = "";
  state.offset = 0;
  state.pick = 0;
  state.aiOpen = true;
  state.screen = "merge";
  state.back = null;
  render();
}

async function openCompare(left, right) {
  state.compare = await api("/api/compare/sessions?limit=500", {
    method: "POST",
    body: JSON.stringify({ left, right }),
  });
  state.compareSheet = state.compare.activeSheet;
  state.compareQ = "";
  state.compareOffset = 0;
  state.comparePick = null;
  state.detail = null;
  state.screen = "compare";
  state.back = null;
  render();
  if (state.settings.autoOpenCompare && state.compare.total > 0) await exportCompare(true);
}

async function reloadMerge() {
  if (!state.merge) return;
  state.offset = 0;
  state.merge = await api(mergeUrl(state.merge.id));
  if (state.pick >= (state.merge.items || []).length) state.pick = 0;
  render();
}

async function postMerge(path) {
  state.merge = await api("/api/merge/sessions/" + state.merge.id + path + query(state.q, state.offset), { method: "POST" });
  state.pick = 0;
  render();
}

async function postMergeBody(path, body) {
  state.merge = await api("/api/merge/sessions/" + state.merge.id + path + query(state.q, state.offset), {
    method: "POST",
    body: JSON.stringify(body),
  });
  render();
}

async function saveOption(key, value) {
  const options = Object.assign({}, state.merge.options, { [key]: value });
  state.merge = await api("/api/merge/sessions/" + state.merge.id + "/options" + query(state.q, 0), {
    method: "PUT",
    body: JSON.stringify(options),
  });
  state.offset = 0;
  state.pick = 0;
  render();
}

async function chooseRow(index, choice) {
  const row = state.merge.items[index];
  state.pick = index;
  state.merge = await api("/api/merge/sessions/" + state.merge.id + "/choice" + query(state.q, state.offset), {
    method: "POST",
    body: JSON.stringify({ sheet: row.conflictSheet, key: row.conflictKey, kind: row.kind, choice }),
  });
  render();
}

async function turnPage(dir) {
  const limit = state.merge.limit || 500;
  state.offset = Math.max(0, (state.merge.offset || 0) + dir * limit);
  state.merge = await api(mergeUrl(state.merge.id));
  state.pick = 0;
  render();
}

async function generate() {
  state.busy = true;
  render();
  try {
    state.merge = await api("/api/merge/sessions/" + state.merge.id + "/generate" + query(state.q, state.offset), {
      method: "POST",
      body: JSON.stringify({ createBackup: true, merged: state.merge.targetPath || null }),
    });
    render();
  } catch (error) {
    alert(error.message);
  } finally {
    state.busy = false;
    render();
  }
}

async function confirmMerge() {
  const result = await api("/api/merge/sessions/" + state.merge.id + "/confirm", { method: "POST", body: "{}" });
  if (result.message) state.merge.status = result.message;
  render();
  if (result.close) hostMessage("close");
  else if (result.message) alert(result.message);
}

async function copyPrompt(detail) {
  const result = await api("/api/merge/sessions/" + state.merge.id + "/ai/prompt", {
    method: "POST",
    body: JSON.stringify({ detail }),
  });
  await copyText(result.text);
  if (state.merge.ai) state.merge.ai.statusText = "已复制到剪贴板";
  render();
}

async function reloadCompare() {
  if (!state.compare) return;
  state.compareOffset = 0;
  state.compare = await api(compareUrl(state.compare.id));
  render();
}

async function postCompare(path) {
  state.detail = null;
  state.comparePick = null;
  state.compare = await api("/api/compare/sessions/" + state.compare.id + path + compareQuery(), { method: "POST" });
  state.compareSheet = state.compare.activeSheet;
  render();
}

async function selectCompareSheet(sheet) {
  state.compareSheet = sheet === "*" ? "" : sheet;
  state.compareOffset = 0;
  state.detail = null;
  state.comparePick = null;
  state.compare = await api(compareUrl(state.compare.id));
  render();
}

async function saveDiffFilter(name, checked) {
  const diffFilter = Object.assign({}, state.settings.diffFilter, { [name]: checked });
  await saveSettings({ diffFilter });
  state.compareOffset = 0;
  state.compare = await api(compareUrl(state.compare.id));
  render();
}

async function turnComparePage(dir) {
  const limit = state.compare.limit || 500;
  state.compareOffset = Math.max(0, (state.compare.offset || 0) + dir * limit);
  state.compare = await api(compareUrl(state.compare.id));
  render();
}

async function loadDetail(index) {
  state.comparePick = index;
  const changed = state.showAllCells ? "false" : "true";
  state.detail = await api("/api/compare/sessions/" + state.compare.id + "/rows/" + index + "?changedOnly=" + changed);
  state.detail.location = locationOf(state.detail.leftRow, state.detail.rightRow);
  render();
}

async function exportCompare(auto) {
  const exported = await api("/api/compare/sessions/" + state.compare.id + "/export", { method: "POST" });
  state.compare.exportPath = exported.path;
  state.compare.status = "已导出 " + exported.path;
  render();
  if (state.settings.autoOpenCompare || auto) {
    await api("/api/open", { method: "POST", body: JSON.stringify({ path: exported.path }) });
  }
}

async function chooseBoot(choice) {
  state.boot = await api("/api/boot/choose", { method: "POST", body: JSON.stringify({ choice }) });
  state.screen = state.boot.screen;
  if (state.boot.mergeId) state.merge = await api(mergeUrl(state.boot.mergeId));
  if (state.boot.compareId) {
    state.compare = await api(compareUrl(state.boot.compareId));
    state.compareSheet = state.compare.activeSheet;
  }
  render();
}

async function cancelBoot() {
  const result = await api("/api/boot/cancel", { method: "POST", body: "{}" });
  if (result.close) hostMessage("close");
}

async function cancel() {
  if (state.boot && state.boot.host === "desktop" && !state.back) {
    hostMessage("close");
    return;
  }
  if (state.merge && state.screen === "merge") {
    await api("/api/merge/sessions/" + state.merge.id, { method: "DELETE" });
    state.merge = null;
  }
  if (state.compare && state.screen === "compare") {
    await api("/api/compare/sessions/" + state.compare.id, { method: "DELETE" });
    state.compare = null;
  }
  state.screen = "settings";
  state.back = null;
  if (!state.fixtures) state.fixtures = await api("/api/fixtures");
  render();
}

function mergeUrl(id) {
  return "/api/merge/sessions/" + id + query(state.q, state.offset);
}

function compareUrl(id) {
  return "/api/compare/sessions/" + id + compareQuery();
}

function query(q, offset) {
  const params = new URLSearchParams();
  if (q) params.set("q", q);
  if (offset) params.set("offset", String(offset));
  params.set("limit", "500");
  return "?" + params.toString();
}

function compareQuery() {
  const params = new URLSearchParams();
  if (state.compareSheet !== undefined) params.set("sheet", state.compareSheet || "");
  if (state.compareQ) params.set("q", state.compareQ);
  if (state.compareOffset) params.set("offset", String(state.compareOffset));
  params.set("limit", "500");
  const filter = (state.settings && state.settings.diffFilter) || {};
  Object.keys(filter).forEach((name) => {
    if (filter[name]) params.append("status", name);
  });
  if (!params.has("status")) params.append("status", "");
  return "?" + params.toString();
}

function locationOf(remoteRow, localRow) {
  if (localRow > 0 && remoteRow > 0 && localRow !== remoteRow) return "本地 " + localRow + " 行 · 线上 " + remoteRow + " 行";
  if (localRow > 0 && remoteRow > 0) return "第 " + localRow + " 行";
  if (localRow > 0) return "本地第 " + localRow + " 行";
  if (remoteRow > 0) return "线上第 " + remoteRow + " 行";
  return "";
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = document.createElement("textarea");
    area.value = text;
    document.body.appendChild(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
}

function hostMessage(text) {
  if (window.chrome && window.chrome.webview) window.chrome.webview.postMessage(text);
}

function restoreFocus() {
  if (!focusRestore) return;
  const el = document.getElementById(focusRestore.id);
  focusRestore = null;
  if (!el) return;
  el.focus();
  try { el.setSelectionRange(el.value.length, el.value.length); } catch { /* number inputs */ }
}

function showError(error) {
  alert(error && error.message ? error.message : String(error));
}

async function api(path, options) {
  const response = await fetch(path, Object.assign({ headers: { "Content-Type": "application/json" } }, options));
  const text = await response.text();
  let data = {};
  if (text) {
    try { data = JSON.parse(text); } catch { data = { error: text }; }
  }
  if (!response.ok) throw new Error(data.error || text || String(response.status));
  return data;
}

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[ch]));
}
