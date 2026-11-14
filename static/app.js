const el = id => document.getElementById(id);
const form = el("new-task");
const files = [];
let options;
let activePreset;
let task = null;
let pollTimer;
let submitting = false;
let reviewSubmitting = false;

const storageKey = "roster-harness-task";
const historyDate = new Intl.DateTimeFormat("zh-CN", {dateStyle: "medium", timeStyle: "short"});
const fileKinds = {
  xlsx: "sheet", xls: "sheet", et: "et", docx: "word", doc: "word", wps: "wps",
  pdf: "pdf", png: "image", jpg: "image", jpeg: "image", webp: "image",
};
const toolNames = {
  list_documents: "查看资料目录",
  read_document: "读取资料",
  submit_roster: "提交名单",
};

function node(tag, text, className) {
  const element = document.createElement(tag);
  if (text !== undefined) element.textContent = text;
  if (className) element.className = className;
  return element;
}

function message(id, text = "") {
  const element = el(id);
  element.textContent = text;
  element.hidden = !text;
}

async function api(path, init) {
  const response = await fetch(path, init);
  let payload = {};
  try { payload = await response.json(); } catch {}
  if (!response.ok) {
    const detail = payload.detail;
    throw new Error(typeof detail === "string" ? detail : "请求未完成，请检查资料或稍后重试。");
  }
  return payload;
}

function remember(id) {
  try {
    if (id) localStorage.setItem(storageKey, id);
    else localStorage.removeItem(storageKey);
  } catch {}
  const url = new URL(location.href);
  if (id) url.searchParams.set("task", id);
  else url.searchParams.delete("task");
  history.replaceState(null, "", url);
}

function updateFileCount() {
  const selected = files.filter(entry => entry.selected).length;
  el("file-count").textContent = files.length ? `已选 ${selected} / ${files.length} 份` : "已选 0 份";
  updateStart();
}

function updateStart() {
  const title = form.elements.title.value.trim();
  const selected = files.some(entry => entry.selected);
  el("start").disabled = submitting || !title || !selected || !activePreset;
}

function addFiles(list) {
  const rejected = [];
  let total = files.reduce((sum, entry) => sum + entry.file.size, 0);
  for (const file of list) {
    const extension = file.name.split(".").at(-1).toLowerCase();
    if (!options.upload.extensions.includes("." + extension)) {
      rejected.push(`${file.name} 的格式不受支持`);
      continue;
    }
    if (file.size > options.upload.max_file_bytes) {
      rejected.push(`${file.name} 超过单份 300 MB`);
      continue;
    }
    if (total + file.size > options.upload.max_total_bytes) {
      rejected.push(`${file.name} 使整批超过 3 GB`);
      continue;
    }
    if (files.some(entry => entry.file.name === file.name && entry.file.size === file.size &&
        entry.file.lastModified === file.lastModified)) continue;
    files.push({file, selected: true});
    total += file.size;
  }
  if (files.length > options.upload.max_files) {
    while (files.length > options.upload.max_files) files.pop();
    rejected.push(`一次最多 ${options.upload.max_files} 份`);
  }
  message("upload-note", rejected.join("；"));
  renderFiles();
}

function renderFiles() {
  const list = el("file-list");
  list.replaceChildren(...files.map((entry, index) => {
    const item = node("li");
    const choice = node("label", "", "file-choice");
    const check = node("input");
    check.type = "checkbox";
    check.checked = entry.selected;
    check.addEventListener("change", () => {
      entry.selected = check.checked;
      updateFileCount();
    });
    choice.append(check, node("span", entry.file.name, "file-name"));
    const extension = entry.file.name.split(".").at(-1).toLowerCase();
    const meta = node("span", "", "file-meta");
    meta.append(node("span", extension.toUpperCase(), "file-type " + (fileKinds[extension] || "file")));
    const remove = node("button", "×", "file-remove");
    remove.type = "button";
    remove.setAttribute("aria-label", `移除 ${entry.file.name}`);
    remove.addEventListener("click", () => {
      files.splice(index, 1);
      renderFiles();
      updateFileCount();
    });
    item.append(choice, meta, remove);
    return item;
  }));
  el("empty-upload").hidden = files.length > 0;
  updateFileCount();
}

function selectPreset(id) {
  activePreset = id;
  for (const radio of el("presets").querySelectorAll("input")) radio.checked = radio.value === id;
  const preset = options.presets.find(item => item.id === id);
  if (!preset) return;
  el("preview-name").textContent = preset.name;
  const sheet = el("preset-sample");
  const image = node("img");
  image.src = `/static/previews/${encodeURIComponent(id)}.png`;
  image.alt = `${preset.name}示例`;
  image.loading = "lazy";
  sheet.replaceChildren(image);
  updateStart();
}

function renderPresets(presets) {
  const root = el("presets");
  root.replaceChildren(...presets.map((preset, index) => {
    const label = node("label", "", "preset");
    const radio = node("input", "", "preset-radio");
    radio.type = "radio";
    radio.name = "preset";
    radio.value = preset.id;
    radio.checked = index === 0;
    radio.addEventListener("change", () => selectPreset(preset.id));
    label.append(radio, node("span", preset.name));
    label.title = preset.description;
    return label;
  }));
  el("template-picker").hidden = false;
  selectPreset(presets[0]?.id);
}

function phaseOf(value) {
  if (!value) return "idle";
  if (value.status === "review") return "review";
  return value.status || "idle";
}

function stateLabel(value) {
  const phase = phaseOf(value);
  return {
    queued: "准备中", running: "整理中", review: "待核对", done: "已完成", error: "未完成", stopped: "已结束",
  }[phase] || "准备中";
}

function setTaskDetail(value) {
  const preset = value.preset_name || value.preset || "";
  const files = value.files || [];
  const done = files.filter(file => file.status === "done").length;
  const progress = files.length > 1 ? `${done}/${files.length} 份完成` : "";
  el("task-detail").textContent = [value.title, value.group, preset, progress].filter(Boolean).join(" · ");
}

function renderTask(value) {
  task = value;
  el("activity").hidden = false;
  el("activity").dataset.phase = phaseOf(value);
  el("task-state-label").textContent = stateLabel(value);
  setTaskDetail(value);
  const phase = phaseOf(value);
  const busy = ["queued", "running"].includes(phase);
  const progress = el("task-progress");
  progress.hidden = !busy && phase !== "done";
  progress.classList.toggle("indeterminate", busy);
  progress.style.setProperty("--progress", phase === "done" ? "100%" : "28%");
  const percent = phase === "done" ? 100 : busy ? null : 0;
  if (percent === null) progress.removeAttribute("aria-valuenow");
  else progress.setAttribute("aria-valuenow", String(percent));

  const word = el("word");
  word.hidden = !value.word_available;
  if (value.word_available) word.href = `/api/tasks/${encodeURIComponent(value.id)}/download/word`;
  el("restart-task").hidden = value.status !== "error";

  message("task-error", value.error || "");
  renderFilesProgress(value);
  renderReview(value);
  renderTrace(value);
  updateStart();
}

function renderFilesProgress(value) {
  const section = el("batch-progress");
  const list = el("batch-files");
  const sourceFiles = value.files || [];
  section.hidden = !sourceFiles.length;
  const trace = value.trace || [];
  const read = new Set(trace.filter(event => event.name === "read_document" && event.detail)
    .map(event => String(event.detail).split(" · ")[0]));
  list.replaceChildren(...sourceFiles.map(file => {
    const item = node("li");
    const copy = node("div", "", "batch-copy");
    copy.append(node("strong", file.name));
    const units = file.units || [];
    const detail = file.suffix === ".pdf" ? `${units.length} 页` :
      units.length === 1 ? `${units[0].label}` : `${units.length} 个工作表`;
    const states = {waiting: "待处理", running: "识别中", done: "已完成", error: "失败", review: "待核对"};
    const state = states[file.status] || (read.has(file.id) ? "已读取" : "待读取");
    copy.append(node("small", `${file.id} · ${detail} · ${state}`, "quiet"));
    if (file.error) copy.append(node("small", file.error, "event-error"));
    item.append(copy);
    if (file.status === "error" && value.status === "error") {
      const actions = node("div", "", "batch-actions");
      const retry = node("button", "重试此文件", "secondary");
      retry.type = "button";
      retry.addEventListener("click", async () => {
        retry.disabled = true;
        try {
          const updated = await api(`/api/tasks/${encodeURIComponent(value.id)}/files/${encodeURIComponent(file.id)}/retry`, {method: "POST"});
          task = updated;
          renderTask(updated);
          poll(updated.id);
        } catch (error) {
          message("task-error", error.message);
          retry.disabled = false;
        }
      });
      actions.append(retry);
      item.append(actions);
    }
    return item;
  }));
}

function reviewSource(value) {
  return (value.review?.sources || [])[0] || "";
}

function sourceParts(value, source) {
  const file = (value.files || []).find(item => source.startsWith(item.id + " "));
  if (!file) return {file: null, name: "", page: null, location: source};
  const pageMatch = source.match(/第 (\d+) 页/);
  return {file, name: file.name, page: pageMatch ? Number(pageMatch[1]) : null, location: source};
}

function renderReview(value) {
  const section = el("review");
  const review = value.review;
  section.hidden = !review;
  if (!review) return;
  el("review-question").textContent = review.question || "";
  const source = reviewSource(value);
  const parts = sourceParts(value, source);
  el("review-context").hidden = !source;
  el("review-filename").textContent = parts.name || "";
  el("review-location").textContent = source || "";
  el("source-open").hidden = !parts.file;
  el("source-open").onclick = () => openSources(value);
  const evidence = el("review-evidence");
  const list = el("review-evidence-list");
  const sources = review.sources || [];
  evidence.hidden = !sources.length;
  list.replaceChildren(...sources.map(item => node("li", item)));
  const options = el("review-options");
  const buttons = el("review-option-list");
  options.hidden = !(review.options || []).length;
  buttons.replaceChildren(...(review.options || []).map((option, index) => {
    const button = node("button", "", "review-option");
    button.type = "button";
    button.disabled = reviewSubmitting;
    button.append(node("strong", "ABCDEF"[index] || String(index + 1)), node("span", option), node("span", "", "chevron"));
    button.addEventListener("click", () => submitReview(option));
    return button;
  }));
  el("answer").value = "";
  el("resume").disabled = true;
  updateReviewControls();
}

function updateReviewControls() {
  el("review-options").disabled = reviewSubmitting;
  el("answer").readOnly = reviewSubmitting;
  el("resume").disabled = reviewSubmitting || !el("answer").value.trim();
}

function renderTrace(value) {
  const events = (value.trace || []).slice(-80);
  el("empty-log").hidden = events.length > 0;
  const list = el("events");
  list.replaceChildren(...events.map(event => {
    const item = node("li");
    item.dataset.tone = event.error ? "error" : event.kind === "review" ? "review" : "";
    const copy = node("div", "", "event-copy");
    const name = event.kind === "error" ? "运行错误" : (toolNames[event.name] || event.name || "处理");
    copy.append(node("strong", name));
    const detail = event.error || event.detail || "";
    if (detail) copy.append(node("span", detail, event.error ? "event-error" : ""));
    item.append(copy);
    return item;
  }));
}

async function poll(id) {
  clearTimeout(pollTimer);
  try {
    const value = await api(`/api/tasks/${encodeURIComponent(id)}`);
    renderTask(value);
    if (["queued", "running"].includes(value.status)) {
      pollTimer = setTimeout(() => poll(id), 1200);
    } else {
      submitting = false;
      updateStart();
    }
  } catch (error) {
    message("submit-error", error.message);
  }
}

async function startTask(event) {
  event.preventDefault();
  if (el("start").disabled) return;
  clearTimeout(pollTimer);
  submitting = true;
  updateStart();
  message("submit-error", "");
  message("task-error", "");
  const selected = files.filter(entry => entry.selected);
  const body = new FormData();
  body.append("title", form.elements.title.value.trim());
  body.append("group", form.elements.group.value.trim());
  body.append("preset", activePreset);
  for (const entry of selected) body.append("files", entry.file, entry.file.name);
  el("start").textContent = "正在上传…";
  try {
    const value = await api("/api/tasks", {method: "POST", body});
    task = value;
    remember(value.id);
    renderTask(value);
    if (["queued", "running"].includes(value.status)) poll(value.id);
  } catch (error) {
    message("submit-error", error.message);
  } finally {
    submitting = false;
    el("start").textContent = "开始整理 →";
    updateStart();
  }
}

async function submitReview(answer) {
  if (!task || !task.review || reviewSubmitting) return;
  const value = (answer || el("answer").value).trim();
  if (!value) return;
  reviewSubmitting = true;
  message("review-error", "");
  updateReviewControls();
  try {
    const updated = await api(`/api/tasks/${encodeURIComponent(task.id)}/review`, {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({answer: value}),
    });
    renderTask(updated);
    poll(task.id);
  } catch (error) {
    message("review-error", error.message);
  } finally {
    reviewSubmitting = false;
    updateReviewControls();
  }
}

function openSources(value) {
  const dialog = el("source-dialog");
  const root = el("sources");
  root.replaceChildren();
  const sources = value.review?.sources || [];
  if (!sources.length) {
    root.append(node("p", "这条核对没有关联的原文位置。", "quiet"));
    dialog.showModal();
    return;
  }
  for (const source of sources) {
    const parts = sourceParts(value, source);
    const details = node("details");
    details.open = true;
    details.append(node("summary", source));
    if (parts.file && /\.pdf$/i.test(parts.name)) {
      const frame = node("iframe", "", "source-frame");
      frame.title = source;
      frame.src = `/api/tasks/${encodeURIComponent(value.id)}/source/${encodeURIComponent(parts.file.id)}` +
        (parts.page ? `#page=${parts.page}` : "");
      details.append(frame);
    } else if (parts.file && /\.(png|jpe?g|webp)$/i.test(parts.name)) {
      const image = node("img", "", "source-image");
      image.src = `/api/tasks/${encodeURIComponent(value.id)}/source/${encodeURIComponent(parts.file.id)}`;
      image.alt = source;
      details.append(image);
    } else if (parts.file) {
      const link = node("a", "下载原资料", "secondary");
      link.href = `/api/tasks/${encodeURIComponent(value.id)}/source/${encodeURIComponent(parts.file.id)}`;
      details.append(link);
    }
    const excerpt = value.trace?.find(event => event.detail === source)?.detail;
    if (excerpt) details.append(node("p", excerpt, "source-excerpt"));
    root.append(details);
  }
  dialog.showModal();
}

async function showHistory() {
  const dialog = el("history-dialog");
  dialog.showModal();
  const list = el("history-list");
  list.replaceChildren();
  message("history-status", "正在读取生成历史…");
  try {
    const entries = await api("/api/history");
    message("history-status", entries.length ? `共 ${entries.length} 条已生成记录` : "还没有生成记录。");
    list.replaceChildren(...entries.map(entry => {
      const item = node("li", "", "history-item");
      const copy = node("div", "", "history-copy");
      copy.append(node("strong", entry.title || "未命名赛事"));
      copy.append(node("small", [entry.group, entry.preset_name, entry.records ? `${entry.records} 条记录` :
        null, entry.generated_at ? historyDate.format(new Date(entry.generated_at)) : null].filter(Boolean).join(" · ")));
      item.append(copy);
      if (entry.word_available) {
        const link = node("a", "", "secondary history-download");
        link.href = `/api/tasks/${encodeURIComponent(entry.id)}/download/word`;
        link.append(node("span", "", "download-icon"), "下载");
        item.append(link);
      }
      return item;
    }));
  } catch (error) {
    message("history-status", error.message);
  }
}

async function restoreTask(id) {
  try {
    const value = await api(`/api/tasks/${encodeURIComponent(id)}`);
    task = value;
    form.elements.title.value = value.title || "";
    form.elements.group.value = value.group || "";
    if (value.preset) selectPreset(value.preset);
    renderTask(value);
    if (["queued", "running"].includes(value.status)) poll(value.id);
    else submitting = false;
  } catch {
    remember(null);
  }
}

async function init() {
  try {
    options = await api("/api/options");
  } catch {
    message("submit-error", "本地服务连接失败，请刷新重试。");
    return;
  }
  const upload = options.upload;
  el("file-input").accept = upload.extensions.join(",");
  el("upload-limit").textContent = `Excel、Word、PDF 或图片 · 单份不超过 ${Math.round(upload.max_file_bytes / 1024 / 1024)} MB`;
  if (!options.key_configured) message("submit-error", "服务端未配置 DASHSCOPE_API_KEY，暂时不能开始。");
  renderPresets(options.presets);
  form.addEventListener("submit", startTask);
  el("history-open").addEventListener("click", showHistory);
  el("activity-toggle").addEventListener("click", () => {
    const open = el("activity-toggle").getAttribute("aria-expanded") !== "true";
    el("activity-toggle").setAttribute("aria-expanded", String(open));
    el("activity-log").hidden = !open;
  });
  el("answer").addEventListener("input", updateReviewControls);
  el("restart-task").addEventListener("click", async () => {
    if (!task || task.status !== "error") return;
    try {
      const value = await api(`/api/tasks/${encodeURIComponent(task.id)}/restart`, {method: "POST"});
      task = value;
      remember(value.id);
      renderTask(value);
      poll(value.id);
    } catch (error) {
      message("task-error", error.message);
    }
  });
  el("review-reply").addEventListener("submit", event => {
    event.preventDefault();
    submitReview();
  });
  el("add-files").addEventListener("click", () => el("file-input").click());
  el("empty-upload").addEventListener("click", () => el("file-input").click());
  el("file-input").addEventListener("change", event => {
    addFiles([...event.target.files]);
    event.target.value = "";
  });
  form.elements.title.addEventListener("input", updateStart);
  const drop = el("drop");
  for (const type of ["dragenter", "dragover"]) {
    drop.addEventListener(type, event => {
      event.preventDefault();
      drop.classList.add("over");
    });
  }
  for (const type of ["dragleave", "drop"]) {
    drop.addEventListener(type, event => {
      event.preventDefault();
      drop.classList.remove("over");
    });
  }
  drop.addEventListener("drop", event => addFiles([...event.dataTransfer.files]));
  const id = new URLSearchParams(location.search).get("task") || localStorage.getItem(storageKey);
  if (id) restoreTask(id);
  updateStart();
}

init();
