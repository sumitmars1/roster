const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

let templatesCache = [];
let projectsCache = [];
let activeProjectId = null;
let replaceTargetMatId = null;
let projPollTimer = null;
let drawPollTimer = null;

function esc(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fmtSize(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

async function apiJson(url, opts = {}) {
  const res = await fetch(url, opts);
  const text = await res.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { detail: text };
  }
  if (!res.ok) {
    throw new Error(data.detail || `请求失败 (${res.status})`);
  }
  return data;
}

// ---------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------

function initTabs() {
  const tabs = $$(".nav-tab");
  const panels = {
    projects: $("#tab-projects"),
    templates: $("#tab-templates"),
    draw: $("#tab-draw"),
    roster: $("#tab-roster"),
  };
  tabs.forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.tab;
      tabs.forEach((t) => t.classList.toggle("is-active", t === btn));
      Object.entries(panels).forEach(([k, el]) => {
        if (el) el.hidden = k !== target;
      });
    });
  });
}

// ---------------------------------------------------------------------------
// Templates
// ---------------------------------------------------------------------------

async function loadTemplates() {
  const data = await apiJson("/api/templates");
  templatesCache = data.templates || [];
  renderTemplateSelect();
  renderTemplateCards();
}

function renderTemplateSelect() {
  const sel = $("#proj-template-select");
  if (!sel) return;
  const prev = sel.value;
  sel.innerHTML = templatesCache
    .map(
      (t) =>
        `<option value="${esc(t.id)}">${esc(t.name)}${t.builtin ? "（内置）" : "（已入库）"}</option>`
    )
    .join("");
  if (prev && templatesCache.some((t) => t.id === prev)) {
    sel.value = prev;
  }
  updateTemplateDesc();
}

function updateTemplateDesc() {
  const sel = $("#proj-template-select");
  const desc = $("#proj-template-desc");
  if (!sel || !desc) return;
  const tpl = templatesCache.find((t) => t.id === sel.value);
  desc.textContent = tpl ? tpl.description || "" : "";
}

function renderTemplateCards() {
  const grid = $("#tpl-cards");
  if (!grid) return;
  grid.innerHTML = templatesCache
    .map((t) => {
      const fmt = t.formatting || {};
      const bodyFont = fmt.fonts?.body || {};
      const chaps = (t.chapters || [])
        .filter((c) => c.toc_visible !== false)
        .map((c) => c.title)
        .join(" / ");
      const hasWm = fmt.watermark?.enabled && fmt.watermark?.asset;
      return `
        <article class="panel tpl-card">
          <div class="row-between">
            <h3 class="minihead">${esc(t.name)}</h3>
            <span class="pill ${t.builtin ? "" : "is-ready"}">${t.builtin ? "内置模板" : "自定义入库"}</span>
          </div>
          <p class="quiet tiny">${esc(t.description || "")}</p>
          <div class="tpl-badges">
            <span class="pill">名单母版: ${esc(t.roster_preset || "inline-roster")}</span>
            <span class="pill">正文: ${esc(bodyFont.name || "仿宋")} ${esc(bodyFont.size_pt || 16)}pt</span>
            ${hasWm ? `<span class="pill is-running">含底图水印</span>` : ""}
          </div>
          <div class="tpl-chapters"><strong>章节顺序：</strong>${esc(chaps || "自动")}</div>
          <div class="row-between">
            ${
              hasWm
                ? `<a class="quiet tiny" href="/api/templates/${encodeURIComponent(t.id)}/watermark" target="_blank">查看模板底图</a>`
                : `<span></span>`
            }
            ${
              t.builtin
                ? ""
                : `<button type="button" class="secondary tiny-btn tpl-del-btn" data-id="${esc(t.id)}">删除</button>`
            }
          </div>
        </article>
      `;
    })
    .join("");

  $$(".tpl-del-btn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      try {
        await apiJson(`/api/templates/${encodeURIComponent(btn.dataset.id)}`, { method: "DELETE" });
        await loadTemplates();
      } catch (e) {
        alert(e.message);
      }
    });
  });
}

function initTemplateForms() {
  const sel = $("#proj-template-select");
  if (sel) sel.addEventListener("change", updateTemplateDesc);

  const parseForm = $("#tpl-parse-form");
  if (parseForm) {
    parseForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const fileInput = $("#tpl-sample-input");
      const nameInput = $("#tpl-name-input");
      const msg = $("#tpl-parse-msg");
      const err = $("#tpl-parse-err");
      const btn = $("#tpl-parse-btn");
      err.hidden = true;
      if (!fileInput.files.length) return;

      const fd = new FormData();
      fd.append("file", fileInput.files[0]);
      fd.append("name", nameInput.value.trim());
      fd.append("use_llm", "true");

      btn.disabled = true;
      msg.textContent = "正在解析样例结构、版式、名单风格及固定底图…";
      try {
        const created = await apiJson("/api/templates/parse", { method: "POST", body: fd });
        msg.textContent = `解析入库成功：${created.name}`;
        parseForm.reset();
        await loadTemplates();
      } catch (ex) {
        err.textContent = ex.message;
        err.hidden = false;
        msg.textContent = "";
      } finally {
        btn.disabled = false;
      }
    });
  }

  const covForm = $("#cover-gen-form");
  if (covForm) {
    covForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const fd = new FormData();
      fd.append("event_title", $("#cover-title-input").value.trim());
      fd.append("footer", $("#cover-footer-input").value);
      const bg = $("#cover-bg-input");
      if (bg.files.length) fd.append("file", bg.files[0]);
      try {
        const data = await apiJson("/api/cover/preview", { method: "POST", body: fd });
        const box = $("#cover-preview-box");
        const img = $("#cover-preview-img");
        const dl = $("#cover-preview-dl");
        const url = `${data.preview_url}?t=${Date.now()}`;
        img.src = url;
        dl.href = data.preview_url;
        box.hidden = false;
      } catch (ex) {
        alert(ex.message);
      }
    });
  }
}

// ---------------------------------------------------------------------------
// Projects
// ---------------------------------------------------------------------------

async function loadProjects() {
  const data = await apiJson("/api/projects");
  projectsCache = data.projects || [];
  renderProjectList();
  if (activeProjectId) {
    await selectProject(activeProjectId, false);
  } else if (projectsCache.length) {
    await selectProject(projectsCache[0].id, false);
  }
}

function renderProjectList() {
  const ul = $("#proj-list");
  if (!ul) return;
  if (!projectsCache.length) {
    ul.innerHTML = `<li class="quiet tiny">暂无项目，请在上方新建。</li>`;
    return;
  }
  ul.innerHTML = projectsCache
    .map(
      (p) => `
      <li class="proj-item ${p.id === activeProjectId ? "is-selected" : ""}" data-id="${esc(p.id)}">
        <div class="proj-item-name">${esc(p.name)}</div>
        <div class="proj-item-sub">
          <span>${(p.materials || []).length} 份资料 · ${(p.versions || []).length} 版</span>
          <span>${esc(p.stage || p.status || "")}</span>
        </div>
      </li>`
    )
    .join("");
  $$(".proj-item").forEach((el) => {
    el.addEventListener("click", () => selectProject(el.dataset.id, true));
  });
}

async function selectProject(projectId, scrollToTop = false) {
  activeProjectId = projectId;
  renderProjectList();
  const detail = await apiJson(`/api/projects/${encodeURIComponent(projectId)}`);
  renderActiveProject(detail);
  if (detail.status === "running") {
    scheduleProjectPoll(projectId);
  } else {
    clearProjectPoll();
  }
  if (scrollToTop) {
    $("#proj-workspace")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function clearProjectPoll() {
  if (projPollTimer) {
    clearTimeout(projPollTimer);
    projPollTimer = null;
  }
}

function scheduleProjectPoll(projectId) {
  clearProjectPoll();
  projPollTimer = setTimeout(async () => {
    if (activeProjectId !== projectId) return;
    try {
      const detail = await apiJson(`/api/projects/${encodeURIComponent(projectId)}`);
      renderActiveProject(detail);
      if (detail.status === "running") {
        scheduleProjectPoll(projectId);
      } else {
        await loadProjects();
      }
    } catch {
      scheduleProjectPoll(projectId);
    }
  }, 2500);
}

function renderActiveProject(proj) {
  $("#proj-empty-hint").hidden = true;
  $("#proj-active").hidden = false;

  $("#proj-active-title").textContent = proj.name;
  $("#proj-active-meta").textContent = `模板：${proj.template?.name || proj.template_id} · 项目ID：${proj.id}`;

  const pill = $("#proj-status-pill");
  const st = proj.status || "draft";
  pill.className = `pill ${st === "running" ? "is-running" : st === "ready" ? "is-ready" : st === "error" ? "is-error" : ""}`;
  pill.textContent = proj.stage || st;

  const genBtn = $("#proj-generate-btn");
  genBtn.disabled = st === "running";
  genBtn.textContent =
    st === "running"
      ? "正在生成中…"
      : proj.latest_version_id
      ? "重新生成整册 Word + PDF"
      : "生成整册 Word + PDF";

  const stageLine = $("#proj-stage-line");
  if (st === "running") {
    stageLine.textContent = `当前阶段：${proj.stage || "处理中"}`;
    stageLine.hidden = false;
  } else {
    stageLine.hidden = true;
  }

  const errEl = $("#proj-error-line");
  if (proj.error) {
    errEl.textContent = proj.error;
    errEl.hidden = false;
  } else {
    errEl.hidden = true;
  }

  // Materials
  const mats = proj.materials || [];
  const matUl = $("#proj-materials-list");
  if (!mats.length) {
    matUl.innerHTML = `<li class="quiet tiny">尚未上传任何资料，请点击右上角“+ 上传/补充资料”。</li>`;
  } else {
    matUl.innerHTML = mats
      .map(
        (m) => `
        <li class="mat-item">
          <div class="mat-item-meta">
            <strong>${esc(m.name)}</strong>
            <span class="quiet">${fmtSize(m.size)}</span>
            ${m.role_hint && m.role_hint !== "auto" ? `<span class="pill">${esc(m.role_hint)}</span>` : ""}
          </div>
          <div class="row-gap">
            <a class="quiet tiny" href="/api/projects/${encodeURIComponent(proj.id)}/materials/${encodeURIComponent(m.id)}/raw" target="_blank">下载原件</a>
            <button type="button" class="secondary tiny-btn mat-replace-btn" data-id="${esc(m.id)}">替换</button>
            <button type="button" class="secondary tiny-btn mat-del-btn" data-id="${esc(m.id)}">删除</button>
          </div>
        </li>`
      )
      .join("");
  }

  $$(".mat-replace-btn").forEach((b) => {
    b.addEventListener("click", () => {
      replaceTargetMatId = b.dataset.id;
      $("#mat-replace-input").click();
    });
  });

  $$(".mat-del-btn").forEach((b) => {
    b.addEventListener("click", async () => {
      try {
        const res = await apiJson(
          `/api/projects/${encodeURIComponent(proj.id)}/materials/${encodeURIComponent(b.dataset.id)}`,
          { method: "DELETE" }
        );
        renderActiveProject({ ...proj, materials: res.project.materials });
      } catch (ex) {
        alert(ex.message);
      }
    });
  });

  // Trace
  const traces = proj.trace || [];
  const traceBox = $("#proj-trace-box");
  const traceList = $("#proj-trace-list");
  if (traces.length) {
    traceBox.hidden = false;
    traceList.innerHTML = traces
      .map((t) => `<li><code>${esc(t.ts?.slice(11, 19) || "")}</code> ${esc(t.stage || "")} ${esc(JSON.stringify(t))}</li>`)
      .join("");
  } else {
    traceBox.hidden = true;
  }

  // Versions & PDF Viewer
  const versions = proj.versions || [];
  const resCard = $("#proj-result-card");
  if (!versions.length) {
    resCard.hidden = true;
    return;
  }
  resCard.hidden = false;
  const verSel = $("#proj-version-select");
  const prevVid = verSel.value;
  verSel.innerHTML = versions
    .slice()
    .reverse()
    .map(
      (v) =>
        `<option value="${esc(v.id)}">${esc(v.id)} (${v.kind === "draw_update" ? "抽签更新" : "整册生成"} · ${v.pdf_pages || "?"} 页)</option>`
    )
    .join("");
  const chosenId = versions.some((v) => v.id === prevVid) ? prevVid : proj.latest_version_id || versions[versions.length - 1].id;
  verSel.value = chosenId;
  renderSelectedVersion(proj, chosenId);
}

function renderSelectedVersion(proj, versionId) {
  const ver = (proj.versions || []).find((v) => v.id === versionId);
  if (!ver) return;

  $("#proj-version-meta").textContent = `版本 ${ver.id} · ${ver.template_name || ""} · 共 ${ver.pdf_pages || "?"} 页 · 生成于 ${ver.created_at || ""}`;
  $("#proj-dl-docx").href = `/api/projects/${encodeURIComponent(proj.id)}/versions/${encodeURIComponent(ver.id)}/download/docx`;
  $("#proj-dl-pdf").href = `/api/projects/${encodeURIComponent(proj.id)}/versions/${encodeURIComponent(ver.id)}/download/pdf`;

  const previewUrl = `/api/projects/${encodeURIComponent(proj.id)}/versions/${encodeURIComponent(ver.id)}/preview/pdf#view=FitH`;
  const frame = $("#proj-pdf-frame");
  if (frame.dataset.loadedUrl !== previewUrl) {
    frame.src = previewUrl;
    frame.dataset.loadedUrl = previewUrl;
  }

  const tocHtml = (ver.toc || [])
    .map((item) => `<li>${esc(item.title)} — 第 ${esc(item.page ?? "?")} 页</li>`)
    .join("");
  const missingHtml = (ver.missing_items || []).length
    ? (ver.missing_items || []).map((x) => `<li>${esc(x)}</li>`).join("")
    : `<li class="quiet">未发现缺失或遗漏项</li>`;
  const exclHtml = [
    ...(ver.excluded_items || []).map((x) => `<li>${esc(x)}</li>`),
    ...(ver.unrecognized_files || []).map((u) => `<li>未识别：${esc(u.file)} — ${esc(u.reason)}</li>`),
  ].join("");
  const drawBlock = ver.draw_summary
    ? `<div class="audit-box"><h4>抽签赛程替换核对</h4><p>已替换 <strong>${ver.draw_summary.replaced_count}</strong> 处签位，保留 <strong>${ver.draw_summary.preserved_count}</strong> 处胜负/排名引用。</p><p class="quiet">${esc(ver.draw_summary.agent_summary || "")}</p></div>`
    : "";

  $("#proj-audit-box").innerHTML = `
    <div class="audit-box"><h4>已生成目录页码</h4><ul>${tocHtml || "<li>无目录</li>"}</ul></div>
    <div class="audit-box"><h4>缺失项与完整性说明</h4><ul>${missingHtml}</ul></div>
    ${exclHtml ? `<div class="audit-box"><h4>范围排除 / 未识别资料说明</h4><ul>${exclHtml}</ul></div>` : ""}
    ${drawBlock}
  `;
}

function initProjectEvents() {
  $("#proj-create-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const fd = new FormData();
    fd.append("name", $("#proj-name-input").value.trim());
    fd.append("template_id", $("#proj-template-select").value);
    try {
      const created = await apiJson("/api/projects", { method: "POST", body: fd });
      $("#proj-name-input").value = "";
      await loadProjects();
      await selectProject(created.id, false);
    } catch (ex) {
      alert(ex.message);
    }
  });

  $("#proj-refresh-btn").addEventListener("click", loadProjects);

  $("#mat-add-btn").addEventListener("click", () => $("#mat-file-input").click());

  $("#mat-file-input").addEventListener("change", async (e) => {
    const files = e.target.files;
    if (!files.length || !activeProjectId) return;
    const fd = new FormData();
    for (const f of files) fd.append("files", f);
    fd.append("role_hint", $("#mat-role-select").value);
    try {
      await apiJson(`/api/projects/${encodeURIComponent(activeProjectId)}/materials`, {
        method: "POST",
        body: fd,
      });
      e.target.value = "";
      await selectProject(activeProjectId, false);
    } catch (ex) {
      alert(ex.message);
    }
  });

  $("#mat-replace-input").addEventListener("change", async (e) => {
    const files = e.target.files;
    if (!files.length || !activeProjectId || !replaceTargetMatId) return;
    const fd = new FormData();
    fd.append("files", files[0]);
    fd.append("replace_material_id", replaceTargetMatId);
    try {
      await apiJson(`/api/projects/${encodeURIComponent(activeProjectId)}/materials`, {
        method: "POST",
        body: fd,
      });
      e.target.value = "";
      replaceTargetMatId = null;
      await selectProject(activeProjectId, false);
    } catch (ex) {
      alert(ex.message);
    }
  });

  $("#proj-generate-btn").addEventListener("click", async () => {
    if (!activeProjectId) return;
    const fd = new FormData();
    fd.append("note", $("#proj-note-input").value.trim());
    try {
      const updated = await apiJson(`/api/projects/${encodeURIComponent(activeProjectId)}/generate`, {
        method: "POST",
        body: fd,
      });
      renderActiveProject(updated);
      scheduleProjectPoll(activeProjectId);
    } catch (ex) {
      alert(ex.message);
    }
  });

  $("#proj-version-select").addEventListener("change", async (e) => {
    if (!activeProjectId) return;
    const proj = await apiJson(`/api/projects/${encodeURIComponent(activeProjectId)}`);
    renderSelectedVersion(proj, e.target.value);
  });

  $("#proj-draw-btn").addEventListener("click", async () => {
    if (!activeProjectId) return;
    const files = $("#proj-draw-files").files;
    if (!files.length) {
      alert("请先选择抽签结果文件。");
      return;
    }
    const fd = new FormData();
    for (const f of files) fd.append("files", f);
    fd.append("source_version_id", $("#proj-version-select").value);
    fd.append("note", $("#proj-draw-note").value.trim());
    try {
      const updated = await apiJson(`/api/projects/${encodeURIComponent(activeProjectId)}/draw-update`, {
        method: "POST",
        body: fd,
      });
      $("#proj-draw-files").value = "";
      renderActiveProject(updated);
      scheduleProjectPoll(activeProjectId);
    } catch (ex) {
      alert(ex.message);
    }
  });
}

// ---------------------------------------------------------------------------
// Standalone Post-Draw Update
// ---------------------------------------------------------------------------

function initStandaloneDraw() {
  const form = $("#standalone-draw-form");
  if (!form) return;
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const bookFile = $("#sd-book-input").files[0];
    const drawFiles = $("#sd-draw-input").files;
    const errEl = $("#sd-error");
    const statusText = $("#sd-status-text");
    const btn = $("#sd-submit-btn");
    errEl.hidden = true;
    if (!bookFile || !drawFiles.length) return;

    const fd = new FormData();
    fd.append("book_file", bookFile);
    for (const f of drawFiles) fd.append("draw_files", f);
    fd.append("note", $("#sd-note-input").value.trim());

    btn.disabled = true;
    statusText.textContent = "已提交，正在读取抽签表并核对赛程上下文…";
    try {
      const job = await apiJson("/api/draw-update", { method: "POST", body: fd });
      pollStandaloneDrawJob(job.id);
    } catch (ex) {
      errEl.textContent = ex.message;
      errEl.hidden = false;
      statusText.textContent = "";
      btn.disabled = false;
    }
  });
}

function pollStandaloneDrawJob(jobId) {
  if (drawPollTimer) clearTimeout(drawPollTimer);
  drawPollTimer = setTimeout(async () => {
    try {
      const job = await apiJson(`/api/draw-jobs/${encodeURIComponent(jobId)}`);
      if (job.status === "running") {
        pollStandaloneDrawJob(jobId);
        return;
      }
      $("#sd-submit-btn").disabled = false;
      if (job.status === "error") {
        $("#sd-status-text").textContent = "";
        $("#sd-error").textContent = job.error || "抽签更新失败";
        $("#sd-error").hidden = false;
        return;
      }
      $("#sd-status-text").textContent = "更新完成";
      const card = $("#sd-result-card");
      card.hidden = false;
      $("#sd-summary-line").textContent = `已替换 ${job.replaced_count} 处抽签签位，完好保留 ${job.preserved_count} 处胜负/排名引用。${job.agent_summary || ""}`;
      $("#sd-dl-docx").href = `/api/draw-jobs/${encodeURIComponent(jobId)}/download/docx`;
      $("#sd-dl-pdf").href = `/api/draw-jobs/${encodeURIComponent(jobId)}/download/pdf`;
      $("#sd-pdf-frame").src = `/api/draw-jobs/${encodeURIComponent(jobId)}/preview/pdf#view=FitH`;

      const chgHtml = (job.changes || [])
        .slice(0, 40)
        .map((c) => `<li><code>${esc(c.location)}</code>: ${esc(c.before)} → <strong>${esc(c.after)}</strong></li>`)
        .join("");
      const presHtml = (job.preserved || [])
        .slice(0, 30)
        .map((p) => `<li><code>${esc(p.location || "")}</code>: ${esc(p.text || "")} ${p.reason ? `(${esc(p.reason)})` : ""}</li>`)
        .join("");
      $("#sd-diff-box").innerHTML = `
        <div class="audit-box"><h4>已替换赛程签位（前40条）</h4><ul>${chgHtml || "<li>无改动</li>"}</ul></div>
        <div class="audit-box"><h4>保持不动的胜负/排名引用</h4><ul>${presHtml || "<li>无特殊引用</li>"}</ul></div>
      `;
    } catch {
      pollStandaloneDrawJob(jobId);
    }
  }, 2000);
}

// Boot
initTabs();
initTemplateForms();
initProjectEvents();
initStandaloneDraw();
loadTemplates().then(loadProjects).catch((e) => console.error(e));
