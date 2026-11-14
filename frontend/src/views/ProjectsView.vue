<template>
  <div class="page-container">
    <div class="page-header">
      <div>
        <h1 class="page-title">整册秩序册生成</h1>
        <p class="page-subtitle">
          选定公文或自定义版式模板，上传全套赛事资料（封面、规程、组织机构、报名表、分组、赛程、对阵等），由 Agent 研判整批资料并直接生成可预览、可下载的 Word 与 PDF 秩序册。
        </p>
      </div>
      <el-button type="primary" size="large" @click="ui.openCreateProjectDialog()">
        <FolderPlus :size="16" />
        <span style="margin-left: 6px">新建秩序册项目</span>
      </el-button>
    </div>

    <div class="workspace-grid">
      <!-- Left Sidebar: Project List -->
      <aside class="sidebar-col">
        <div class="white-card sidebar-card">
          <div class="sidebar-head">
            <span class="section-title">
              <Layers :size="16" class="accent-icon" />
              <span>项目列表</span>
              <el-tag size="small" round effect="plain" type="info">{{ projectsStore.projects.length }}</el-tag>
            </span>
            <el-button
              circle
              size="small"
              :loading="projectsStore.loadingList"
              title="刷新项目"
              @click="projectsStore.fetchProjects({ autoSelect: false })"
            >
              <RefreshCw :size="13" />
            </el-button>
          </div>

          <el-input
            v-model="searchKeyword"
            placeholder="搜索项目名称…"
            clearable
            size="default"
            class="sidebar-search"
          >
            <template #prefix>
              <Search :size="14" class="muted-icon" />
            </template>
          </el-input>

          <el-skeleton v-if="projectsStore.loadingList && !projectsStore.projects.length" :rows="4" animated />

          <el-empty
            v-else-if="!filteredProjects.length"
            :description="searchKeyword ? '未找到匹配项目' : '暂无项目，点击右上角新建'"
            :image-size="60"
          />

          <ul v-else class="project-list">
            <li
              v-for="proj in filteredProjects"
              :key="proj.id"
              class="project-item"
              :class="{ 'is-active': proj.id === projectsStore.activeProjectId }"
              @click="handleSelectProject(proj.id)"
            >
              <div class="project-item-top">
                <span class="status-dot" :class="`is-${projectStatusMeta(proj.status).dot}`" />
                <span class="project-item-name">{{ proj.name }}</span>
              </div>
              <div class="project-item-meta tabular">
                <span>{{ (proj.materials || []).length }} 份资料</span>
                <span>·</span>
                <span>{{ (proj.versions || []).length }} 个版本</span>
                <span class="project-item-stage">{{ proj.stage || projectStatusMeta(proj.status).label }}</span>
              </div>
            </li>
          </ul>
        </div>
      </aside>

      <!-- Main Area -->
      <section class="main-col">
        <div v-if="!projectsStore.activeProject && !projectsStore.loadingDetail" class="white-card empty-workspace">
          <el-empty description="在左侧选择一个秩序册项目，或创建新项目开始上传赛事资料" :image-size="92">
            <el-button type="primary" @click="ui.openCreateProjectDialog()">
              <FolderPlus :size="15" />
              <span style="margin-left: 6px">立即创建项目</span>
            </el-button>
          </el-empty>
        </div>

        <div v-else-if="projectsStore.loadingDetail && !projectsStore.activeProject" class="white-card card-pad">
          <el-skeleton :rows="8" animated />
        </div>

        <template v-else-if="active">
          <!-- Active Project Header Card -->
          <div class="white-card card-pad">
            <div class="proj-header-row">
              <div class="proj-header-info">
                <div class="proj-title-line">
                  <h2 class="proj-title">{{ active.name }}</h2>
                  <el-button
                    v-if="active.status !== 'running'"
                    size="small"
                    text
                    type="primary"
                    @click="openEditDialog"
                  >
                    <Pencil :size="13" />
                    <span style="margin-left: 4px">修改名称 / 模板</span>
                  </el-button>
                </div>
                <div class="proj-sub-meta">
                  <el-tag size="small" effect="plain" type="info">
                    模板：{{ active.template?.name || active.template_id }}
                  </el-tag>
                  <span class="muted-text tabular">ID: {{ active.id }}</span>
                </div>
              </div>

              <el-tag
                :type="projectStatusMeta(active.status).type"
                effect="light"
                size="large"
                round
                class="status-pill"
              >
                <span class="status-pill-inner">
                  <Loader2 v-if="active.status === 'running'" :size="14" class="spin-icon" />
                  <CheckCircle2 v-else-if="active.status === 'ready'" :size="14" />
                  <AlertCircle v-else-if="active.status === 'error'" :size="14" />
                  <Clock v-else :size="14" />
                  <span>{{ active.stage || projectStatusMeta(active.status).label }}</span>
                </span>
              </el-tag>
            </div>

            <div v-if="active.status === 'running'" class="running-banner">
              <Loader2 :size="18" class="spin-icon accent-icon" />
              <div class="running-copy">
                <strong>正在智能生成中…</strong>
                <span>{{ active.stage || '正在研判资料并排版秩序册，请稍候' }}</span>
              </div>
            </div>

            <el-alert
              v-if="active.error"
              type="error"
              :title="active.error"
              show-icon
              :closable="false"
              style="margin-top: 14px"
            />
          </div>

          <!-- Step 1: Materials Management -->
          <div class="white-card card-pad">
            <div class="card-head-row">
              <div>
                <h3 class="section-title">
                  <FolderOpen :size="16" class="accent-icon" />
                  <span>1. 项目原始资料</span>
                  <el-tag size="small" round type="info" effect="plain">
                    {{ (active.materials || []).length }} 份
                  </el-tag>
                </h3>
                <p class="muted-text" style="margin: 4px 0 0">
                  支持 Word (.docx/.doc/.wps)、Excel (.xlsx/.xls/.et)、PDF 及图片（封面、底图、尾页、对阵图等），可随时增删或替换。
                </p>
              </div>

              <div class="role-upload-controls">
                <el-select v-model="roleHint" style="width: 170px" size="default">
                  <el-option
                    v-for="opt in ROLE_HINT_OPTIONS"
                    :key="opt.value"
                    :label="opt.label"
                    :value="opt.value"
                  />
                </el-select>
                <el-button
                  type="primary"
                  plain
                  :loading="projectsStore.uploadingMaterials"
                  @click="triggerMaterialPicker"
                >
                  <Upload :size="15" />
                  <span style="margin-left: 6px">选择文件上传</span>
                </el-button>
              </div>
            </div>

            <div
              class="dropzone"
              :class="{ 'is-over': draggingMaterials }"
              @dragenter.prevent="draggingMaterials = true"
              @dragover.prevent="draggingMaterials = true"
              @dragleave.prevent="draggingMaterials = false"
              @drop.prevent="handleDropMaterials"
              @click="triggerMaterialPicker"
            >
              <UploadCloud :size="24" class="accent-icon" />
              <div class="dropzone-text">
                <strong>点击或将全套赛事资料拖拽到此处上传</strong>
                <span>当前上传用途预设：{{ roleHintLabel(roleHint) }}（支持批量多选）</span>
              </div>
            </div>

            <input
              ref="materialInputRef"
              type="file"
              multiple
              hidden
              @change="handleMaterialFileChange"
            />
            <input
              ref="replaceInputRef"
              type="file"
              hidden
              @change="handleReplaceFileChange"
            />

            <div v-if="(active.materials || []).length" class="materials-list">
              <div v-for="mat in active.materials" :key="mat.id" class="material-row">
                <div class="material-left">
                  <FileKindIcon :filename="mat.name" />
                  <div class="material-info">
                    <div class="material-name-line">
                      <span class="material-name">{{ mat.name }}</span>
                      <el-tag
                        v-if="mat.role_hint && mat.role_hint !== 'auto'"
                        size="small"
                        type="primary"
                        effect="light"
                      >
                        {{ roleHintLabel(mat.role_hint) }}
                      </el-tag>
                    </div>
                    <span class="muted-text tabular">{{ formatSize(mat.size) }}</span>
                  </div>
                </div>

                <div class="material-actions">
                  <a
                    :href="api.materialRawUrl(active.id, mat.id)"
                    target="_blank"
                    class="action-anchor"
                  >
                    <el-button size="small" text>
                      <Download :size="14" />
                      <span style="margin-left: 4px">原件</span>
                    </el-button>
                  </a>
                  <el-button size="small" text type="primary" @click="triggerReplace(mat.id)">
                    <Replace :size="14" />
                    <span style="margin-left: 4px">替换</span>
                  </el-button>
                  <el-popconfirm
                    title="确定删除这份资料吗？"
                    confirm-button-text="删除"
                    cancel-button-text="取消"
                    @confirm="handleDeleteMaterial(mat.id)"
                  >
                    <template #reference>
                      <el-button size="small" text type="danger">
                        <Trash2 :size="14" />
                      </el-button>
                    </template>
                  </el-popconfirm>
                </div>
              </div>
            </div>

            <!-- Step 2: Generation Action Bar -->
            <div class="gen-console">
              <div class="gen-input-wrap">
                <el-input
                  v-model="genNote"
                  placeholder="可选补充要求，例如：按规程排除未开赛组别、赛程表横向排版…"
                  maxlength="300"
                  clearable
                >
                  <template #prefix>
                    <Sparkles :size="15" class="accent-icon" />
                  </template>
                </el-input>
                <div class="quick-chips">
                  <span class="muted-text">快捷指令：</span>
                  <button
                    v-for="chip in quickNoteChips"
                    :key="chip"
                    type="button"
                    class="chip-btn"
                    @click="appendQuickNote(chip)"
                  >
                    + {{ chip }}
                  </button>
                </div>
              </div>

              <el-button
                type="primary"
                size="large"
                class="gen-submit-btn"
                :disabled="active.status === 'running' || !(active.materials || []).length"
                :loading="active.status === 'running' || projectsStore.actionBusy"
                @click="handleGenerate"
              >
                <Wand2 v-if="active.status !== 'running'" :size="16" />
                <span style="margin-left: 6px">
                  {{
                    active.status === 'running'
                      ? '正在生成中…'
                      : active.latest_version_id
                      ? '重新生成整册 Word + PDF'
                      : '生成整册 Word + PDF'
                  }}
                </span>
              </el-button>
            </div>

            <!-- Trace Log -->
            <el-collapse v-if="(active.trace || []).length" class="trace-collapse">
              <el-collapse-item name="trace">
                <template #title>
                  <span class="trace-summary-title">
                    <Terminal :size="14" />
                    <span>查看 Agent 生成过程记录（{{ active.trace.length }} 条）</span>
                  </span>
                </template>
                <div class="trace-list">
                  <div v-for="(ev, idx) in active.trace" :key="idx" class="trace-item">
                    <code class="trace-ts tabular">{{ (ev.ts || '').slice(11, 19) }}</code>
                    <el-tag size="small" effect="plain" type="info">{{ ev.stage || 'log' }}</el-tag>
                    <span class="trace-detail">{{ formatTraceDetail(ev) }}</span>
                  </div>
                </div>
              </el-collapse-item>
            </el-collapse>
          </div>

          <!-- Step 3: Generated Result & Preview -->
          <div v-if="currentVer" class="white-card card-pad">
            <div class="card-head-row" style="margin-bottom: 16px">
              <div>
                <h3 class="section-title">
                  <BookCheck :size="16" class="success-icon" />
                  <span>2. 已生成的秩序册成册</span>
                </h3>
                <p class="muted-text tabular" style="margin: 4px 0 0">
                  版本 {{ currentVer.id }} · {{ currentVer.template_name || active.template?.name }} · 共
                  {{ currentVer.pdf_pages || '?' }} 页 · {{ formatFullTime(currentVer.created_at) }}
                </p>
              </div>

              <el-select
                :model-value="currentVer.id"
                style="width: 260px"
                @update:model-value="projectsStore.setSelectedVersion($event)"
              >
                <el-option
                  v-for="v in projectsStore.activeVersions"
                  :key="v.id"
                  :label="`${v.id} (${v.kind === 'draw_update' ? '抽签更新' : '整册生成'} · ${v.pdf_pages || '?'} 页)`"
                  :value="v.id"
                />
              </el-select>
            </div>

            <!-- Audit Grid -->
            <div class="audit-grid">
              <div class="audit-card">
                <div class="audit-card-head">
                  <ListOrdered :size="15" class="accent-icon" />
                  <span>已生成目录页码</span>
                </div>
                <ul v-if="(currentVer.toc || []).length" class="audit-list">
                  <li v-for="(item, idx) in currentVer.toc" :key="idx">
                    <span>{{ item.title }}</span>
                    <span class="toc-page tabular">第 {{ item.page ?? '?' }} 页</span>
                  </li>
                </ul>
                <p v-else class="muted-text">暂无目录信息</p>
              </div>

              <div class="audit-card">
                <div class="audit-card-head">
                  <ShieldCheck :size="15" class="success-icon" />
                  <span>缺失项与完整性核对</span>
                </div>
                <ul v-if="(currentVer.missing_items || []).length" class="audit-list">
                  <li v-for="(item, idx) in currentVer.missing_items" :key="idx" class="warn-item">
                    {{ item }}
                  </li>
                </ul>
                <div v-else class="ok-empty">
                  <CheckCircle2 :size="15" />
                  <span>未发现缺失或遗漏项</span>
                </div>
              </div>

              <div
                v-if="(currentVer.excluded_items || []).length || (currentVer.unrecognized_files || []).length"
                class="audit-card"
              >
                <div class="audit-card-head">
                  <Info :size="15" class="warning-icon" />
                  <span>范围排除 / 未识别资料</span>
                </div>
                <ul class="audit-list">
                  <li v-for="(x, idx) in currentVer.excluded_items || []" :key="`ex-${idx}`">
                    {{ x }}
                  </li>
                  <li v-for="(u, idx) in currentVer.unrecognized_files || []" :key="`un-${idx}`">
                    未识别：{{ u.file }} — {{ u.reason }}
                  </li>
                </ul>
              </div>

              <div v-if="currentVer.draw_summary" class="audit-card">
                <div class="audit-card-head">
                  <Shuffle :size="15" class="accent-icon" />
                  <span>抽签赛程替换核对</span>
                </div>
                <p class="draw-stat">
                  已替换 <strong>{{ currentVer.draw_summary.replaced_count }}</strong> 处签位，保留
                  <strong>{{ currentVer.draw_summary.preserved_count }}</strong> 处胜负/排名引用。
                </p>
                <p v-if="currentVer.draw_summary.agent_summary" class="muted-text">
                  {{ currentVer.draw_summary.agent_summary }}
                </p>
              </div>
            </div>

            <!-- In-Project Draw Update Bar -->
            <div class="draw-inline-panel">
              <div class="draw-inline-head">
                <div class="section-title" style="font-size: 14px">
                  <Shuffle :size="15" class="accent-icon" />
                  <span>抽签后更新此版本赛程</span>
                </div>
                <span class="muted-text">上传抽签结果表，自动将签位（如 A1 VS A2）精准替换为参赛队名</span>
              </div>
              <div class="draw-inline-controls">
                <el-button plain @click="triggerDrawPicker">
                  <Paperclip :size="14" />
                  <span style="margin-left: 6px">
                    {{ drawFiles.length ? `已选 ${drawFiles.length} 份抽签表` : '选择抽签结果表' }}
                  </span>
                </el-button>
                <input
                  ref="drawInputRef"
                  type="file"
                  multiple
                  hidden
                  @change="handleDrawFileChange"
                />
                <el-input
                  v-model="drawNote"
                  placeholder="可选说明（默认替换抽签签位、保留胜负/名次引用）"
                  style="flex: 1; min-width: 220px"
                  clearable
                />
                <el-button
                  type="primary"
                  plain
                  :disabled="!drawFiles.length || active.status === 'running'"
                  :loading="projectsStore.actionBusy"
                  @click="handleProjectDrawUpdate"
                >
                  提交抽签结果并更新
                </el-button>
              </div>
              <div v-if="drawFiles.length" class="draw-picked-tags">
                <el-tag
                  v-for="(f, idx) in drawFiles"
                  :key="idx"
                  closable
                  size="small"
                  effect="plain"
                  @close="drawFiles.splice(idx, 1)"
                >
                  {{ f.name }}
                </el-tag>
              </div>
            </div>

            <!-- PDF Preview -->
            <PdfPreviewCard
              :title="`秩序册预览 (${currentVer.id})`"
              :preview-url="api.versionPdfPreviewUrl(active.id, currentVer.id)"
              :docx-url="api.versionDocxUrl(active.id, currentVer.id)"
              :pdf-url="api.versionPdfUrl(active.id, currentVer.id)"
            />
          </div>
        </template>
      </section>
    </div>

    <!-- Edit Project Dialog -->
    <el-dialog v-model="editDialogOpen" title="修改项目设置" width="460px" append-to-body>
      <el-form label-position="top">
        <el-form-item label="项目 / 赛事名称" required>
          <el-input v-model="editName" maxlength="120" />
        </el-form-item>
        <el-form-item label="版式模板" required>
          <el-select v-model="editTemplateId" style="width: 100%">
            <el-option
              v-for="tpl in templatesStore.templates"
              :key="tpl.id"
              :label="`${tpl.name}${tpl.builtin ? '（内置）' : '（已入库）'}`"
              :value="tpl.id"
            />
          </el-select>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="editDialogOpen = false">取消</el-button>
        <el-button type="primary" :loading="savingEdit" @click="handleSaveEdit">保存修改</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import {
  AlertCircle,
  BookCheck,
  CheckCircle2,
  Clock,
  Download,
  FolderOpen,
  FolderPlus,
  Info,
  Layers,
  ListOrdered,
  Loader2,
  Paperclip,
  Pencil,
  RefreshCw,
  Replace,
  Search,
  ShieldCheck,
  Shuffle,
  Sparkles,
  Terminal,
  Trash2,
  Upload,
  UploadCloud,
  Wand2,
} from 'lucide-vue-next'
import { api } from '@/api/client'
import FileKindIcon from '@/components/FileKindIcon.vue'
import PdfPreviewCard from '@/components/PdfPreviewCard.vue'
import { useProjectsStore } from '@/stores/projects'
import { useTemplatesStore } from '@/stores/templates'
import { useUiStore } from '@/stores/ui'
import {
  ROLE_HINT_OPTIONS,
  formatFullTime,
  formatSize,
  projectStatusMeta,
  roleHintLabel,
} from '@/utils/format'

const route = useRoute()
const router = useRouter()
const projectsStore = useProjectsStore()
const templatesStore = useTemplatesStore()
const ui = useUiStore()

const searchKeyword = ref('')
const roleHint = ref('auto')
const draggingMaterials = ref(false)
const genNote = ref('')
const drawFiles = ref([])
const drawNote = ref('')
const replaceTargetId = ref(null)

const materialInputRef = ref(null)
const replaceInputRef = ref(null)
const drawInputRef = ref(null)

const editDialogOpen = ref(false)
const editName = ref('')
const editTemplateId = ref('')
const savingEdit = ref(false)

const quickNoteChips = [
  '按规程排除未开赛组别',
  '赛程表横向排版',
  '严格保持报名表队伍顺序',
]

const active = computed(() => projectsStore.activeProject)
const currentVer = computed(() => projectsStore.currentVersion)

const filteredProjects = computed(() => {
  const kw = searchKeyword.value.trim().toLowerCase()
  if (!kw) return projectsStore.projects
  return projectsStore.projects.filter((p) =>
    String(p.name || '').toLowerCase().includes(kw)
  )
})

onMounted(async () => {
  const routeId = route.params.id ? String(route.params.id) : null
  await projectsStore.fetchProjects({ autoSelect: true, preferredId: routeId })
})

watch(
  () => route.params.id,
  (newId) => {
    if (newId && newId !== projectsStore.activeProjectId) {
      projectsStore.selectProject(String(newId)).catch(() => {})
    }
  }
)

onBeforeUnmount(() => {
  projectsStore.stopPolling()
})

async function handleSelectProject(id) {
  router.push({ name: 'projects', params: { id } })
  await projectsStore.selectProject(id)
}

function triggerMaterialPicker() {
  materialInputRef.value?.click()
}

async function uploadSelectedMaterials(fileList) {
  if (!active.value || !fileList?.length) return
  try {
    await projectsStore.uploadMaterials(active.value.id, Array.from(fileList), roleHint.value)
    ElMessage.success(`已上传 ${fileList.length} 份资料`)
  } catch (err) {
    ElMessage.error(err.message || '上传资料失败')
  }
}

async function handleMaterialFileChange(e) {
  const files = e.target.files
  if (files?.length) {
    await uploadSelectedMaterials(files)
  }
  e.target.value = ''
}

async function handleDropMaterials(e) {
  draggingMaterials.value = false
  const files = e.dataTransfer?.files
  if (files?.length) {
    await uploadSelectedMaterials(files)
  }
}

function triggerReplace(materialId) {
  replaceTargetId.value = materialId
  replaceInputRef.value?.click()
}

async function handleReplaceFileChange(e) {
  const file = e.target.files?.[0]
  if (file && active.value && replaceTargetId.value) {
    try {
      await projectsStore.replaceMaterial(active.value.id, replaceTargetId.value, file)
      ElMessage.success(`已替换为 ${file.name}`)
    } catch (err) {
      ElMessage.error(err.message || '替换文件失败')
    }
  }
  replaceTargetId.value = null
  e.target.value = ''
}

async function handleDeleteMaterial(materialId) {
  if (!active.value) return
  try {
    await projectsStore.deleteMaterial(active.value.id, materialId)
    ElMessage.success('已删除资料')
  } catch (err) {
    ElMessage.error(err.message || '删除失败')
  }
}

function appendQuickNote(chip) {
  if (!genNote.value.includes(chip)) {
    genNote.value = genNote.value.trim() ? `${genNote.value.trim()}；${chip}` : chip
  }
}

async function handleGenerate() {
  if (!active.value) return
  try {
    await projectsStore.generateBook(active.value.id, genNote.value)
    ElMessage.info('已提交整册秩序册生成任务')
  } catch (err) {
    ElMessage.error(err.message || '启动生成失败')
  }
}

function triggerDrawPicker() {
  drawInputRef.value?.click()
}

function handleDrawFileChange(e) {
  const files = Array.from(e.target.files || [])
  if (files.length) {
    drawFiles.value = [...drawFiles.value, ...files]
  }
  e.target.value = ''
}

async function handleProjectDrawUpdate() {
  if (!active.value || !drawFiles.value.length) return
  try {
    await projectsStore.submitDrawUpdate(active.value.id, {
      files: drawFiles.value,
      sourceVersionId: currentVer.value?.id,
      note: drawNote.value,
    })
    drawFiles.value = []
    ElMessage.info('已提交抽签后赛程更新任务')
  } catch (err) {
    ElMessage.error(err.message || '提交抽签更新失败')
  }
}

function openEditDialog() {
  if (!active.value) return
  editName.value = active.value.name
  editTemplateId.value = active.value.template_id
  editDialogOpen.value = true
}

async function handleSaveEdit() {
  if (!active.value || !editName.value.trim()) {
    ElMessage.warning('请填写项目名称')
    return
  }
  savingEdit.value = true
  try {
    await projectsStore.updateProject(active.value.id, {
      name: editName.value,
      templateId: editTemplateId.value,
    })
    editDialogOpen.value = false
    ElMessage.success('项目设置已更新')
  } catch (err) {
    ElMessage.error(err.message || '更新失败')
  } finally {
    savingEdit.value = false
  }
}

function formatTraceDetail(ev) {
  const clone = { ...ev }
  delete clone.ts
  delete clone.stage
  if (clone.message) return clone.message
  if (clone.detail) return clone.detail
  const keys = Object.keys(clone)
  if (!keys.length) return '完成阶段处理'
  return keys.map((k) => `${k}: ${typeof clone[k] === 'string' ? clone[k] : JSON.stringify(clone[k])}`).join(' · ')
}
</script>

<style scoped>
.workspace-grid {
  display: grid;
  grid-template-columns: 310px minmax(0, 1fr);
  gap: 20px;
  align-items: start;
}

.sidebar-col {
  position: sticky;
  top: 84px;
}

.sidebar-card {
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.sidebar-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.project-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: calc(100vh - 280px);
  overflow-y: auto;
}

.project-item {
  padding: 11px 12px;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  background: #ffffff;
  cursor: pointer;
  transition: all 150ms ease;
}

.project-item:hover {
  border-color: var(--accent-border);
  background: var(--bg-canvas);
}

.project-item.is-active {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.project-item-top {
  display: flex;
  align-items: center;
  gap: 8px;
}

.project-item-name {
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text-primary);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.project-item-meta {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 5px;
  margin-top: 5px;
  padding-left: 16px;
  font-size: 11.5px;
  color: var(--text-muted);
}

.project-item-stage {
  margin-left: auto;
  color: var(--text-secondary);
}

.main-col {
  display: flex;
  flex-direction: column;
  gap: 18px;
  min-width: 0;
}

.card-pad {
  padding: 20px 22px;
}

.empty-workspace {
  padding: 48px 24px;
}

.proj-header-row {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  flex-wrap: wrap;
}

.proj-title-line {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.proj-title {
  margin: 0;
  font-size: 20px;
  font-weight: 700;
  color: var(--text-primary);
}

.proj-sub-meta {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-top: 8px;
  flex-wrap: wrap;
}

.status-pill-inner {
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.running-banner {
  margin-top: 14px;
  padding: 12px 14px;
  background: var(--accent-soft);
  border: 1px solid var(--accent-border);
  border-radius: var(--radius-md);
  display: flex;
  align-items: center;
  gap: 12px;
}

.running-copy {
  display: flex;
  flex-direction: column;
  font-size: 13px;
}

.running-copy strong {
  color: var(--accent);
}

.running-copy span {
  color: var(--text-secondary);
}

.card-head-row {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 14px;
  flex-wrap: wrap;
}

.role-upload-controls {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.dropzone {
  margin-top: 14px;
  padding: 20px 16px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-md);
  background: #ffffff;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 14px;
  cursor: pointer;
  transition: all 150ms ease;
}

.dropzone:hover,
.dropzone.is-over {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.dropzone-text {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.dropzone-text strong {
  font-size: 13.5px;
  color: var(--text-primary);
}

.dropzone-text span {
  font-size: 12px;
  color: var(--text-muted);
}

.materials-list {
  margin-top: 14px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.material-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 12px;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  background: #ffffff;
  transition: border-color 150ms ease;
}

.material-row:hover {
  border-color: var(--border-strong);
}

.material-left {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
  flex: 1;
}

.material-info {
  min-width: 0;
}

.material-name-line {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.material-name {
  font-size: 13.5px;
  font-weight: 500;
  color: var(--text-primary);
  word-break: break-all;
}

.material-actions {
  display: flex;
  align-items: center;
  gap: 2px;
  flex-shrink: 0;
}

.action-anchor {
  text-decoration: none;
}

.gen-console {
  margin-top: 18px;
  padding-top: 16px;
  border-top: 1px solid var(--border-hairline);
  display: flex;
  align-items: flex-start;
  gap: 12px;
  flex-wrap: wrap;
}

.gen-input-wrap {
  flex: 1;
  min-width: 260px;
}

.quick-chips {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: 8px;
}

.chip-btn {
  border: 1px solid var(--border-hairline);
  background: var(--bg-canvas);
  color: var(--text-secondary);
  border-radius: 999px;
  padding: 2px 10px;
  font-size: 11.5px;
  cursor: pointer;
  transition: all 140ms ease;
}

.chip-btn:hover {
  border-color: var(--accent-border);
  color: var(--accent);
  background: var(--accent-soft);
}

.gen-submit-btn {
  flex-shrink: 0;
  padding-inline: 22px;
}

.trace-collapse {
  margin-top: 14px;
  border-top: 1px solid var(--border-hairline);
  border-bottom: 0;
}

.trace-summary-title {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
  color: var(--text-secondary);
}

.trace-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-height: 240px;
  overflow-y: auto;
}

.trace-item {
  display: flex;
  align-items: baseline;
  gap: 8px;
  font-size: 12.5px;
  padding: 4px 0;
}

.trace-ts {
  color: var(--text-muted);
  font-size: 11.5px;
}

.trace-detail {
  color: var(--text-secondary);
  word-break: break-word;
}

.audit-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 12px;
  margin-bottom: 16px;
}

.audit-card {
  padding: 14px 16px;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  background: var(--bg-canvas);
}

.audit-card-head {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
  font-weight: 600;
  color: var(--text-primary);
  margin-bottom: 10px;
}

.audit-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
  font-size: 12.5px;
  color: var(--text-secondary);
  max-height: 190px;
  overflow-y: auto;
}

.audit-list li {
  display: flex;
  justify-content: space-between;
  gap: 10px;
}

.toc-page {
  color: var(--text-muted);
  flex-shrink: 0;
}

.warn-item {
  color: var(--warning);
}

.ok-empty {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: var(--success);
  font-size: 13px;
  font-weight: 500;
}

.draw-stat {
  margin: 0 0 6px;
  font-size: 13px;
  color: var(--text-primary);
}

.draw-inline-panel {
  padding: 14px 16px;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  margin-bottom: 16px;
}

.draw-inline-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  flex-wrap: wrap;
  margin-bottom: 10px;
}

.draw-inline-controls {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.draw-picked-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: 10px;
}

.accent-icon {
  color: var(--accent);
}

.success-icon {
  color: var(--success);
}

.warning-icon {
  color: var(--warning);
}

.muted-icon {
  color: var(--text-muted);
}

.spin-icon {
  animation: spin 1s linear infinite;
}

@keyframes spin {
  from {
    transform: rotate(0deg);
  }
  to {
    transform: rotate(360deg);
  }
}

@media (max-width: 980px) {
  .workspace-grid {
    grid-template-columns: 1fr;
  }
  .sidebar-col {
    position: static;
  }
}
</style>
