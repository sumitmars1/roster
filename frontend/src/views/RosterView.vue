<template>
  <div class="page-container">
    <div class="page-header">
      <div>
        <h1 class="page-title">单表秩序册名单整理</h1>
        <p class="page-subtitle">
          汇总各类报名表格或扫描件，由 DeepSeek Harness 按需读取、校验冲突并直接生成可编辑的 Word 名单文件。
        </p>
      </div>
    </div>

    <div class="roster-setup-grid">
      <!-- Left: Upload Panel -->
      <section class="white-card panel-pad">
        <div class="panel-head">
          <div>
            <h2 class="section-title">
              <FileSpreadsheet :size="16" class="accent-icon" />
              <span>报名资料</span>
            </h2>
            <span class="muted-text tabular" style="margin-left: 8px">
              {{
                roster.selectedFiles.length
                  ? `已选 ${roster.selectedCount} / ${roster.selectedFiles.length} 份`
                  : '已选 0 份'
              }}
            </span>
          </div>
          <el-button size="small" plain @click="fileInputRef?.click()">
            <Plus :size="14" />
            <span style="margin-left: 4px">添加文件</span>
          </el-button>
        </div>

        <input
          ref="fileInputRef"
          type="file"
          multiple
          :accept="acceptExtensions"
          hidden
          @change="handleFileInput"
        />

        <div
          class="roster-dropzone"
          :class="{ 'is-over': dragging, 'is-compact': roster.selectedFiles.length > 0 }"
          @dragenter.prevent="dragging = true"
          @dragover.prevent="dragging = true"
          @dragleave.prevent="dragging = false"
          @drop.prevent="handleDrop"
          @click="fileInputRef?.click()"
        >
          <UploadCloud :size="22" class="accent-icon" />
          <div class="drop-copy">
            <strong>选择或拖入报名资料</strong>
            <span>{{ uploadLimitText }}</span>
          </div>
        </div>

        <el-alert
          v-if="roster.uploadNote"
          type="warning"
          :title="roster.uploadNote"
          show-icon
          closable
          style="margin-top: 10px"
          @close="roster.uploadNote = ''"
        />

        <div v-if="roster.selectedFiles.length" class="roster-file-list">
          <div v-for="entry in roster.selectedFiles" :key="entry.uid" class="roster-file-row">
            <el-checkbox v-model="entry.selected" />
            <FileKindIcon :filename="entry.file.name" :size="15" />
            <div class="roster-file-info">
              <span class="roster-file-name">{{ entry.file.name }}</span>
              <span class="muted-text tabular">{{ formatSize(entry.file.size) }}</span>
            </div>
            <el-tag size="small" effect="plain" type="info">
              {{ getExtension(entry.file.name).toUpperCase() }}
            </el-tag>
            <el-button size="small" text type="danger" @click="roster.removeFile(entry.uid)">
              <X :size="14" />
            </el-button>
          </div>
        </div>
      </section>

      <!-- Right: Settings & Preset Picker -->
      <section class="white-card panel-pad">
        <div class="panel-head">
          <h2 class="section-title">
            <SlidersHorizontal :size="16" class="accent-icon" />
            <span>输出设置与名单版式</span>
          </h2>
        </div>

        <div class="settings-inputs">
          <div class="input-field">
            <label class="field-label">赛事名称 <span class="req">*</span></label>
            <el-input
              v-model="roster.title"
              placeholder="如：职工乒乓球赛"
              maxlength="100"
              clearable
            />
          </div>
          <div class="input-field">
            <label class="field-label">组别 <span class="muted-text">（选填）</span></label>
            <el-input
              v-model="roster.group"
              placeholder="如：混合双打组"
              maxlength="70"
              clearable
            />
          </div>
        </div>

        <div v-if="presets.length" class="preset-section">
          <div class="preset-label-row">
            <span class="field-label">选择名单版式</span>
            <span class="muted-text">{{ activePresetObj?.description }}</span>
          </div>

          <div class="preset-layout">
            <div class="preset-list">
              <button
                v-for="preset in presets"
                :key="preset.id"
                type="button"
                class="preset-btn"
                :class="{ 'is-active': roster.activePreset === preset.id }"
                @click="roster.activePreset = preset.id"
              >
                <span class="preset-dot" />
                <div class="preset-btn-copy">
                  <strong>{{ preset.name }}</strong>
                  <span>{{ preset.description }}</span>
                </div>
              </button>
            </div>

            <div v-if="activePresetObj" class="preset-preview-card">
              <div class="preview-card-head">
                <strong>{{ activePresetObj.name }}</strong>
                <el-button size="small" text type="primary" @click="presetZoomOpen = true">
                  <ZoomIn :size="14" />
                  <span style="margin-left: 4px">放大示例</span>
                </el-button>
              </div>
              <div class="preview-img-frame" @click="presetZoomOpen = true">
                <img
                  :src="`/static/previews/${encodeURIComponent(activePresetObj.id)}.png`"
                  :alt="`${activePresetObj.name}示例`"
                />
              </div>
            </div>
          </div>
        </div>

        <div class="start-row">
          <el-button
            type="primary"
            size="large"
            style="width: 100%"
            :disabled="!roster.canStart"
            :loading="roster.submitting"
            @click="handleStartTask"
          >
            <span>{{ roster.submitting ? '正在上传资料并启动…' : '开始整理名单' }}</span>
            <ArrowRight v-if="!roster.submitting" :size="16" style="margin-left: 6px" />
          </el-button>
          <el-alert
            v-if="roster.submitError"
            type="error"
            :title="roster.submitError"
            show-icon
            :closable="false"
            style="margin-top: 12px"
          />
        </div>
      </section>
    </div>

    <!-- Active Task Section -->
    <section v-if="task" class="white-card task-card">
      <div class="task-header">
        <div>
          <div class="task-title-line">
            <h2 class="task-title">整理任务</h2>
            <el-tag :type="rosterStatusMeta(task.status).type" effect="light" round>
              <span class="tag-inner">
                <Loader2 v-if="isBusy" :size="13" class="spin-icon" />
                <CheckCircle2 v-else-if="task.status === 'done'" :size="13" />
                <AlertTriangle v-else-if="task.status === 'review'" :size="13" />
                <AlertCircle v-else-if="task.status === 'error'" :size="13" />
                <span>{{ rosterStatusMeta(task.status).label }}</span>
              </span>
            </el-tag>
          </div>
          <p class="muted-text" style="margin: 6px 0 0">{{ taskDetailSummary }}</p>
        </div>

        <div class="task-actions">
          <el-button v-if="task.status === 'error'" plain @click="handleRestart">
            <RotateCcw :size="14" />
            <span style="margin-left: 5px">重新整理</span>
          </el-button>
          <a v-if="task.word_available" :href="api.taskPlacementUrl(task.id)" download class="btn-anchor">
            <el-button plain>
              <FileJson :size="14" />
              <span style="margin-left: 5px">落位 JSON</span>
            </el-button>
          </a>
          <a v-if="task.word_available" :href="api.taskWordUrl(task.id)" download class="btn-anchor">
            <el-button type="primary">
              <Download :size="15" />
              <span style="margin-left: 6px">下载 Word 名单</span>
            </el-button>
          </a>
        </div>
      </div>

      <el-progress
        v-if="isBusy || task.status === 'done'"
        :percentage="task.status === 'done' ? 100 : 45"
        :indeterminate="isBusy"
        :status="task.status === 'done' ? 'success' : ''"
        :show-text="false"
        style="margin-top: 14px"
      />

      <el-alert
        v-if="task.error"
        type="error"
        :title="task.error"
        show-icon
        :closable="false"
        style="margin-top: 14px"
      />

      <!-- Review Box (Human-in-the-loop) -->
      <div v-if="task.review" class="review-box">
        <div class="review-head">
          <div class="review-title">
            <HelpCircle :size="18" class="warning-icon" />
            <h3>需要你核对</h3>
          </div>
          <el-button
            v-if="primarySourceFile"
            size="small"
            plain
            @click="sourceDialogOpen = true"
          >
            <Eye :size="14" />
            <span style="margin-left: 5px">查看报名表原件</span>
          </el-button>
        </div>

        <p class="review-question">{{ task.review.question }}</p>

        <div v-if="reviewSources.length" class="review-sources">
          <el-tag
            v-for="(src, idx) in reviewSources"
            :key="idx"
            size="small"
            effect="plain"
            type="warning"
          >
            {{ src }}
          </el-tag>
        </div>

        <div v-if="(task.review.options || []).length" class="review-options">
          <button
            v-for="(opt, idx) in task.review.options"
            :key="idx"
            type="button"
            class="review-option-btn"
            :disabled="roster.reviewSubmitting"
            @click="handleSubmitReview(opt)"
          >
            <span class="opt-letter">{{ 'ABCDEF'[idx] || idx + 1 }}</span>
            <span class="opt-text">{{ opt }}</span>
            <ChevronRight :size="15" class="muted-icon" />
          </button>
        </div>

        <div class="review-reply-row">
          <el-input
            v-model="reviewAnswer"
            type="textarea"
            :rows="2"
            maxlength="2500"
            placeholder="或直接输入你的核对答复…"
            :disabled="roster.reviewSubmitting"
          />
          <el-button
            type="primary"
            :disabled="!reviewAnswer.trim() || roster.reviewSubmitting"
            :loading="roster.reviewSubmitting"
            @click="handleSubmitReview(reviewAnswer)"
          >
            <Send :size="15" />
            <span style="margin-left: 5px">发送答复</span>
          </el-button>
        </div>

        <el-alert
          v-if="roster.reviewError"
          type="error"
          :title="roster.reviewError"
          show-icon
          :closable="false"
          style="margin-top: 10px"
        />
      </div>

      <!-- Batch Files Progress -->
      <div v-if="(task.files || []).length" class="batch-section">
        <h3 class="section-title" style="font-size: 14px; margin-bottom: 10px">
          <Files :size="15" class="accent-icon" />
          <span>资料处理进度</span>
        </h3>
        <div class="batch-grid">
          <div v-for="file in task.files" :key="file.id" class="batch-item">
            <FileKindIcon :filename="file.name" :size="15" />
            <div class="batch-copy">
              <strong>{{ file.name }}</strong>
              <span class="muted-text">{{ fileDetailText(file) }}</span>
              <span v-if="file.error" class="batch-err">{{ file.error }}</span>
            </div>
            <el-button
              v-if="file.status === 'error' && task.status === 'error'"
              size="small"
              plain
              @click="handleRetryFile(file.id)"
            >
              重试
            </el-button>
          </div>
        </div>
      </div>

      <!-- Trace Log -->
      <el-collapse v-if="(task.trace || []).length" style="margin-top: 14px; border-bottom: 0">
        <el-collapse-item name="log">
          <template #title>
            <span class="trace-head">
              <Terminal :size="14" />
              <span>处理记录（{{ task.trace.length }} 条）</span>
            </span>
          </template>
          <div class="trace-list">
            <div v-for="(ev, idx) in task.trace.slice(-80)" :key="idx" class="trace-row">
              <el-tag
                size="small"
                :type="ev.error ? 'danger' : ev.kind === 'review' ? 'warning' : 'info'"
                effect="plain"
              >
                {{ ev.kind === 'error' ? '运行错误' : toolNames[ev.name] || ev.name || '处理' }}
              </el-tag>
              <span class="trace-msg">{{ ev.error || ev.detail || '' }}</span>
            </div>
          </div>
        </el-collapse-item>
      </el-collapse>
    </section>

    <!-- Preset Zoom Dialog -->
    <el-dialog
      v-model="presetZoomOpen"
      :title="`${activePresetObj?.name || ''} 版式示例`"
      width="780px"
      append-to-body
    >
      <div v-if="activePresetObj" class="zoom-img-wrap">
        <img
          :src="`/static/previews/${encodeURIComponent(activePresetObj.id)}.png`"
          :alt="activePresetObj.name"
          class="zoom-img"
        />
      </div>
    </el-dialog>

    <!-- Source Evidence Dialog -->
    <el-dialog v-model="sourceDialogOpen" title="报名表原件预览" width="820px" append-to-body>
      <div v-if="task" class="sources-stack">
        <div v-for="(src, idx) in reviewSources" :key="idx" class="source-item">
          <div class="source-item-head">
            <strong>{{ src }}</strong>
            <a
              v-if="parseSource(src).file"
              :href="api.taskSourceUrl(task.id, parseSource(src).file.id)"
              download
              class="btn-anchor"
            >
              <el-button size="small" plain>下载原件</el-button>
            </a>
          </div>
          <iframe
            v-if="parseSource(src).file && /\.pdf$/i.test(parseSource(src).name)"
            :src="
              api.taskSourceUrl(task.id, parseSource(src).file.id) +
              (parseSource(src).page ? `#page=${parseSource(src).page}` : '')
            "
            class="source-iframe"
          />
          <img
            v-else-if="parseSource(src).file && /\.(png|jpe?g|webp)$/i.test(parseSource(src).name)"
            :src="api.taskSourceUrl(task.id, parseSource(src).file.id)"
            class="source-img"
          />
        </div>
      </div>
    </el-dialog>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import { ElMessage } from 'element-plus'
import {
  AlertCircle,
  AlertTriangle,
  ArrowRight,
  CheckCircle2,
  ChevronRight,
  Download,
  Eye,
  FileJson,
  FileSpreadsheet,
  Files,
  HelpCircle,
  Loader2,
  Plus,
  RotateCcw,
  Send,
  SlidersHorizontal,
  Terminal,
  UploadCloud,
  X,
  ZoomIn,
} from 'lucide-vue-next'
import { api } from '@/api/client'
import FileKindIcon from '@/components/FileKindIcon.vue'
import { useRosterStore } from '@/stores/roster'
import { formatSize, getExtension, rosterStatusMeta } from '@/utils/format'

const roster = useRosterStore()

const fileInputRef = ref(null)
const dragging = ref(false)
const presetZoomOpen = ref(false)
const sourceDialogOpen = ref(false)
const reviewAnswer = ref('')

const toolNames = {
  list_documents: '查看资料目录',
  read_document: '读取资料',
  submit_roster: '提交名单',
}

const presets = computed(() => roster.options?.presets || [])
const activePresetObj = computed(() =>
  presets.value.find((p) => p.id === roster.activePreset) || presets.value[0] || null
)
const acceptExtensions = computed(() => (roster.options?.upload?.extensions || []).join(','))
const uploadLimitText = computed(() => {
  const maxMb = roster.options?.upload?.max_file_bytes
    ? Math.round(roster.options.upload.max_file_bytes / 1024 / 1024)
    : 300
  return `支持 Excel、Word、PDF 或图片 · 单份不超过 ${maxMb} MB`
})

const task = computed(() => roster.task)
const isBusy = computed(() => ['queued', 'running'].includes(task.value?.status))

const taskDetailSummary = computed(() => {
  if (!task.value) return ''
  const files = task.value.files || []
  const done = files.filter((f) => f.status === 'done').length
  const progress = files.length > 1 ? `${done}/${files.length} 份完成` : ''
  return [task.value.title, task.value.group, task.value.preset_name || task.value.preset, progress]
    .filter(Boolean)
    .join(' · ')
})

const reviewSources = computed(() => task.value?.review?.sources || [])
const primarySourceFile = computed(() => {
  const first = reviewSources.value[0]
  return first ? parseSource(first).file : null
})

onMounted(async () => {
  try {
    await roster.fetchOptions()
    const savedId = roster.getSavedTaskId()
    if (savedId && !roster.task) {
      await roster.restoreTask(savedId)
    }
  } catch {
    // handled in store
  }
})

onBeforeUnmount(() => {
  roster.stopPolling()
})

function handleFileInput(e) {
  if (e.target.files?.length) {
    roster.addFiles(Array.from(e.target.files))
  }
  e.target.value = ''
}

function handleDrop(e) {
  dragging.value = false
  if (e.dataTransfer?.files?.length) {
    roster.addFiles(Array.from(e.dataTransfer.files))
  }
}

async function handleStartTask() {
  try {
    await roster.startTask()
    ElMessage.info('已启动名单整理任务')
  } catch (err) {
    ElMessage.error(err.message || '提交失败')
  }
}

async function handleSubmitReview(ans) {
  try {
    await roster.submitReview(ans)
    reviewAnswer.value = ''
    ElMessage.success('已发送核对答复，继续处理')
  } catch (err) {
    ElMessage.error(err.message || '提交答复失败')
  }
}

async function handleRetryFile(fileId) {
  try {
    await roster.retryFile(fileId)
    ElMessage.info('正在重试该文件')
  } catch (err) {
    ElMessage.error(err.message || '重试失败')
  }
}

async function handleRestart() {
  try {
    await roster.restartTask()
    ElMessage.info('已重新启动整理任务')
  } catch (err) {
    ElMessage.error(err.message || '重新整理失败')
  }
}

function fileDetailText(file) {
  const units = file.units || []
  const unitInfo =
    file.suffix === '.pdf'
      ? `${units.length} 页`
      : units.length === 1
      ? units[0]?.label || '1 个单元'
      : `${units.length} 个工作表`
  const states = {
    waiting: '待处理',
    running: '识别中',
    done: '已完成',
    error: '失败',
    review: '待核对',
  }
  return `${file.id} · ${unitInfo} · ${states[file.status] || '处理中'}`
}

function parseSource(source) {
  const files = task.value?.files || []
  const file = files.find((item) => source.startsWith(`${item.id} `))
  if (!file) return { file: null, name: '', page: null, location: source }
  const pageMatch = source.match(/第 (\d+) 页/)
  return {
    file,
    name: file.name,
    page: pageMatch ? Number(pageMatch[1]) : null,
    location: source,
  }
}
</script>

<style scoped>
.roster-setup-grid {
  display: grid;
  grid-template-columns: minmax(0, 1fr) minmax(0, 1.35fr);
  gap: 20px;
  align-items: start;
}

.panel-pad {
  padding: 20px 22px;
}

.panel-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 14px;
}

.roster-dropzone {
  padding: 28px 16px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-md);
  background: #ffffff;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  text-align: center;
  gap: 8px;
  cursor: pointer;
  transition: all 150ms ease;
}

.roster-dropzone.is-compact {
  padding: 14px 16px;
  flex-direction: row;
  text-align: left;
  justify-content: flex-start;
  gap: 12px;
}

.roster-dropzone:hover,
.roster-dropzone.is-over {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.drop-copy {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.drop-copy strong {
  font-size: 13.5px;
  color: var(--text-primary);
}

.drop-copy span {
  font-size: 12px;
  color: var(--text-muted);
}

.roster-file-list {
  margin-top: 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: 360px;
  overflow-y: auto;
}

.roster-file-row {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 10px;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-sm);
  background: #ffffff;
}

.roster-file-info {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
}

.roster-file-name {
  font-size: 13px;
  font-weight: 500;
  color: var(--text-primary);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.settings-inputs {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 14px;
}

.field-label {
  display: block;
  font-size: 13px;
  font-weight: 600;
  color: var(--text-primary);
  margin-bottom: 6px;
}

.req {
  color: var(--danger);
}

.preset-section {
  margin-top: 18px;
}

.preset-label-row {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 10px;
  flex-wrap: wrap;
  margin-bottom: 8px;
}

.preset-layout {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 260px;
  gap: 14px;
  align-items: stretch;
}

.preset-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.preset-btn {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  padding: 9px 11px;
  text-align: left;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-sm);
  cursor: pointer;
  transition: all 140ms ease;
}

.preset-btn:hover {
  border-color: var(--accent-border);
  background: var(--bg-canvas);
}

.preset-btn.is-active {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.preset-dot {
  width: 10px;
  height: 10px;
  border-radius: 50%;
  border: 2px solid var(--border-strong);
  margin-top: 5px;
  flex-shrink: 0;
}

.preset-btn.is-active .preset-dot {
  border-color: var(--accent);
  background: var(--accent);
}

.preset-btn-copy {
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.preset-btn-copy strong {
  font-size: 13px;
  color: var(--text-primary);
}

.preset-btn-copy span {
  font-size: 11.5px;
  color: var(--text-muted);
  line-height: 1.4;
}

.preset-preview-card {
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  background: var(--bg-canvas);
  padding: 10px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.preview-card-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-size: 12.5px;
}

.preview-img-frame {
  flex: 1;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-sm);
  overflow: hidden;
  cursor: zoom-in;
  display: flex;
  align-items: center;
  justify-content: center;
}

.preview-img-frame img {
  width: 100%;
  height: auto;
  display: block;
}

.start-row {
  margin-top: 18px;
}

.task-card {
  margin-top: 20px;
  padding: 22px 24px;
}

.task-header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 14px;
  flex-wrap: wrap;
}

.task-title-line {
  display: flex;
  align-items: center;
  gap: 10px;
}

.task-title {
  margin: 0;
  font-size: 18px;
  font-weight: 700;
}

.tag-inner {
  display: inline-flex;
  align-items: center;
  gap: 5px;
}

.task-actions {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.btn-anchor {
  text-decoration: none;
}

.review-box {
  margin-top: 16px;
  padding: 16px 18px;
  background: var(--warning-soft);
  border: 1px solid #fde68a;
  border-radius: var(--radius-md);
}

.review-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
}

.review-title {
  display: flex;
  align-items: center;
  gap: 8px;
}

.review-title h3 {
  margin: 0;
  font-size: 15px;
  font-weight: 700;
  color: #92400e;
}

.review-question {
  margin: 10px 0;
  font-size: 14px;
  color: var(--text-primary);
  line-height: 1.6;
}

.review-sources {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 12px;
}

.review-options {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 12px;
}

.review-option-btn {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  background: #ffffff;
  border: 1px solid #fcd34d;
  border-radius: var(--radius-sm);
  cursor: pointer;
  text-align: left;
  transition: all 140ms ease;
}

.review-option-btn:hover:not(:disabled) {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.opt-letter {
  width: 22px;
  height: 22px;
  border-radius: 6px;
  background: var(--accent-soft);
  color: var(--accent);
  font-weight: 700;
  font-size: 12px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
}

.opt-text {
  flex: 1;
  font-size: 13.5px;
  color: var(--text-primary);
}

.review-reply-row {
  display: flex;
  gap: 10px;
  align-items: flex-end;
}

.batch-section {
  margin-top: 18px;
}

.batch-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 10px;
}

.batch-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-sm);
  background: var(--bg-canvas);
}

.batch-copy {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
}

.batch-copy strong {
  font-size: 13px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.batch-err {
  font-size: 12px;
  color: var(--danger);
}

.trace-head {
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

.trace-row {
  display: flex;
  align-items: baseline;
  gap: 8px;
  font-size: 12.5px;
}

.trace-msg {
  color: var(--text-secondary);
  word-break: break-word;
}

.zoom-img-wrap {
  background: var(--bg-canvas);
  padding: 12px;
  border-radius: var(--radius-md);
  text-align: center;
}

.zoom-img {
  max-width: 100%;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border-hairline);
}

.sources-stack {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.source-item {
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  padding: 12px;
}

.source-item-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 10px;
}

.source-iframe {
  width: 100%;
  height: 460px;
  border: 0;
  border-radius: var(--radius-sm);
}

.source-img {
  max-width: 100%;
  border-radius: var(--radius-sm);
}

.accent-icon {
  color: var(--accent);
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

@media (max-width: 1020px) {
  .roster-setup-grid,
  .preset-layout,
  .settings-inputs {
    grid-template-columns: 1fr;
  }
}
</style>
