<template>
  <div class="page-container">
    <div class="page-header">
      <div>
        <h1 class="page-title">抽签后赛程更新</h1>
        <p class="page-subtitle">
          针对赛前按抽签签位（如 A1 VS A2）编排的秩序册 Word（.docx），上传抽签结果表后自动核对赛程上下文，精准替换为参赛队名，并完好保留“68胜——69胜”与交叉排名引用。
        </p>
      </div>
      <el-button v-if="drawStore.job" plain @click="handleReset">
        <RotateCcw :size="15" />
        <span style="margin-left: 6px">重置任务</span>
      </el-button>
    </div>

    <div class="white-card setup-card">
      <div class="steps-grid">
        <!-- Step 1: Target Word Book -->
        <div class="step-box">
          <div class="step-badge">1</div>
          <h2 class="step-title">上传待更新秩序册 Word</h2>
          <p class="muted-text">必须为 .docx 格式，系统将保留原稿样式就地替换</p>

          <div
            class="drop-box"
            :class="{ 'is-over': draggingBook, 'has-file': Boolean(bookFile) }"
            @dragenter.prevent="draggingBook = true"
            @dragover.prevent="draggingBook = true"
            @dragleave.prevent="draggingBook = false"
            @drop.prevent="handleDropBook"
            @click="bookInputRef?.click()"
          >
            <FileText :size="24" class="accent-icon" />
            <div class="drop-copy">
              <strong>{{ bookFile ? bookFile.name : '点击或拖入秩序册 (.docx)' }}</strong>
              <span>{{ bookFile ? formatSize(bookFile.size) : '支持单个 .docx 文件' }}</span>
            </div>
          </div>
          <input ref="bookInputRef" type="file" accept=".docx" hidden @change="handleBookChange" />
        </div>

        <!-- Step 2: Draw Result Files -->
        <div class="step-box">
          <div class="step-badge">2</div>
          <h2 class="step-title">上传抽签结果表</h2>
          <p class="muted-text">支持多份 Excel、Word、PDF 或抽签结果图片</p>

          <div
            class="drop-box"
            :class="{ 'is-over': draggingDraw }"
            @dragenter.prevent="draggingDraw = true"
            @dragover.prevent="draggingDraw = true"
            @dragleave.prevent="draggingDraw = false"
            @drop.prevent="handleDropDraw"
            @click="drawInputRef?.click()"
          >
            <UploadCloud :size="24" class="accent-icon" />
            <div class="drop-copy">
              <strong>点击或拖入抽签结果表（可多选）</strong>
              <span>{{ drawFiles.length ? `已添加 ${drawFiles.length} 份文件` : 'Excel / Word / PDF / 图片' }}</span>
            </div>
          </div>
          <input ref="drawInputRef" type="file" multiple hidden @change="handleDrawChange" />

          <div v-if="drawFiles.length" class="picked-files">
            <div v-for="(f, idx) in drawFiles" :key="idx" class="picked-chip">
              <FileKindIcon :filename="f.name" :size="14" />
              <span class="picked-name">{{ f.name }}</span>
              <button type="button" class="remove-chip" @click="drawFiles.splice(idx, 1)">
                <X :size="13" />
              </button>
            </div>
          </div>
        </div>
      </div>

      <!-- Step 3: Note & Submit -->
      <div class="submit-bar">
        <div class="note-col">
          <el-input
            v-model="note"
            placeholder="可选更新要求备注，例如：区分部门组与街镇组 A1/A2，保留淘汰赛胜负关系"
            clearable
          >
            <template #prefix>
              <MessageSquare :size="15" class="muted-icon" />
            </template>
          </el-input>
        </div>
        <el-button
          type="primary"
          size="large"
          :disabled="!bookFile || !drawFiles.length || drawStore.submitting"
          :loading="drawStore.submitting"
          @click="handleSubmit"
        >
          <Shuffle v-if="!drawStore.submitting" :size="16" />
          <span style="margin-left: 6px">
            {{ drawStore.submitting ? '正在核对赛程并替换签位…' : '开始执行抽签后更新' }}
          </span>
        </el-button>
      </div>

      <div v-if="drawStore.job?.status === 'running'" class="running-note">
        <Loader2 :size="16" class="spin-icon accent-icon" />
        <span>已提交任务 {{ drawStore.job.id }}，正在读取抽签表并核对赛程上下文…</span>
      </div>

      <el-alert
        v-if="drawStore.error"
        type="error"
        :title="drawStore.error"
        show-icon
        :closable="false"
        style="margin-top: 14px"
      />
    </div>

    <!-- Result Workspace -->
    <div v-if="drawStore.job?.status === 'done'" class="white-card result-card">
      <div class="result-head">
        <div>
          <div class="result-title-row">
            <CheckCircle2 :size="18" class="success-icon" />
            <h2>抽签更新完成</h2>
          </div>
          <p class="result-summary">
            已替换 <strong>{{ drawStore.job.replaced_count }}</strong> 处抽签签位，完好保留
            <strong>{{ drawStore.job.preserved_count }}</strong> 处胜负/排名引用。
            <span v-if="drawStore.job.agent_summary">{{ drawStore.job.agent_summary }}</span>
          </p>
        </div>

        <div class="kpi-strip">
          <div class="kpi-pill is-blue">
            <span class="kpi-num tabular">{{ drawStore.job.replaced_count ?? 0 }}</span>
            <span class="kpi-lbl">已替换签位</span>
          </div>
          <div class="kpi-pill is-green">
            <span class="kpi-num tabular">{{ drawStore.job.preserved_count ?? 0 }}</span>
            <span class="kpi-lbl">保留引用</span>
          </div>
        </div>
      </div>

      <div class="diff-tables">
        <div class="diff-panel">
          <h3 class="section-title" style="margin-bottom: 10px">
            <ArrowRightLeft :size="15" class="accent-icon" />
            <span>已替换赛程签位（前 40 条）</span>
          </h3>
          <el-table
            :data="(drawStore.job.changes || []).slice(0, 40)"
            size="small"
            border
            max-height="280"
            empty-text="无签位改动"
          >
            <el-table-column prop="location" label="位置" min-width="110" />
            <el-table-column prop="before" label="原签位" min-width="90" />
            <el-table-column label="替换为" min-width="110">
              <template #default="{ row }">
                <strong style="color: var(--accent)">{{ row.after }}</strong>
              </template>
            </el-table-column>
          </el-table>
        </div>

        <div class="diff-panel">
          <h3 class="section-title" style="margin-bottom: 10px">
            <ShieldCheck :size="15" class="success-icon" />
            <span>保持不动的胜负 / 排名引用</span>
          </h3>
          <el-table
            :data="(drawStore.job.preserved || []).slice(0, 30)"
            size="small"
            border
            max-height="280"
            empty-text="无特殊引用"
          >
            <el-table-column prop="location" label="位置" min-width="100" />
            <el-table-column prop="text" label="保留文本" min-width="110" />
            <el-table-column prop="reason" label="原因" min-width="120" />
          </el-table>
        </div>
      </div>

      <PdfPreviewCard
        title="抽签更新版秩序册预览"
        :preview-url="api.drawJobPdfPreviewUrl(drawStore.job.id)"
        :docx-url="api.drawJobDocxUrl(drawStore.job.id)"
        :pdf-url="api.drawJobPdfUrl(drawStore.job.id)"
      />
    </div>
  </div>
</template>

<script setup>
import { onBeforeUnmount, ref } from 'vue'
import { ElMessage } from 'element-plus'
import {
  ArrowRightLeft,
  CheckCircle2,
  FileText,
  Loader2,
  MessageSquare,
  RotateCcw,
  ShieldCheck,
  Shuffle,
  UploadCloud,
  X,
} from 'lucide-vue-next'
import { api } from '@/api/client'
import FileKindIcon from '@/components/FileKindIcon.vue'
import PdfPreviewCard from '@/components/PdfPreviewCard.vue'
import { useDrawStore } from '@/stores/draw'
import { formatSize } from '@/utils/format'

const drawStore = useDrawStore()

const bookFile = ref(null)
const drawFiles = ref([])
const note = ref('')
const draggingBook = ref(false)
const draggingDraw = ref(false)

const bookInputRef = ref(null)
const drawInputRef = ref(null)

onBeforeUnmount(() => {
  drawStore.stopPolling()
})

function handleBookChange(e) {
  const f = e.target.files?.[0]
  if (f) bookFile.value = f
  e.target.value = ''
}

function handleDropBook(e) {
  draggingBook.value = false
  const f = e.dataTransfer?.files?.[0]
  if (f) bookFile.value = f
}

function handleDrawChange(e) {
  const list = Array.from(e.target.files || [])
  if (list.length) drawFiles.value = [...drawFiles.value, ...list]
  e.target.value = ''
}

function handleDropDraw(e) {
  draggingDraw.value = false
  const list = Array.from(e.dataTransfer?.files || [])
  if (list.length) drawFiles.value = [...drawFiles.value, ...list]
}

async function handleSubmit() {
  if (!bookFile.value || !drawFiles.value.length) return
  try {
    await drawStore.startJob({
      bookFile: bookFile.value,
      drawFiles: drawFiles.value,
      note: note.value,
    })
    ElMessage.info('已提交抽签后赛程更新任务')
  } catch (err) {
    ElMessage.error(err.message || '提交失败')
  }
}

function handleReset() {
  drawStore.resetJob()
  bookFile.value = null
  drawFiles.value = []
  note.value = ''
}
</script>

<style scoped>
.setup-card {
  padding: 22px 24px;
  margin-bottom: 20px;
}

.steps-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 20px;
}

.step-box {
  padding: 18px;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  background: var(--bg-canvas);
}

.step-badge {
  width: 24px;
  height: 24px;
  border-radius: 50%;
  background: var(--accent);
  color: #ffffff;
  font-size: 12.5px;
  font-weight: 700;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-bottom: 8px;
}

.step-title {
  margin: 0 0 4px;
  font-size: 15px;
  font-weight: 600;
  color: var(--text-primary);
}

.drop-box {
  margin-top: 12px;
  padding: 18px 14px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-md);
  background: #ffffff;
  display: flex;
  align-items: center;
  gap: 12px;
  cursor: pointer;
  transition: all 150ms ease;
}

.drop-box:hover,
.drop-box.is-over {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.drop-box.has-file {
  border-style: solid;
  border-color: var(--accent-border);
}

.drop-copy {
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.drop-copy strong {
  font-size: 13.5px;
  color: var(--text-primary);
  word-break: break-all;
}

.drop-copy span {
  font-size: 12px;
  color: var(--text-muted);
}

.picked-files {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-top: 12px;
}

.picked-chip {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  padding: 4px 10px 4px 6px;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: 999px;
  font-size: 12.5px;
}

.picked-name {
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.remove-chip {
  border: 0;
  background: transparent;
  color: var(--text-muted);
  cursor: pointer;
  display: inline-flex;
  padding: 2px;
}

.remove-chip:hover {
  color: var(--danger);
}

.submit-bar {
  margin-top: 18px;
  padding-top: 18px;
  border-top: 1px solid var(--border-hairline);
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}

.note-col {
  flex: 1;
  min-width: 240px;
}

.running-note {
  margin-top: 14px;
  padding: 10px 14px;
  background: var(--accent-soft);
  border: 1px solid var(--accent-border);
  border-radius: var(--radius-md);
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: var(--accent);
}

.result-card {
  padding: 22px 24px;
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.result-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  flex-wrap: wrap;
}

.result-title-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.result-title-row h2 {
  margin: 0;
  font-size: 18px;
  font-weight: 700;
}

.result-summary {
  margin: 6px 0 0;
  font-size: 13.5px;
  color: var(--text-secondary);
}

.kpi-strip {
  display: flex;
  gap: 10px;
}

.kpi-pill {
  padding: 8px 16px;
  border-radius: var(--radius-md);
  display: flex;
  flex-direction: column;
  align-items: center;
  min-width: 96px;
  border: 1px solid var(--border-hairline);
}

.kpi-pill.is-blue {
  background: var(--accent-soft);
  border-color: var(--accent-border);
  color: var(--accent);
}

.kpi-pill.is-green {
  background: var(--success-soft);
  border-color: #a7f3d0;
  color: var(--success);
}

.kpi-num {
  font-size: 20px;
  font-weight: 700;
  line-height: 1.1;
}

.kpi-lbl {
  font-size: 11.5px;
  margin-top: 2px;
}

.diff-tables {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 16px;
}

.diff-panel {
  min-width: 0;
}

.accent-icon {
  color: var(--accent);
}

.success-icon {
  color: var(--success);
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

@media (max-width: 900px) {
  .steps-grid,
  .diff-tables {
    grid-template-columns: 1fr;
  }
}
</style>
