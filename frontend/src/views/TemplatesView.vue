<template>
  <div class="page-container">
    <div class="page-header">
      <div>
        <h1 class="page-title">秩序册模板库与封面工坊</h1>
        <p class="page-subtitle">
          内置多类公文与综合赛事版式；您也可以上传一份历史 Word 秩序册样例，系统将解析其章节编排、版式字体与固定底图（不保留历史人员与赛程数据），沉淀为可长期复用的模板。
        </p>
      </div>
    </div>

    <div class="tpl-workspace">
      <!-- Left Studio Panel -->
      <aside class="tpl-studio">
        <!-- Sample Parser Card -->
        <div class="white-card studio-card">
          <div class="studio-head">
            <span class="section-title">
              <FileUp :size="16" class="accent-icon" />
              <span>从 Word 样例解析入库</span>
            </span>
          </div>

          <el-form label-position="top" @submit.prevent="handleParseSample">
            <el-form-item label="模板名称（留空则自动命名）">
              <el-input
                v-model="sampleName"
                placeholder="如：全市系统运动会标准模板"
                maxlength="80"
                clearable
              />
            </el-form-item>

            <el-form-item label="样例秩序册 (.docx / .doc / .wps)" required>
              <div
                class="mini-dropzone"
                :class="{ 'is-over': draggingSample }"
                @dragenter.prevent="draggingSample = true"
                @dragover.prevent="draggingSample = true"
                @dragleave.prevent="draggingSample = false"
                @drop.prevent="handleDropSample"
                @click="sampleInputRef?.click()"
              >
                <UploadCloud :size="20" class="accent-icon" />
                <div class="mini-drop-copy">
                  <strong>{{ sampleFile ? sampleFile.name : '点击或拖入样例 Word 文件' }}</strong>
                  <span>{{ sampleFile ? formatSize(sampleFile.size) : '支持 .docx / .doc / .wps' }}</span>
                </div>
              </div>
              <input
                ref="sampleInputRef"
                type="file"
                accept=".docx,.doc,.wps"
                hidden
                @change="handleSampleFileChange"
              />
            </el-form-item>

            <div class="llm-switch-row">
              <span class="muted-text">启用大模型智能识别章节结构</span>
              <el-switch v-model="useLlm" />
            </div>

            <el-button
              type="primary"
              style="width: 100%; margin-top: 10px"
              :disabled="!sampleFile"
              :loading="templatesStore.parsing"
              native-type="submit"
            >
              <Sparkles :size="15" />
              <span style="margin-left: 6px">解析并保存模板</span>
            </el-button>

            <el-alert
              v-if="templatesStore.parseMessage && !templatesStore.parseError"
              type="success"
              :title="templatesStore.parseMessage"
              show-icon
              :closable="false"
              style="margin-top: 12px"
            />
            <el-alert
              v-if="templatesStore.parseError"
              type="error"
              :title="templatesStore.parseError"
              show-icon
              :closable="false"
              style="margin-top: 12px"
            />
          </el-form>
        </div>

        <!-- Cover Generator Card -->
        <div class="white-card studio-card">
          <div class="studio-head">
            <span class="section-title">
              <ImagePlus :size="16" class="accent-icon" />
              <span>A4 封面主画面合成</span>
            </span>
          </div>

          <el-form label-position="top" @submit.prevent="handleGenerateCover">
            <el-form-item label="赛事主标题" required>
              <el-input
                v-model="coverTitle"
                placeholder="如：2026年重庆市职工篮球赛"
                clearable
              />
            </el-form-item>

            <el-form-item label="副标题">
              <el-input v-model="coverSubtitle" placeholder="秩  序  册" clearable />
            </el-form-item>

            <el-form-item label="落款 / 主办单位或日期（每行一条）">
              <el-input
                v-model="coverFooter"
                type="textarea"
                :rows="2"
                placeholder="重庆市总工会&#10;2026年10月"
              />
            </el-form-item>

            <el-form-item label="主画面 / 底图（可选）">
              <div class="cover-file-row">
                <el-button plain @click="coverInputRef?.click()">
                  <ImageIcon :size="14" />
                  <span style="margin-left: 6px">
                    {{ coverBgFile ? coverBgFile.name : '选择背景图片' }}
                  </span>
                </el-button>
                <el-button
                  v-if="coverBgFile"
                  text
                  type="danger"
                  size="small"
                  @click="coverBgFile = null"
                >
                  清除
                </el-button>
              </div>
              <input
                ref="coverInputRef"
                type="file"
                accept=".png,.jpg,.jpeg,.webp"
                hidden
                @change="handleCoverBgChange"
              />
            </el-form-item>

            <el-button
              plain
              type="primary"
              style="width: 100%"
              :disabled="!coverTitle.trim()"
              :loading="templatesStore.coverGenerating"
              native-type="submit"
            >
              <Wand2 :size="15" />
              <span style="margin-left: 6px">合成 A4 封面图</span>
            </el-button>

            <el-alert
              v-if="templatesStore.coverError"
              type="error"
              :title="templatesStore.coverError"
              show-icon
              :closable="false"
              style="margin-top: 12px"
            />
          </el-form>

          <div v-if="templatesStore.coverPreviewUrl" class="cover-result">
            <img :src="templatesStore.coverPreviewUrl" alt="生成封面预览" class="cover-img" />
            <el-button
              tag="a"
              :href="templatesStore.coverPreviewUrl"
              download="封面.png"
              size="small"
              type="primary"
              plain
              class="dl-anchor"
              style="width: 100%"
            >
              <Download :size="14" />
              <span style="margin-left: 6px">下载封面 PNG</span>
            </el-button>
          </div>
        </div>
      </aside>

      <!-- Right Gallery -->
      <section class="tpl-gallery-col">
        <div class="white-card gallery-toolbar">
          <el-input
            v-model="filterKeyword"
            placeholder="搜索模板名称或说明…"
            clearable
            style="max-width: 300px"
          >
            <template #prefix>
              <Search :size="14" class="muted-icon" />
            </template>
          </el-input>

          <el-radio-group v-model="filterType" size="default">
            <el-radio-button value="all">全部 ({{ templatesStore.templates.length }})</el-radio-button>
            <el-radio-button value="builtin">内置模板</el-radio-button>
            <el-radio-button value="custom">自定义入库</el-radio-button>
          </el-radio-group>
        </div>

        <el-skeleton v-if="templatesStore.loading && !templatesStore.templates.length" :rows="6" animated />

        <el-empty
          v-else-if="!filteredTemplates.length"
          description="没有符合条件的模板"
          :image-size="80"
        />

        <div v-else class="tpl-cards-grid">
          <article v-for="tpl in filteredTemplates" :key="tpl.id" class="white-card tpl-card">
            <div class="tpl-card-head">
              <div class="tpl-title-wrap">
                <h3 class="tpl-name">{{ tpl.name }}</h3>
              </div>
              <el-tag
                size="small"
                :type="tpl.builtin ? 'info' : 'success'"
                effect="light"
                round
              >
                {{ tpl.builtin ? '内置' : '自定义入库' }}
              </el-tag>
            </div>

            <p class="tpl-desc">{{ tpl.description || '无描述' }}</p>

            <div class="tpl-spec-pills">
              <el-tag size="small" effect="plain" type="info">
                名单母版: {{ tpl.roster_preset || 'inline-roster' }}
              </el-tag>
              <el-tag size="small" effect="plain" type="info">
                正文: {{ tpl.formatting?.fonts?.body?.name || '仿宋' }}
                {{ tpl.formatting?.fonts?.body?.size_pt || 16 }}pt
              </el-tag>
              <el-tag v-if="tpl.formatting?.page?.width_mm" size="small" effect="plain" type="info">
                页面: {{ Math.round(tpl.formatting.page.width_mm) }}×{{ Math.round(tpl.formatting.page.height_mm) }}mm
              </el-tag>
              <el-tag v-if="tpl.formatting?.page_number_style" size="small" effect="plain" type="info">
                页码: {{ tpl.formatting.page_number_style }}
              </el-tag>
              <el-tag
                v-if="hasWatermark(tpl)"
                size="small"
                effect="light"
                type="primary"
                class="wm-tag"
                @click="openWatermarkPreview(tpl)"
              >
                <Eye :size="12" style="margin-right: 4px" />
                查看底图水印
              </el-tag>
            </div>

            <div class="tpl-chapter-flow">
              <span class="chapter-label">章节编排：</span>
              <div class="chapter-chips">
                <span
                  v-for="ch in visibleChapters(tpl)"
                  :key="ch.id"
                  class="chapter-chip"
                >
                  {{ ch.title }}
                </span>
                <span v-if="!visibleChapters(tpl).length" class="muted-text">自动识别编排</span>
              </div>
            </div>

            <div class="tpl-card-footer">
              <el-button
                size="small"
                type="primary"
                plain
                @click="useTemplateForProject(tpl.id)"
              >
                <FolderPlus :size="14" />
                <span style="margin-left: 5px">用此模板新建项目</span>
              </el-button>

              <el-popconfirm
                v-if="!tpl.builtin"
                title="确定删除这个自定义模板吗？"
                confirm-button-text="删除"
                cancel-button-text="取消"
                @confirm="handleDeleteTemplate(tpl.id)"
              >
                <template #reference>
                  <el-button size="small" text type="danger">
                    <Trash2 :size="14" />
                    <span style="margin-left: 4px">删除</span>
                  </el-button>
                </template>
              </el-popconfirm>
            </div>
          </article>
        </div>
      </section>
    </div>

    <!-- Watermark Lightbox Dialog -->
    <el-dialog v-model="watermarkDialogOpen" :title="`${watermarkTitle} · 模板底图`" width="560px" append-to-body>
      <div class="wm-preview-wrap">
        <img :src="watermarkUrl" alt="模板底图" class="wm-preview-img" />
      </div>
    </el-dialog>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'
import { ElMessage } from 'element-plus'
import {
  Download,
  Eye,
  FileUp,
  FolderPlus,
  Image as ImageIcon,
  ImagePlus,
  Search,
  Sparkles,
  Trash2,
  UploadCloud,
  Wand2,
} from 'lucide-vue-next'
import { api } from '@/api/client'
import { useTemplatesStore } from '@/stores/templates'
import { useUiStore } from '@/stores/ui'
import { formatSize } from '@/utils/format'

const templatesStore = useTemplatesStore()
const ui = useUiStore()

const sampleName = ref('')
const sampleFile = ref(null)
const useLlm = ref(true)
const draggingSample = ref(false)
const sampleInputRef = ref(null)

const coverTitle = ref('')
const coverSubtitle = ref('秩  序  册')
const coverFooter = ref('')
const coverBgFile = ref(null)
const coverInputRef = ref(null)

const filterKeyword = ref('')
const filterType = ref('all')

const watermarkDialogOpen = ref(false)
const watermarkUrl = ref('')
const watermarkTitle = ref('')

const filteredTemplates = computed(() => {
  const kw = filterKeyword.value.trim().toLowerCase()
  return templatesStore.templates.filter((t) => {
    if (filterType.value === 'builtin' && !t.builtin) return false
    if (filterType.value === 'custom' && t.builtin) return false
    if (kw) {
      return (
        String(t.name || '').toLowerCase().includes(kw) ||
        String(t.description || '').toLowerCase().includes(kw)
      )
    }
    return true
  })
})

function hasWatermark(tpl) {
  return Boolean(tpl.formatting?.watermark?.enabled && tpl.formatting?.watermark?.asset)
}

function visibleChapters(tpl) {
  return (tpl.chapters || []).filter((c) => c.toc_visible !== false)
}

function handleSampleFileChange(e) {
  const file = e.target.files?.[0]
  if (file) sampleFile.value = file
  e.target.value = ''
}

function handleDropSample(e) {
  draggingSample.value = false
  const file = e.dataTransfer?.files?.[0]
  if (file) sampleFile.value = file
}

async function handleParseSample() {
  if (!sampleFile.value) return
  try {
    const created = await templatesStore.parseSample({
      file: sampleFile.value,
      name: sampleName.value,
      useLlm: useLlm.value,
    })
    ElMessage.success(`模板“${created.name}”解析入库成功`)
    sampleFile.value = null
    sampleName.value = ''
  } catch (err) {
    ElMessage.error(err.message || '解析模板失败')
  }
}

function handleCoverBgChange(e) {
  const file = e.target.files?.[0]
  if (file) coverBgFile.value = file
  e.target.value = ''
}

async function handleGenerateCover() {
  if (!coverTitle.value.trim()) return
  try {
    await templatesStore.generateCover({
      eventTitle: coverTitle.value,
      subtitle: coverSubtitle.value,
      footer: coverFooter.value,
      file: coverBgFile.value,
    })
    ElMessage.success('封面合成成功')
  } catch (err) {
    ElMessage.error(err.message || '合成封面失败')
  }
}

function openWatermarkPreview(tpl) {
  watermarkTitle.value = tpl.name
  watermarkUrl.value = api.templateWatermarkUrl(tpl.id)
  watermarkDialogOpen.value = true
}

function useTemplateForProject(templateId) {
  ui.openCreateProjectDialog(templateId)
}

async function handleDeleteTemplate(id) {
  try {
    await templatesStore.deleteTemplate(id)
    ElMessage.success('自定义模板已删除')
  } catch (err) {
    ElMessage.error(err.message || '删除失败')
  }
}
</script>

<style scoped>
.tpl-workspace {
  display: grid;
  grid-template-columns: 350px minmax(0, 1fr);
  gap: 20px;
  align-items: start;
}

.tpl-studio {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.studio-card {
  padding: 18px 20px;
}

.studio-head {
  margin-bottom: 14px;
}

.mini-dropzone {
  width: 100%;
  padding: 14px 12px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-md);
  background: #ffffff;
  display: flex;
  align-items: center;
  gap: 12px;
  cursor: pointer;
  transition: all 150ms ease;
}

.mini-dropzone:hover,
.mini-dropzone.is-over {
  border-color: var(--accent);
  background: var(--accent-soft);
}

.mini-drop-copy {
  display: flex;
  flex-direction: column;
  min-width: 0;
  line-height: 1.35;
}

.mini-drop-copy strong {
  font-size: 13px;
  color: var(--text-primary);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.mini-drop-copy span {
  font-size: 11.5px;
  color: var(--text-muted);
}

.llm-switch-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-top: 2px;
}

.cover-file-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.cover-result {
  margin-top: 14px;
  padding-top: 14px;
  border-top: 1px solid var(--border-hairline);
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.cover-img {
  width: 100%;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border-hairline);
}

.dl-anchor {
  text-decoration: none;
}

.tpl-gallery-col {
  display: flex;
  flex-direction: column;
  gap: 16px;
  min-width: 0;
}

.gallery-toolbar {
  padding: 12px 16px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  flex-wrap: wrap;
}

.tpl-cards-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 16px;
}

.tpl-card {
  padding: 18px 20px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  transition: border-color 150ms ease, transform 150ms ease;
}

.tpl-card:hover {
  border-color: var(--accent-border);
}

.tpl-card-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 10px;
}

.tpl-name {
  margin: 0;
  font-size: 15.5px;
  font-weight: 650;
  color: var(--text-primary);
  line-height: 1.35;
}

.tpl-desc {
  margin: 0;
  font-size: 12.5px;
  color: var(--text-secondary);
  line-height: 1.55;
}

.tpl-spec-pills {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.wm-tag {
  cursor: pointer;
}

.tpl-chapter-flow {
  padding: 10px 12px;
  background: var(--bg-canvas);
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-sm);
}

.chapter-label {
  display: block;
  font-size: 11.5px;
  font-weight: 600;
  color: var(--text-secondary);
  margin-bottom: 6px;
}

.chapter-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 5px;
}

.chapter-chip {
  font-size: 11.5px;
  padding: 2px 8px;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: 999px;
  color: var(--text-secondary);
}

.tpl-card-footer {
  margin-top: auto;
  padding-top: 8px;
  border-top: 1px solid var(--border-hairline);
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.wm-preview-wrap {
  display: flex;
  justify-content: center;
  background: var(--bg-canvas);
  padding: 12px;
  border-radius: var(--radius-md);
}

.wm-preview-img {
  max-width: 100%;
  max-height: 68vh;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border-hairline);
}

.accent-icon {
  color: var(--accent);
}

.muted-icon {
  color: var(--text-muted);
}

@media (max-width: 980px) {
  .tpl-workspace {
    grid-template-columns: 1fr;
  }
}
</style>
