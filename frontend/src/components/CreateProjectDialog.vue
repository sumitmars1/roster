<template>
  <el-dialog
    v-model="ui.createProjectDialogOpen"
    title="新建秩序册项目"
    width="540px"
    destroy-on-close
    append-to-body
  >
    <el-form label-position="top" @submit.prevent="handleSubmit">
      <el-form-item label="项目 / 赛事名称" required>
        <el-input
          v-model="name"
          placeholder="如：綦江区第三届运动会五人制篮球比赛"
          maxlength="120"
          show-word-limit
          autofocus
        />
      </el-form-item>

      <el-form-item label="选择秩序册模板" required>
        <el-select v-model="templateId" style="width: 100%" placeholder="请选择模板">
          <el-option
            v-for="tpl in templatesStore.templates"
            :key="tpl.id"
            :label="`${tpl.name}${tpl.builtin ? '（内置）' : '（已入库）'}`"
            :value="tpl.id"
          />
        </el-select>
      </el-form-item>

      <div v-if="selectedTemplate" class="tpl-preview-box">
        <div class="tpl-preview-head">
          <strong>{{ selectedTemplate.name }}</strong>
          <el-tag size="small" :type="selectedTemplate.builtin ? 'info' : 'success'" effect="light">
            {{ selectedTemplate.builtin ? '内置模板' : '自定义入库' }}
          </el-tag>
        </div>
        <p class="tpl-preview-desc">{{ selectedTemplate.description }}</p>
        <div class="tpl-chapters">
          <span
            v-for="ch in visibleChapters"
            :key="ch.id"
            class="chapter-chip"
          >
            {{ ch.title }}
          </span>
        </div>
      </div>

      <el-alert
        v-if="error"
        type="error"
        :title="error"
        show-icon
        :closable="false"
        style="margin-top: 12px"
      />
    </el-form>

    <template #footer>
      <div class="dialog-footer">
        <el-button @click="ui.closeCreateProjectDialog()">取消</el-button>
        <el-button type="primary" :loading="submitting" @click="handleSubmit">
          <FolderPlus :size="15" />
          <span style="margin-left: 6px">创建项目</span>
        </el-button>
      </div>
    </template>
  </el-dialog>
</template>

<script setup>
import { computed, ref, watch } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { FolderPlus } from 'lucide-vue-next'
import { useProjectsStore } from '@/stores/projects'
import { useTemplatesStore } from '@/stores/templates'
import { useUiStore } from '@/stores/ui'

const ui = useUiStore()
const templatesStore = useTemplatesStore()
const projectsStore = useProjectsStore()
const router = useRouter()

const name = ref('')
const templateId = ref('')
const submitting = ref(false)
const error = ref('')

watch(
  () => ui.createProjectDialogOpen,
  (open) => {
    if (open) {
      error.value = ''
      name.value = ''
      if (ui.createProjectTemplateId) {
        templateId.value = ui.createProjectTemplateId
      } else if (!templateId.value && templatesStore.templates.length) {
        templateId.value = templatesStore.templates[0].id
      }
    }
  }
)

const selectedTemplate = computed(() =>
  templatesStore.templates.find((t) => t.id === templateId.value)
)

const visibleChapters = computed(() =>
  (selectedTemplate.value?.chapters || []).filter((c) => c.toc_visible !== false)
)

async function handleSubmit() {
  if (!name.value.trim()) {
    error.value = '请填写项目或赛事名称'
    return
  }
  if (!templateId.value) {
    error.value = '请选择秩序册模板'
    return
  }
  submitting.value = true
  error.value = ''
  try {
    const created = await projectsStore.createProject({
      name: name.value,
      templateId: templateId.value,
    })
    ui.closeCreateProjectDialog()
    ElMessage.success(`项目“${created.name}”已创建`)
    router.push({ name: 'projects', params: { id: created.id } })
  } catch (err) {
    error.value = err.message || '创建项目失败'
  } finally {
    submitting.value = false
  }
}
</script>

<style scoped>
.tpl-preview-box {
  background: var(--bg-canvas);
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  padding: 12px 14px;
}

.tpl-preview-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}

.tpl-preview-desc {
  margin: 6px 0 10px;
  font-size: 12.5px;
  color: var(--text-secondary);
  line-height: 1.5;
}

.tpl-chapters {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.chapter-chip {
  font-size: 11.5px;
  padding: 2px 8px;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: 999px;
  color: var(--text-secondary);
}

.dialog-footer {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
</style>
