<template>
  <div class="pdf-preview-box">
    <div class="pdf-toolbar">
      <div class="pdf-toolbar-left">
        <FileText :size="16" class="pdf-icon" />
        <span class="pdf-label">{{ title }}</span>
      </div>
      <div class="pdf-toolbar-actions">
        <el-button v-if="docxUrl" tag="a" :href="docxUrl" download size="small" plain class="btn-anchor">
          <Download :size="14" />
          <span style="margin-left: 5px">下载 Word (.docx)</span>
        </el-button>
        <el-button v-if="pdfUrl" tag="a" :href="pdfUrl" download size="small" type="primary" class="btn-anchor">
          <Download :size="14" />
          <span style="margin-left: 5px">下载 PDF (.pdf)</span>
        </el-button>
        <el-button size="small" plain @click="fullscreen = true" title="全屏预览">
          <Maximize2 :size="14" />
        </el-button>
      </div>
    </div>

    <div class="pdf-frame-wrap">
      <iframe :src="`${previewUrl}#view=FitH`" :title="title" class="pdf-iframe" />
    </div>

    <el-dialog
      v-model="fullscreen"
      :title="title"
      width="92vw"
      top="3vh"
      destroy-on-close
      append-to-body
    >
      <iframe :src="`${previewUrl}#view=FitH`" :title="title" class="pdf-iframe-full" />
    </el-dialog>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { Download, FileText, Maximize2 } from 'lucide-vue-next'

defineProps({
  title: { type: String, default: 'PDF 在线预览' },
  previewUrl: { type: String, required: true },
  docxUrl: { type: String, default: '' },
  pdfUrl: { type: String, default: '' },
})

const fullscreen = ref(false)
</script>

<style scoped>
.pdf-preview-box {
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  background: #ffffff;
  overflow: hidden;
}

.pdf-toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 14px;
  background: #ffffff;
  border-bottom: 1px solid var(--border-hairline);
  flex-wrap: wrap;
}

.pdf-toolbar-left {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.pdf-icon {
  color: var(--danger);
}

.pdf-label {
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text-primary);
}

.pdf-toolbar-actions {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.btn-anchor {
  text-decoration: none;
}

.pdf-frame-wrap {
  background: #f1f5f9;
}

.pdf-iframe {
  display: block;
  width: 100%;
  height: 760px;
  border: 0;
}

.pdf-iframe-full {
  display: block;
  width: 100%;
  height: 82vh;
  border: 0;
  border-radius: 8px;
}
</style>
