import { defineStore } from 'pinia'
import { ref } from 'vue'
import { api } from '@/api/client'

export const useTemplatesStore = defineStore('templates', () => {
  const templates = ref([])
  const loading = ref(false)
  const error = ref('')

  const parsing = ref(false)
  const parseMessage = ref('')
  const parseError = ref('')

  const coverGenerating = ref(false)
  const coverPreviewUrl = ref('')
  const coverError = ref('')

  async function fetchTemplates() {
    loading.value = true
    error.value = ''
    try {
      const data = await api.listTemplates()
      templates.value = data.templates || []
    } catch (err) {
      error.value = err.message || '加载模板失败'
    } finally {
      loading.value = false
    }
  }

  async function parseSample({ file, name = '', useLlm = true }) {
    parsing.value = true
    parseError.value = ''
    parseMessage.value = '正在解析样例章节结构、版式字体、名单风格及固定底图…'
    try {
      const fd = new FormData()
      fd.append('file', file)
      fd.append('name', name.trim())
      fd.append('use_llm', useLlm ? 'true' : 'false')
      const created = await api.parseTemplate(fd)
      parseMessage.value = `解析入库成功：${created.name}`
      await fetchTemplates()
      return created
    } catch (err) {
      parseError.value = err.message || '解析模板失败'
      parseMessage.value = ''
      throw err
    } finally {
      parsing.value = false
    }
  }

  async function deleteTemplate(id) {
    await api.deleteTemplate(id)
    await fetchTemplates()
  }

  async function generateCover({ eventTitle, subtitle = '秩  序  册', footer = '', file = null }) {
    coverGenerating.value = true
    coverError.value = ''
    try {
      const fd = new FormData()
      fd.append('event_title', eventTitle.trim())
      fd.append('subtitle', subtitle.trim() || '秩  序  册')
      fd.append('footer', footer)
      if (file) {
        fd.append('file', file)
      }
      const data = await api.generateCoverPreview(fd)
      coverPreviewUrl.value = `${data.preview_url}?t=${Date.now()}`
      return data
    } catch (err) {
      coverError.value = err.message || '生成封面失败'
      throw err
    } finally {
      coverGenerating.value = false
    }
  }

  return {
    templates,
    loading,
    error,
    parsing,
    parseMessage,
    parseError,
    coverGenerating,
    coverPreviewUrl,
    coverError,
    fetchTemplates,
    parseSample,
    deleteTemplate,
    generateCover,
  }
})
