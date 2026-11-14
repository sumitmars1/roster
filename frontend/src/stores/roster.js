import { defineStore } from 'pinia'
import { computed, ref } from 'vue'
import { api } from '@/api/client'

const STORAGE_KEY = 'roster-harness-task'

export const useRosterStore = defineStore('roster', () => {
  const options = ref(null)
  const loadingOptions = ref(false)
  const optionsError = ref('')

  const history = ref([])
  const loadingHistory = ref(false)
  const historyError = ref('')

  const selectedFiles = ref([])
  const title = ref('')
  const group = ref('')
  const activePreset = ref('')
  const uploadNote = ref('')

  const task = ref(null)
  const submitting = ref(false)
  const submitError = ref('')
  const reviewSubmitting = ref(false)
  const reviewError = ref('')

  let pollTimer = null
  let uidCounter = 1

  const selectedCount = computed(() => selectedFiles.value.filter((f) => f.selected).length)
  const canStart = computed(
    () =>
      !submitting.value &&
      title.value.trim().length > 0 &&
      selectedCount.value > 0 &&
      Boolean(activePreset.value)
  )

  function stopPolling() {
    if (pollTimer) {
      clearTimeout(pollTimer)
      pollTimer = null
    }
  }

  function rememberTask(id) {
    try {
      if (id) localStorage.setItem(STORAGE_KEY, id)
      else localStorage.removeItem(STORAGE_KEY)
    } catch {
      // ignore storage errors
    }
  }

  function getSavedTaskId() {
    try {
      const fromUrl = new URLSearchParams(window.location.search).get('task')
      return fromUrl || localStorage.getItem(STORAGE_KEY)
    } catch {
      return null
    }
  }

  async function fetchOptions() {
    if (options.value) return options.value
    loadingOptions.value = true
    optionsError.value = ''
    try {
      const data = await api.getOptions()
      options.value = data
      if (!activePreset.value && data.presets?.length) {
        activePreset.value = data.presets[0].id
      }
      if (!data.key_configured) {
        submitError.value = '服务端未配置 DASHSCOPE_API_KEY，暂时不能开始整理。'
      }
      return data
    } catch (err) {
      optionsError.value = err.message || '本地服务连接失败，请刷新重试。'
      throw err
    } finally {
      loadingOptions.value = false
    }
  }

  async function fetchHistory() {
    loadingHistory.value = true
    historyError.value = ''
    try {
      history.value = await api.getHistory()
    } catch (err) {
      historyError.value = err.message || '读取生成历史失败'
    } finally {
      loadingHistory.value = false
    }
  }

  function addFiles(fileList) {
    if (!options.value?.upload) return
    const limits = options.value.upload
    const rejected = []
    let total = selectedFiles.value.reduce((sum, entry) => sum + entry.file.size, 0)

    for (const file of fileList) {
      const ext = file.name.split('.').pop()?.toLowerCase() || ''
      if (!limits.extensions.includes('.' + ext)) {
        rejected.push(`${file.name} 的格式不受支持`)
        continue
      }
      if (file.size > limits.max_file_bytes) {
        rejected.push(`${file.name} 超过单份 ${Math.round(limits.max_file_bytes / 1024 / 1024)} MB`)
        continue
      }
      if (total + file.size > limits.max_total_bytes) {
        rejected.push(`${file.name} 使整批超过 3 GB`)
        continue
      }
      const duplicate = selectedFiles.value.some(
        (entry) =>
          entry.file.name === file.name &&
          entry.file.size === file.size &&
          entry.file.lastModified === file.lastModified
      )
      if (duplicate) continue
      if (selectedFiles.value.length >= limits.max_files) {
        rejected.push(`一次最多 ${limits.max_files} 份`)
        break
      }
      selectedFiles.value.push({
        uid: `rf-${uidCounter++}`,
        file,
        selected: true,
      })
      total += file.size
    }
    uploadNote.value = rejected.join('；')
  }

  function removeFile(uid) {
    selectedFiles.value = selectedFiles.value.filter((f) => f.uid !== uid)
  }

  function clearFiles() {
    selectedFiles.value = []
    uploadNote.value = ''
  }

  function schedulePoll(id) {
    stopPolling()
    pollTimer = setTimeout(async () => {
      try {
        const value = await api.getTask(id)
        task.value = value
        if (['queued', 'running'].includes(value.status)) {
          schedulePoll(id)
        } else {
          submitting.value = false
          if (value.status === 'done') {
            fetchHistory().catch(() => {})
          }
        }
      } catch (err) {
        submitting.value = false
        submitError.value = err.message || '获取任务状态失败'
      }
    }, 1200)
  }

  function resumePolling() {
    if (task.value?.id && ['queued', 'running'].includes(task.value.status)) {
      schedulePoll(task.value.id)
    }
  }

  async function startTask() {
    if (!canStart.value) return
    stopPolling()
    submitting.value = true
    submitError.value = ''
    try {
      const fd = new FormData()
      fd.append('title', title.value.trim())
      fd.append('group', group.value.trim())
      fd.append('preset', activePreset.value)
      for (const entry of selectedFiles.value) {
        if (entry.selected) {
          fd.append('files', entry.file, entry.file.name)
        }
      }
      const value = await api.createTask(fd)
      task.value = value
      rememberTask(value.id)
      if (['queued', 'running'].includes(value.status)) {
        schedulePoll(value.id)
      } else {
        submitting.value = false
      }
      return value
    } catch (err) {
      submitting.value = false
      submitError.value = err.message || '提交整理任务失败'
      throw err
    }
  }

  async function restoreTask(id) {
    if (!id) return
    try {
      const value = await api.getTask(id)
      task.value = value
      rememberTask(value.id)
      title.value = value.title || ''
      group.value = value.group || ''
      if (value.preset) activePreset.value = value.preset
      if (['queued', 'running'].includes(value.status)) {
        schedulePoll(value.id)
      } else {
        submitting.value = false
      }
      return value
    } catch {
      rememberTask(null)
    }
  }

  async function submitReview(answer) {
    if (!task.value || !task.value.review || reviewSubmitting.value) return
    const text = String(answer || '').trim()
    if (!text) return
    reviewSubmitting.value = true
    reviewError.value = ''
    try {
      const updated = await api.submitReview(task.value.id, text)
      task.value = updated
      schedulePoll(updated.id)
      return updated
    } catch (err) {
      reviewError.value = err.message || '提交核对答复失败'
      throw err
    } finally {
      reviewSubmitting.value = false
    }
  }

  async function retryFile(fileId) {
    if (!task.value) return
    const updated = await api.retryTaskFile(task.value.id, fileId)
    task.value = updated
    schedulePoll(updated.id)
    return updated
  }

  async function restartTask() {
    if (!task.value || task.value.status !== 'error') return
    const updated = await api.restartTask(task.value.id)
    task.value = updated
    rememberTask(updated.id)
    schedulePoll(updated.id)
    return updated
  }

  return {
    options,
    loadingOptions,
    optionsError,
    history,
    loadingHistory,
    historyError,
    selectedFiles,
    title,
    group,
    activePreset,
    uploadNote,
    task,
    submitting,
    submitError,
    reviewSubmitting,
    reviewError,
    selectedCount,
    canStart,
    fetchOptions,
    fetchHistory,
    addFiles,
    removeFile,
    clearFiles,
    startTask,
    restoreTask,
    getSavedTaskId,
    submitReview,
    retryFile,
    restartTask,
    resumePolling,
    stopPolling,
  }
})
