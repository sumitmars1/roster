import { defineStore } from 'pinia'
import { ref } from 'vue'
import { api } from '@/api/client'

export const useDrawStore = defineStore('draw', () => {
  const job = ref(null)
  const submitting = ref(false)
  const error = ref('')

  let pollTimer = null

  function stopPolling() {
    if (pollTimer) {
      clearTimeout(pollTimer)
      pollTimer = null
    }
  }

  function schedulePoll(jobId) {
    stopPolling()
    pollTimer = setTimeout(async () => {
      try {
        const res = await api.getDrawJob(jobId)
        job.value = res
        if (res.status === 'running') {
          schedulePoll(jobId)
        } else {
          submitting.value = false
          if (res.status === 'error') {
            error.value = res.error || '抽签更新失败'
          }
        }
      } catch {
        schedulePoll(jobId)
      }
    }, 2000)
  }

  async function startJob({ bookFile, drawFiles, note = '' }) {
    stopPolling()
    submitting.value = true
    error.value = ''
    try {
      const fd = new FormData()
      fd.append('book_file', bookFile)
      for (const f of drawFiles) {
        fd.append('draw_files', f)
      }
      fd.append('note', note.trim())
      const created = await api.startDrawUpdate(fd)
      job.value = created
      schedulePoll(created.id)
      return created
    } catch (err) {
      submitting.value = false
      error.value = err.message || '提交抽签更新失败'
      throw err
    }
  }

  function resumePolling() {
    if (job.value?.id && job.value.status === 'running') {
      submitting.value = true
      schedulePoll(job.value.id)
    }
  }

  function resetJob() {
    stopPolling()
    job.value = null
    error.value = ''
    submitting.value = false
  }

  return {
    job,
    submitting,
    error,
    startJob,
    resetJob,
    resumePolling,
    stopPolling,
  }
})
