import { defineStore } from 'pinia'
import { computed, ref } from 'vue'
import { api } from '@/api/client'

export const useProjectsStore = defineStore('projects', () => {
  const projects = ref([])
  const activeProjectId = ref(null)
  const activeProject = ref(null)
  const selectedVersionId = ref(null)

  const loadingList = ref(false)
  const loadingDetail = ref(false)
  const uploadingMaterials = ref(false)
  const actionBusy = ref(false)
  const listError = ref('')

  let pollTimer = null

  const activeVersions = computed(() => {
    const list = activeProject.value?.versions || []
    return list.slice().reverse()
  })

  const currentVersion = computed(() => {
    const list = activeProject.value?.versions || []
    if (!list.length) return null
    if (selectedVersionId.value) {
      const found = list.find((v) => v.id === selectedVersionId.value)
      if (found) return found
    }
    const latestId = activeProject.value?.latest_version_id
    return list.find((v) => v.id === latestId) || list[list.length - 1]
  })

  function stopPolling() {
    if (pollTimer) {
      clearTimeout(pollTimer)
      pollTimer = null
    }
  }

  function schedulePolling(projectId) {
    stopPolling()
    pollTimer = setTimeout(async () => {
      if (activeProjectId.value !== projectId) return
      try {
        const detail = await api.getProject(projectId)
        applyProjectDetail(detail)
        if (detail.status === 'running') {
          schedulePolling(projectId)
        } else {
          await fetchProjects({ autoSelect: false })
        }
      } catch {
        schedulePolling(projectId)
      }
    }, 2400)
  }

  function applyProjectDetail(detail) {
    const sameProject = activeProject.value?.id === detail?.id
    const prevLatest = sameProject ? activeProject.value?.latest_version_id : null
    activeProject.value = detail
    const versions = detail?.versions || []
    if (!versions.length) {
      selectedVersionId.value = null
    } else if (
      !sameProject ||
      (detail?.latest_version_id && detail.latest_version_id !== prevLatest) ||
      !selectedVersionId.value ||
      !versions.some((v) => v.id === selectedVersionId.value)
    ) {
      selectedVersionId.value = detail.latest_version_id || versions[versions.length - 1].id
    }
  }

  function resumePolling() {
    if (activeProject.value?.id && activeProject.value.status === 'running') {
      schedulePolling(activeProject.value.id)
    }
  }

  async function fetchProjects({ autoSelect = true, preferredId = null } = {}) {
    loadingList.value = true
    listError.value = ''
    try {
      const data = await api.listProjects()
      projects.value = data.projects || []
      if (!autoSelect) return
      const targetId =
        (preferredId && projects.value.some((p) => p.id === preferredId) && preferredId) ||
        (activeProjectId.value && projects.value.some((p) => p.id === activeProjectId.value) && activeProjectId.value) ||
        projects.value[0]?.id ||
        null
      if (targetId) {
        await selectProject(targetId)
      } else {
        activeProjectId.value = null
        activeProject.value = null
        selectedVersionId.value = null
        stopPolling()
      }
    } catch (err) {
      listError.value = err.message || '加载项目列表失败'
    } finally {
      loadingList.value = false
    }
  }

  async function selectProject(projectId) {
    if (!projectId) return
    activeProjectId.value = projectId
    loadingDetail.value = true
    try {
      const detail = await api.getProject(projectId)
      applyProjectDetail(detail)
      if (detail.status === 'running') {
        schedulePolling(projectId)
      } else {
        stopPolling()
      }
      return detail
    } finally {
      loadingDetail.value = false
    }
  }

  async function createProject({ name, templateId }) {
    const fd = new FormData()
    fd.append('name', name.trim())
    fd.append('template_id', templateId)
    const created = await api.createProject(fd)
    await fetchProjects({ autoSelect: false })
    await selectProject(created.id)
    return created
  }

  async function updateProject(projectId, { name, templateId }) {
    const fd = new FormData()
    if (name !== undefined) fd.append('name', name.trim())
    if (templateId !== undefined) fd.append('template_id', templateId)
    await api.updateProject(projectId, fd)
    await fetchProjects({ autoSelect: false })
    await selectProject(projectId)
  }

  async function uploadMaterials(projectId, files, roleHint = 'auto') {
    if (!files?.length) return
    uploadingMaterials.value = true
    try {
      const fd = new FormData()
      for (const f of files) {
        fd.append('files', f)
      }
      fd.append('role_hint', roleHint || 'auto')
      await api.uploadMaterials(projectId, fd)
      await selectProject(projectId)
      await fetchProjects({ autoSelect: false })
    } finally {
      uploadingMaterials.value = false
    }
  }

  async function replaceMaterial(projectId, materialId, file) {
    uploadingMaterials.value = true
    try {
      const fd = new FormData()
      fd.append('files', file)
      fd.append('replace_material_id', materialId)
      await api.uploadMaterials(projectId, fd)
      await selectProject(projectId)
      await fetchProjects({ autoSelect: false })
    } finally {
      uploadingMaterials.value = false
    }
  }

  async function deleteMaterial(projectId, materialId) {
    await api.deleteMaterial(projectId, materialId)
    await selectProject(projectId)
    await fetchProjects({ autoSelect: false })
  }

  async function generateBook(projectId, note = '') {
    actionBusy.value = true
    try {
      const fd = new FormData()
      fd.append('note', note.trim())
      const updated = await api.generateBook(projectId, fd)
      applyProjectDetail({ ...activeProject.value, ...updated })
      schedulePolling(projectId)
      await fetchProjects({ autoSelect: false })
      return updated
    } finally {
      actionBusy.value = false
    }
  }

  async function submitDrawUpdate(projectId, { files, sourceVersionId, note = '' }) {
    actionBusy.value = true
    try {
      const fd = new FormData()
      for (const f of files) {
        fd.append('files', f)
      }
      if (sourceVersionId) fd.append('source_version_id', sourceVersionId)
      fd.append('note', note.trim())
      const updated = await api.projectDrawUpdate(projectId, fd)
      applyProjectDetail({ ...activeProject.value, ...updated })
      schedulePolling(projectId)
      await fetchProjects({ autoSelect: false })
      return updated
    } finally {
      actionBusy.value = false
    }
  }

  function setSelectedVersion(versionId) {
    selectedVersionId.value = versionId
  }

  return {
    projects,
    activeProjectId,
    activeProject,
    selectedVersionId,
    activeVersions,
    currentVersion,
    loadingList,
    loadingDetail,
    uploadingMaterials,
    actionBusy,
    listError,
    fetchProjects,
    selectProject,
    createProject,
    updateProject,
    uploadMaterials,
    replaceMaterial,
    deleteMaterial,
    generateBook,
    submitDrawUpdate,
    setSelectedVersion,
    resumePolling,
    stopPolling,
  }
})
