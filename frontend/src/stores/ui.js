import { defineStore } from 'pinia'
import { ref } from 'vue'

export const useUiStore = defineStore('ui', () => {
  const historyDrawerOpen = ref(false)
  const createProjectDialogOpen = ref(false)
  const createProjectTemplateId = ref('')

  function openHistoryDrawer() {
    historyDrawerOpen.value = true
  }

  function closeHistoryDrawer() {
    historyDrawerOpen.value = false
  }

  function openCreateProjectDialog(templateId = '') {
    createProjectTemplateId.value = templateId || ''
    createProjectDialogOpen.value = true
  }

  function closeCreateProjectDialog() {
    createProjectDialogOpen.value = false
  }

  return {
    historyDrawerOpen,
    createProjectDialogOpen,
    createProjectTemplateId,
    openHistoryDrawer,
    closeHistoryDrawer,
    openCreateProjectDialog,
    closeCreateProjectDialog,
  }
})
