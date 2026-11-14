<template>
  <div class="app-shell">
    <AppHeader />
    <main class="app-main">
      <router-view />
    </main>
    <HistoryDrawer />
    <CreateProjectDialog />
  </div>
</template>

<script setup>
import { onMounted } from 'vue'
import AppHeader from '@/components/AppHeader.vue'
import CreateProjectDialog from '@/components/CreateProjectDialog.vue'
import HistoryDrawer from '@/components/HistoryDrawer.vue'
import { useTemplatesStore } from '@/stores/templates'

const templatesStore = useTemplatesStore()

onMounted(() => {
  templatesStore.fetchTemplates().catch(() => {})
})
</script>

<style scoped>
.app-shell {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  background: var(--bg-canvas);
}

.app-main {
  flex: 1;
}
</style>
