<template>
  <el-drawer
    v-model="ui.historyDrawerOpen"
    title="单表名单生成历史"
    size="460px"
    :append-to-body="true"
  >
    <div class="drawer-body">
      <div class="drawer-toolbar">
        <el-input
          v-model="keyword"
          placeholder="搜索赛事名称、组别或版式…"
          clearable
        >
          <template #prefix>
            <Search :size="15" class="muted-icon" />
          </template>
        </el-input>
        <el-button @click="roster.fetchHistory()" :loading="roster.loadingHistory" circle title="刷新历史">
          <RefreshCw :size="15" />
        </el-button>
      </div>

      <el-alert
        v-if="roster.historyError"
        type="error"
        :title="roster.historyError"
        show-icon
        :closable="false"
      />

      <div v-if="roster.loadingHistory && !roster.history.length" class="drawer-empty">
        <el-skeleton :rows="5" animated />
      </div>

      <el-empty
        v-else-if="!filteredHistory.length"
        description="暂无符合条件的名单记录"
        :image-size="72"
      />

      <div v-else class="history-list">
        <div
          v-for="entry in filteredHistory"
          :key="entry.id"
          class="history-card"
        >
          <div class="history-card-main" @click="openTask(entry.id)">
            <div class="history-title-row">
              <strong>{{ entry.title || '未命名赛事' }}</strong>
              <el-tag v-if="entry.preset_name" size="small" effect="plain" type="info">
                {{ entry.preset_name }}
              </el-tag>
            </div>
            <div class="history-meta">
              <span v-if="entry.group">{{ entry.group }}</span>
              <span v-if="entry.records">{{ entry.records }} 条记录</span>
              <span v-if="entry.generated_at">{{ formatFullTime(entry.generated_at) }}</span>
            </div>
          </div>
          <div class="history-actions">
            <el-button size="small" text type="primary" @click="openTask(entry.id)">
              <ExternalLink :size="14" />
              <span style="margin-left: 4px">查看</span>
            </el-button>
            <el-button
              v-if="entry.word_available"
              tag="a"
              :href="api.taskWordUrl(entry.id)"
              download
              size="small"
              plain
              class="dl-link"
            >
              <Download :size="14" />
              <span style="margin-left: 4px">Word</span>
            </el-button>
          </div>
        </div>
      </div>
    </div>
  </el-drawer>
</template>

<script setup>
import { computed, ref, watch } from 'vue'
import { useRouter } from 'vue-router'
import { Download, ExternalLink, RefreshCw, Search } from 'lucide-vue-next'
import { api } from '@/api/client'
import { useRosterStore } from '@/stores/roster'
import { useUiStore } from '@/stores/ui'
import { formatFullTime } from '@/utils/format'

const ui = useUiStore()
const roster = useRosterStore()
const router = useRouter()

const keyword = ref('')

watch(
  () => ui.historyDrawerOpen,
  (open) => {
    if (open) {
      roster.fetchHistory()
    }
  }
)

const filteredHistory = computed(() => {
  const kw = keyword.value.trim().toLowerCase()
  if (!kw) return roster.history
  return roster.history.filter((item) =>
    [item.title, item.group, item.preset_name].some((v) =>
      String(v || '').toLowerCase().includes(kw)
    )
  )
})

async function openTask(id) {
  ui.closeHistoryDrawer()
  await roster.restoreTask(id)
  router.push({ name: 'roster', query: { task: id } })
}
</script>

<style scoped>
.drawer-body {
  display: flex;
  flex-direction: column;
  gap: 14px;
}

.drawer-toolbar {
  display: flex;
  gap: 8px;
  align-items: center;
}

.muted-icon {
  color: var(--text-muted);
}

.history-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.history-card {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 12px 14px;
  background: #ffffff;
  border: 1px solid var(--border-hairline);
  border-radius: var(--radius-md);
  transition: border-color 150ms ease, box-shadow 150ms ease;
}

.history-card:hover {
  border-color: var(--accent-border);
  box-shadow: var(--shadow-xs);
}

.history-card-main {
  min-width: 0;
  flex: 1;
  cursor: pointer;
}

.history-title-row {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.history-title-row strong {
  font-size: 14px;
  color: var(--text-primary);
}

.history-meta {
  display: flex;
  flex-wrap: wrap;
  gap: 6px 10px;
  margin-top: 4px;
  font-size: 12px;
  color: var(--text-muted);
}

.history-actions {
  display: flex;
  align-items: center;
  gap: 4px;
  flex-shrink: 0;
}

.dl-link {
  text-decoration: none;
}
</style>
