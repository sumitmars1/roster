<template>
  <header class="app-header">
    <div class="header-inner">
      <router-link to="/projects" class="brand-link">
        <div class="brand-logo">
          <Trophy :size="18" />
        </div>
        <div class="brand-copy">
          <span class="brand-name">赛事秩序册工作台</span>
          <span class="brand-sub">整册编排 · 智能抽取 · 抽签更新</span>
        </div>
      </router-link>

      <nav class="header-nav" aria-label="功能模块">
        <router-link
          v-for="item in navItems"
          :key="item.name"
          :to="item.to"
          class="nav-pill"
          :class="{ 'is-active': route.name === item.name }"
        >
          <component :is="item.icon" :size="16" />
          <span>{{ item.label }}</span>
        </router-link>
      </nav>

      <div class="header-actions">
        <el-button class="history-btn" @click="ui.openHistoryDrawer()">
          <History :size="15" />
          <span style="margin-left: 6px">名单历史</span>
        </el-button>
      </div>
    </div>
  </header>
</template>

<script setup>
import { useRoute } from 'vue-router'
import { BookOpen, History, LayoutTemplate, Shuffle, Trophy, Users } from 'lucide-vue-next'
import { useUiStore } from '@/stores/ui'

const route = useRoute()
const ui = useUiStore()

const navItems = [
  { name: 'projects', to: '/projects', label: '秩序册生成', icon: BookOpen },
  { name: 'templates', to: '/templates', label: '模板库与封面', icon: LayoutTemplate },
  { name: 'draw', to: '/draw', label: '抽签后更新', icon: Shuffle },
  { name: 'roster', to: '/roster', label: '单表名单', icon: Users },
]
</script>

<style scoped>
.app-header {
  position: sticky;
  top: 0;
  z-index: 50;
  background: rgba(255, 255, 255, 0.94);
  backdrop-filter: blur(12px);
  border-bottom: 1px solid var(--border-hairline);
}

.header-inner {
  max-width: 1360px;
  margin: 0 auto;
  padding: 0 24px;
  min-height: 64px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
}

.brand-link {
  display: inline-flex;
  align-items: center;
  gap: 11px;
  text-decoration: none;
  color: inherit;
  flex-shrink: 0;
}

.brand-logo {
  width: 36px;
  height: 36px;
  border-radius: 10px;
  background: var(--accent-soft);
  border: 1px solid var(--accent-border);
  color: var(--accent);
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.brand-copy {
  display: flex;
  flex-direction: column;
  line-height: 1.2;
}

.brand-name {
  font-size: 15px;
  font-weight: 700;
  color: var(--text-primary);
  letter-spacing: -0.01em;
}

.brand-sub {
  font-size: 11px;
  color: var(--text-muted);
  margin-top: 2px;
}

.header-nav {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px;
  background: var(--bg-canvas);
  border: 1px solid var(--border-hairline);
  border-radius: 999px;
}

.nav-pill {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 7px 14px;
  border-radius: 999px;
  font-size: 13.5px;
  font-weight: 500;
  color: var(--text-secondary);
  text-decoration: none;
  transition: all 150ms ease;
  white-space: nowrap;
}

.nav-pill:hover {
  color: var(--text-primary);
  background: rgba(255, 255, 255, 0.7);
}

.nav-pill.is-active {
  background: #ffffff;
  color: var(--accent);
  font-weight: 600;
  box-shadow: var(--shadow-xs);
  border: 1px solid var(--border-hairline);
}

.header-actions {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.history-btn {
  border-radius: 999px;
  border-color: var(--border-hairline);
  color: var(--text-secondary);
  background: #ffffff;
}

.history-btn:hover {
  color: var(--accent);
  border-color: var(--accent-border);
  background: var(--accent-soft);
}

@media (max-width: 960px) {
  .brand-sub {
    display: none;
  }
  .header-inner {
    padding: 10px 14px;
    flex-wrap: wrap;
  }
  .header-nav {
    order: 3;
    width: 100%;
    overflow-x: auto;
    justify-content: flex-start;
  }
}
</style>
