import { createRouter, createWebHistory } from 'vue-router'
import ProjectsView from '@/views/ProjectsView.vue'
import TemplatesView from '@/views/TemplatesView.vue'
import DrawUpdateView from '@/views/DrawUpdateView.vue'
import RosterView from '@/views/RosterView.vue'

const routes = [
  {
    path: '/',
    redirect: '/projects',
  },
  {
    path: '/projects/:id?',
    name: 'projects',
    component: ProjectsView,
    meta: { title: '整册秩序册生成' },
  },
  {
    path: '/templates',
    name: 'templates',
    component: TemplatesView,
    meta: { title: '模板库与封面工坊' },
  },
  {
    path: '/draw',
    name: 'draw',
    component: DrawUpdateView,
    meta: { title: '抽签后赛程更新' },
  },
  {
    path: '/roster',
    name: 'roster',
    component: RosterView,
    meta: { title: '单表名单整理' },
  },
  {
    path: '/:pathMatch(.*)*',
    redirect: '/projects',
  },
]

const router = createRouter({
  history: createWebHistory(),
  routes,
})

router.afterEach((to) => {
  const baseTitle = '体育赛事秩序册整册生成平台'
  document.title = to.meta?.title ? `${to.meta.title} · ${baseTitle}` : baseTitle
})

export default router
