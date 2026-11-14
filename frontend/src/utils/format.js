export function formatSize(bytes) {
  if (bytes === null || bytes === undefined) return ''
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

const dateFormatter = new Intl.DateTimeFormat('zh-CN', {
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
})

const fullDateFormatter = new Intl.DateTimeFormat('zh-CN', {
  dateStyle: 'medium',
  timeStyle: 'short',
})

export function formatShortTime(iso) {
  if (!iso) return ''
  try {
    return dateFormatter.format(new Date(iso))
  } catch {
    return String(iso)
  }
}

export function formatFullTime(iso) {
  if (!iso) return ''
  try {
    return fullDateFormatter.format(new Date(iso))
  } catch {
    return String(iso)
  }
}

export function getExtension(filename = '') {
  const parts = String(filename).split('.')
  return parts.length > 1 ? parts.pop().toLowerCase() : ''
}

export function getFileKind(filename = '') {
  const ext = getExtension(filename)
  if (['xlsx', 'xls', 'et', 'csv'].includes(ext)) return 'sheet'
  if (['docx', 'doc', 'wps'].includes(ext)) return 'word'
  if (ext === 'pdf') return 'pdf'
  if (['png', 'jpg', 'jpeg', 'webp', 'gif'].includes(ext)) return 'image'
  return 'file'
}

export const ROLE_HINT_OPTIONS = [
  { value: 'auto', label: '自动智能识别用途', short: '智能识别' },
  { value: 'cover', label: '作为封面', short: '封面' },
  { value: 'background', label: '作为封面底图', short: '封面底图' },
  { value: 'back_cover', label: '作为尾页', short: '尾页' },
  { value: 'watermark', label: '作为内页底图/水印', short: '内页底图' },
]

export function roleHintLabel(value) {
  const found = ROLE_HINT_OPTIONS.find((o) => o.value === value)
  return found ? found.short : value
}

export function projectStatusMeta(status) {
  switch (status) {
    case 'running':
      return { label: '生成中', type: 'primary', dot: 'running' }
    case 'ready':
      return { label: '已生成', type: 'success', dot: 'ready' }
    case 'error':
      return { label: '异常', type: 'danger', dot: 'error' }
    default:
      return { label: '待生成', type: 'info', dot: 'draft' }
  }
}

export function rosterStatusMeta(status) {
  const map = {
    queued: { label: '准备中', type: 'primary' },
    running: { label: '整理中', type: 'primary' },
    review: { label: '待核对', type: 'warning' },
    done: { label: '已完成', type: 'success' },
    error: { label: '未完成', type: 'danger' },
    stopped: { label: '已结束', type: 'info' },
  }
  return map[status] || { label: '准备中', type: 'info' }
}
