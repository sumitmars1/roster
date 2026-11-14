export class ApiError extends Error {
  constructor(message, status) {
    super(message)
    this.name = 'ApiError'
    this.status = status
  }
}

export async function requestJson(url, options = {}) {
  const res = await fetch(url, options)
  const text = await res.text()
  let data = {}
  try {
    data = text ? JSON.parse(text) : {}
  } catch {
    data = { detail: text }
  }
  if (!res.ok) {
    const msg =
      typeof data.detail === 'string'
        ? data.detail
        : Array.isArray(data.detail)
        ? data.detail.map((d) => d.msg || JSON.stringify(d)).join('；')
        : `请求失败 (${res.status})`
    throw new ApiError(msg, res.status)
  }
  return data
}

export const api = {
  // Templates & Cover
  listTemplates: () => requestJson('/api/templates'),
  getTemplate: (id) => requestJson(`/api/templates/${encodeURIComponent(id)}`),
  parseTemplate: (formData) =>
    requestJson('/api/templates/parse', { method: 'POST', body: formData }),
  deleteTemplate: (id) =>
    requestJson(`/api/templates/${encodeURIComponent(id)}`, { method: 'DELETE' }),
  generateCoverPreview: (formData) =>
    requestJson('/api/cover/preview', { method: 'POST', body: formData }),
  templateWatermarkUrl: (id) => `/api/templates/${encodeURIComponent(id)}/watermark`,

  // Projects
  listProjects: () => requestJson('/api/projects'),
  createProject: (formData) =>
    requestJson('/api/projects', { method: 'POST', body: formData }),
  getProject: (id) => requestJson(`/api/projects/${encodeURIComponent(id)}`),
  updateProject: (id, formData) =>
    requestJson(`/api/projects/${encodeURIComponent(id)}`, { method: 'PATCH', body: formData }),
  uploadMaterials: (projectId, formData) =>
    requestJson(`/api/projects/${encodeURIComponent(projectId)}/materials`, {
      method: 'POST',
      body: formData,
    }),
  deleteMaterial: (projectId, materialId) =>
    requestJson(
      `/api/projects/${encodeURIComponent(projectId)}/materials/${encodeURIComponent(materialId)}`,
      { method: 'DELETE' }
    ),
  materialRawUrl: (projectId, materialId) =>
    `/api/projects/${encodeURIComponent(projectId)}/materials/${encodeURIComponent(materialId)}/raw`,
  generateBook: (projectId, formData) =>
    requestJson(`/api/projects/${encodeURIComponent(projectId)}/generate`, {
      method: 'POST',
      body: formData,
    }),
  projectDrawUpdate: (projectId, formData) =>
    requestJson(`/api/projects/${encodeURIComponent(projectId)}/draw-update`, {
      method: 'POST',
      body: formData,
    }),
  versionDocxUrl: (projectId, versionId = 'latest') =>
    `/api/projects/${encodeURIComponent(projectId)}/versions/${encodeURIComponent(versionId)}/download/docx`,
  versionPdfUrl: (projectId, versionId = 'latest') =>
    `/api/projects/${encodeURIComponent(projectId)}/versions/${encodeURIComponent(versionId)}/download/pdf`,
  versionPdfPreviewUrl: (projectId, versionId = 'latest') =>
    `/api/projects/${encodeURIComponent(projectId)}/versions/${encodeURIComponent(versionId)}/preview/pdf`,

  // Standalone Draw Update
  startDrawUpdate: (formData) =>
    requestJson('/api/draw-update', { method: 'POST', body: formData }),
  getDrawJob: (jobId) => requestJson(`/api/draw-jobs/${encodeURIComponent(jobId)}`),
  drawJobDocxUrl: (jobId) => `/api/draw-jobs/${encodeURIComponent(jobId)}/download/docx`,
  drawJobPdfUrl: (jobId) => `/api/draw-jobs/${encodeURIComponent(jobId)}/download/pdf`,
  drawJobPdfPreviewUrl: (jobId) => `/api/draw-jobs/${encodeURIComponent(jobId)}/preview/pdf`,

  // Single-Table Roster
  getOptions: () => requestJson('/api/options'),
  getHistory: () => requestJson('/api/history'),
  createTask: (formData) => requestJson('/api/tasks', { method: 'POST', body: formData }),
  getTask: (id) => requestJson(`/api/tasks/${encodeURIComponent(id)}`),
  retryTaskFile: (taskId, fileId) =>
    requestJson(`/api/tasks/${encodeURIComponent(taskId)}/files/${encodeURIComponent(fileId)}/retry`, {
      method: 'POST',
    }),
  restartTask: (taskId) =>
    requestJson(`/api/tasks/${encodeURIComponent(taskId)}/restart`, { method: 'POST' }),
  submitReview: (taskId, answer) =>
    requestJson(`/api/tasks/${encodeURIComponent(taskId)}/review`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ answer }),
    }),
  taskWordUrl: (taskId) => `/api/tasks/${encodeURIComponent(taskId)}/download/word`,
  taskPlacementUrl: (taskId) => `/api/tasks/${encodeURIComponent(taskId)}/download/placement`,
  taskSourceUrl: (taskId, fileId) =>
    `/api/tasks/${encodeURIComponent(taskId)}/source/${encodeURIComponent(fileId)}`,
}
