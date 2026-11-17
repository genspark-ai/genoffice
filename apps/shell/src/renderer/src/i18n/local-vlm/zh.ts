export const zh = {
  note: '在本地通过 llama.cpp 运行 Liquid AI LFM3-VL-2B (Q4_K_M) 视觉语言模型，无需 API Key 或联网。',
  autoFallback: '未安装或未登录 Antigravity CLI 时自动回退到本地模型',
  autoFallbackHint: '当 agy 不可用时，聊天、图像理解与扫描 PDF OCR 自动使用本地 LFM3-VL-2B。',
  statusReady: 'LFM3-VL-2B (Q4_K_M) 与 llama-server 已就绪。',
  statusRunning: '本地推理服务运行中（端口 {port}）。',
  statusMissing: '尚未安装模型或推理引擎（{detail}）。',
  downloadBtn: '下载 LFM3-VL-2B (Q4_K_M)',
  downloading: '正在下载 {file}（{percent}%）',
  cancelDownload: '取消',
  startBtn: '启动服务',
  stopBtn: '停止服务',
  browse: '浏览',
  modelPathLabel: 'GGUF 模型路径',
  mmprojPathLabel: '视觉投影器 (mmproj) 路径',
  runtimePathLabel: 'llama-server 可执行文件',
  ramWarn: '当前可用内存较低（{mb} MB），LFM3-VL-2B Q4_K_M 建议预留约 2.2 GB 可用内存。',
  agyFallbackReady: '未检测到 Antigravity CLI —— 已启用本地 LFM3-VL-2B (Q4_K_M) 作为备选。',
  agyFallbackMissing:
    '未检测到 Antigravity CLI。可在“Local AI (LFM3-VL-2B)”中下载离线模型作为备选。',
  ocrBackendLabel: 'OCR 增强后端',
  ocrBackendAuto: '自动（Antigravity CLI，不可用时回退本地 LFM3-VL-2B）',
  ocrBackendAgy: '仅 Antigravity CLI',
  ocrBackendLocal: '本地 LFM3-VL-2B (Q4_K_M)',
  ocrBackendLocalActive: '当前使用本地 LFM3-VL-2B (Q4_K_M) 处理 OCR —— 不消耗 Antigravity 配额。',
} as const

export type LocalVlmKey = keyof typeof zh
