import type { zh } from './zh'

export const copy = {
  note: 'Runs Liquid AI LFM3-VL-2B (Q4_K_M) locally via llama.cpp — no API key or internet required once downloaded.',
  autoFallback: 'Auto-fallback when Antigravity CLI is unavailable',
  autoFallbackHint:
    'Automatically uses local LFM3-VL-2B for chat, image analysis, and scanned-PDF OCR when agy is missing or signed out.',
  statusReady: 'LFM3-VL-2B (Q4_K_M) and llama-server are ready on this device.',
  statusRunning: 'Local inference server running on port {port}.',
  statusMissing: 'Model or runtime not installed yet ({detail}).',
  downloadBtn: 'Download LFM3-VL-2B (Q4_K_M)',
  downloading: 'Downloading {file} ({percent}%)',
  cancelDownload: 'Cancel',
  startBtn: 'Start server',
  stopBtn: 'Stop server',
  browse: 'Browse',
  modelPathLabel: 'GGUF model path',
  mmprojPathLabel: 'Vision projector (mmproj) path',
  runtimePathLabel: 'llama-server binary',
  ramWarn:
    'Low free memory ({mb} MB available). LFM3-VL-2B Q4_K_M works best with ~2.2 GB free RAM.',
  agyFallbackReady: 'Antigravity CLI unavailable — local LFM3-VL-2B (Q4_K_M) fallback is ready.',
  agyFallbackMissing:
    'Antigravity CLI unavailable. Download LFM3-VL-2B (Q4_K_M) under Local AI (LFM3-VL-2B) for offline fallback.',
  ocrBackendLabel: 'OCR escalation backend',
  ocrBackendAuto: 'Automatic (Antigravity CLI, local LFM3-VL-2B fallback)',
  ocrBackendAgy: 'Antigravity CLI only',
  ocrBackendLocal: 'Local LFM3-VL-2B (Q4_K_M)',
  ocrBackendLocalActive: 'Using local LFM3-VL-2B (Q4_K_M) on-device — no Antigravity quota used.',
} satisfies Record<keyof typeof zh, string>
