import type { zh } from './zh'

export const copy = {
  note: 'Chạy model Liquid AI LFM3-VL-2B (Q4_K_M) trực tiếp trên máy qua llama.cpp — không cần API key hay mạng sau khi tải.',
  autoFallback: 'Tự động dùng model cục bộ khi không có Antigravity CLI',
  autoFallbackHint:
    'Khi agy chưa cài hoặc chưa đăng nhập, Chat AI, phân tích ảnh và OCR PDF scan tự động chuyển sang LFM3-VL-2B trên máy.',
  statusReady: 'LFM3-VL-2B (Q4_K_M) và llama-server đã sẵn sàng trên máy.',
  statusRunning: 'Server cục bộ đang chạy tại cổng {port}.',
  statusMissing: 'Chưa cài đủ model hoặc runtime ({detail}).',
  downloadBtn: 'Tải LFM3-VL-2B (Q4_K_M)',
  downloading: 'Đang tải {file} ({percent}%)',
  cancelDownload: 'Hủy',
  startBtn: 'Khởi động server',
  stopBtn: 'Dừng server',
  browse: 'Chọn file',
  modelPathLabel: 'Đường dẫn model GGUF',
  mmprojPathLabel: 'Đường dẫn vision projector (mmproj)',
  runtimePathLabel: 'Tệp thực thi llama-server',
  ramWarn:
    'RAM trống thấp ({mb} MB). LFM3-VL-2B Q4_K_M hoạt động tốt nhất với khoảng 2.2 GB RAM trống.',
  agyFallbackReady:
    'Không có Antigravity CLI — đã sẵn sàng fallback sang LFM3-VL-2B (Q4_K_M) cục bộ.',
  agyFallbackMissing:
    'Không có Antigravity CLI. Bạn có thể tải LFM3-VL-2B (Q4_K_M) ở mục Local AI để dùng dự phòng offline.',
  ocrBackendLabel: 'Backend OCR nâng cao',
  ocrBackendAuto: 'Tự động (Antigravity CLI, dự phòng LFM3-VL-2B cục bộ)',
  ocrBackendAgy: 'Chỉ Antigravity CLI',
  ocrBackendLocal: 'LFM3-VL-2B (Q4_K_M) cục bộ',
  ocrBackendLocalActive:
    'Đang dùng LFM3-VL-2B (Q4_K_M) trực tiếp trên máy — không tốn hạn mức Antigravity.',
} satisfies Record<keyof typeof zh, string>
