/**
 * Konva-drawn editing chrome colours. The redact frame rides into thumbnails
 * over the picture it marks, so it is drawn the same in both UI themes.
 */
export const CANVAS_COLORS = {
  redactFrame: {
    stroke: 'rgba(0,0,0,0.55)',
    chip: 'rgba(0,0,0,0.55)',
    label: '#ffffff',
  },
} as const
