/** Zoom gesture live-preview event (dispatched by App's rAF handler while the CSS
 * transform is animating, before any React commit). detail = the pending zoom.
 * Dependency-free so DOM-only modules (TextEditOverlay) can listen without pulling
 * in the Konva canvas stack. */
export const ZOOM_PREVIEW_EVENT = 'slides:zoom-preview'
