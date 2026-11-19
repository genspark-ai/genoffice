/** Reference viewport width the render tree is laid out against (px). */
export const FIT_WIDTH = 1280

/** Pixels per inch in the FIT_WIDTH viewport (for ruler/grid; 1280px ↔ 13.33in slide) */
export const PX_PER_INCH = 96

/** Arrow-key nudge step: PowerPoint's default grid, measured 7.2 pt = 0.1 in on Mac (Windows default 0.083 in). */
export const NUDGE_STEP_PX = PX_PER_INCH / 10
