/**
 * Public API — the two entry points consumed by the rest of the application.
 *
 * The conversion pipeline for both formats follows the same high-level steps:
 * 1. Parse the file header to determine logical bounds and canvas dimensions.
 * 2. Create an in-memory canvas (OffscreenCanvas preferred, HTMLCanvasElement fallback).
 * 3. Replay every metafile record onto the canvas context in order.
 * 4. Resolve "deferred images" — bitmap / embedded-metafile draws that require
 *    async image decoding (via {@link createImageBitmap}).
 * 5. Export the canvas contents as a `data:image/png;base64,…` URL.
 *
 * @module emf-converter
 */
/**
 * Configuration options for EMF/WMF conversion.
 */
interface EmfConvertOptions {
    /** Maximum output width in pixels. */
    maxWidth?: number;
    /** Maximum output height in pixels. */
    maxHeight?: number;
    /**
     * DPI scale factor for higher-resolution output.
     * Default is 2 (HiDPI). Set to 1 for 1:1 pixel mapping.
     * Values above 4 are clamped to 4 to prevent excessive memory usage.
     */
    dpiScale?: number;
    /**
     * Hard cap on the output canvas width/height in pixels. Guards against
     * pathological metafiles. Defaults to {@link MAX_CANVAS_DIMENSION} (8192).
     */
    maxCanvasDimension?: number;
    /**
     * Maximum number of records processed per stream before replay stops.
     * Defaults to 200,000 for GDI/WMF and 500,000 for the finer-grained EMF+
     * stream. Supplying a value overrides both with the same cap.
     */
    maxRecords?: number;
    /**
     * Optional map from a (case-insensitive) Windows face name to a CSS font
     * family available in the rendering environment, e.g. `{ calibri: 'Carlito',
     * 'ms shell dlg': 'Tahoma' }`. Applied to GDI, WMF, and EMF+ text.
     */
    fontFamilyMap?: Record<string, string>;
}
/**
 * Converts an EMF (Enhanced Metafile) binary buffer to a PNG data-URL string
 * by parsing the EMF header, iterating over all EMR records, and replaying
 * them onto an in-memory canvas. Embedded EMF+ (GDI+) records found inside
 * EMR_COMMENT payloads are handled transparently.
 *
 * The canvas is rendered at a configurable DPI scale (default 2x) to produce
 * sharper output when displayed at CSS logical-pixel sizes. This is important
 * for presentations viewed on HiDPI/Retina displays.
 *
 * Returns `null` when:
 * - The buffer does not begin with a valid EMR_HEADER record.
 * - The logical bounds are zero-sized or negative.
 * - No canvas API is available (e.g. headless test environment).
 *
 * @param buffer  - The raw EMF file bytes.
 * @param options - Optional {@link EmfConvertOptions} controlling output size,
 *   DPI scale, record limits, and font mapping.
 * @returns A `data:image/png;base64,…` string, or `null` on failure.
 */
declare function convertEmfToDataUrl(buffer: ArrayBuffer, options?: EmfConvertOptions, recursionDepth?: number): Promise<string | null>;
/**
 * Converts a WMF (Windows Metafile) binary buffer to a PNG data-URL string.
 *
 * WMF is the older 16-bit metafile format with simpler record types than EMF.
 * This function parses the optional Aldus placeable header, the WMF header,
 * and then replays all META_* records onto a canvas.
 *
 * The canvas is rendered at a configurable DPI scale (default 2x) for
 * sharper output on HiDPI displays.
 *
 * Returns `null` when:
 * - The header cannot be parsed or reports invalid dimensions.
 * - No canvas API is available.
 *
 * @param buffer  - The raw WMF file bytes.
 * @param options - Optional {@link EmfConvertOptions} controlling output size,
 *   DPI scale, record limits, and font mapping.
 * @returns A `data:image/png;base64,…` string, or `null` on failure.
 */
declare function convertWmfToDataUrl(buffer: ArrayBuffer, options?: EmfConvertOptions, recursionDepth?: number): Promise<string | null>;

/**
 * Canvas creation, styling, string reading, stock objects, and export helpers.
 */

/**
 * The default DPI scale factor used for EMF/WMF rendering.
 * Set to 1 for 1:1 pixel mapping (default).
 * Set to 2 for HiDPI output (2x resolution).
 */
declare const DEFAULT_DPI_SCALE = 1;

export { DEFAULT_DPI_SCALE, type EmfConvertOptions, convertEmfToDataUrl, convertWmfToDataUrl };
