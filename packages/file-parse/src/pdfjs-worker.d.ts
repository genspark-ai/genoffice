/**
 * The pdfjs worker build artifact ships as plain ESM without type
 * declarations. We import it only for its top-level side effect — it
 * registers `globalThis.pdfjsWorker`, which lets pdfjs use its in-process
 * "fake worker" instead of resolving `pdf.worker.mjs` by path at runtime
 * (that lookup fails once the code is bundled).
 *
 * Declaring the module here (rather than a `@ts-expect-error` at the import
 * site) keeps the import type-safe no matter which tsconfig compiles this
 * package — a caller with a different `lib`/`types` set would otherwise see
 * the directive as unused.
 */
declare module 'pdfjs-dist/legacy/build/pdf.worker.mjs'
