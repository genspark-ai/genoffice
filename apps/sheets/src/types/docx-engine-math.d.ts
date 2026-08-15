/**
 * Type shim: @genoffice/docx-engine ships as TS source and fails under
 * sheets' stricter compiler options (same recipe as file-parse.d.ts).
 * tsconfig paths point type resolution here; runtime bundling still uses the
 * real source. Keep in sync with packages/docx-engine/src/math.ts.
 */

/** LaTeX (subset) → OMML fragment; throws on unsupported commands */
export function latexToOmml(latex: string): string

/** OMML (`<m:oMath>…</m:oMath>`) → MathML markup for native rendering */
export function ommlToMathML(ommlXml: string): string
