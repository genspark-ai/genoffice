export interface SchemaProblem {
  part: string
  message: string
}
export function xmllintAvailable(): boolean
export function mcePreprocess(xml: string): string
export function validatePptx(input: string | Uint8Array): Promise<SchemaProblem[]>
export function newProblems(base: SchemaProblem[], edited: SchemaProblem[]): SchemaProblem[]
