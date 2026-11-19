/** Minimal typings for bidi-js (ships no .d.ts). API per its README. */
declare module 'bidi-js' {
  export interface EmbeddingLevels {
    /** per-code-unit bidi embedding level (odd = RTL) */
    levels: Uint8Array
    paragraphs: Array<{ start: number; end: number; level: number }>
  }

  export interface Bidi {
    getEmbeddingLevels(text: string, explicitDirection?: 'ltr' | 'rtl'): EmbeddingLevels
    /** indices such that result[i] = source index of the unit displayed at position i */
    getReorderedIndices(
      text: string,
      embeddingLevels: EmbeddingLevels,
      start?: number,
      end?: number,
    ): number[]
    getReorderSegments(
      text: string,
      embeddingLevels: EmbeddingLevels,
      start?: number,
      end?: number,
    ): Array<[number, number]>
    getReorderedString(
      text: string,
      embeddingLevels: EmbeddingLevels,
      start?: number,
      end?: number,
    ): string
    getMirroredCharacter(char: string): string | null
    getMirroredCharactersMap(
      text: string,
      embeddingLevels: EmbeddingLevels,
      start?: number,
      end?: number,
    ): Map<number, string>
    getBidiCharType(char: string): number
    getBidiCharTypeName(char: string): string
  }

  export default function bidiFactory(): Bidi
}
