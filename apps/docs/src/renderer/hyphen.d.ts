declare module 'hyphen' {
  const createHyphenator: (
    patterns: unknown,
    options?: { minWordLength?: number; hyphenChar?: string },
  ) => (text: string) => string
  export default createHyphenator
}

declare module 'hyphen/patterns/*' {
  const patterns: unknown
  export default patterns
}
