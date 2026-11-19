/** Vite-style raw text import (the op guides ship as markdown); bundlers outside Vite need a matching loader */
declare module '*.md?raw' {
  const text: string
  export default text
}
