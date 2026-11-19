/** electron-vite main-process asset imports: ?asset copies into the bundle and returns the runtime path */
declare module '*?asset' {
  const path: string
  export default path
}
