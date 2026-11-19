/** electron-vite main-process asset imports: ?asset copies into the bundle and returns the runtime path */
declare module '*?asset' {
  const path: string
  export default path
}

/** electron-vite bundles the module as its own chunk and returns its runtime path (worker threads) */
declare module '*?modulePath' {
  const path: string
  export default path
}
