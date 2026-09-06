import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// renderer-only dev server (embedded by the shell via DOCS_RENDERER_URL for HMR; no standalone Electron)
export default defineConfig({
  root: 'src/renderer',
  plugins: [react()],
  server: {
    port: Number(process.env.DOCS_DEV_PORT) || 5173,
    strictPort: true,
    // web version: same-origin proxy to the HTTP IPC bridge inside the running
    // Electron main process (keeps the page CSP's connect-src 'self' intact)
    proxy: {
      '/api': {
        target: `http://127.0.0.1:${Number(process.env.DOCS_IPC_PORT) || 5273}`,
        changeOrigin: true,
      },
    },
  },
})
