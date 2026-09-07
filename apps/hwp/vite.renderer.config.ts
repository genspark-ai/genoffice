import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// renderer-only dev server (embedded by shell via HWP_RENDERER_URL for HMR)
export default defineConfig({
  root: 'src/renderer',
  plugins: [react()],
  server: {
    port: Number(process.env.HWP_DEV_PORT) || 5178,
    strictPort: true,
  },
})
