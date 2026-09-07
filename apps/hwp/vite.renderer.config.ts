import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { rhwpStudioPlugin } from './studio-vite-plugin'

// renderer-only dev server (embedded by shell via HWP_RENDERER_URL for HMR)
export default defineConfig({
  root: 'src/renderer',
  plugins: [react(), rhwpStudioPlugin()],
  optimizeDeps: {
    exclude: ['@rhwp/editor'],
  },
  server: {
    port: Number(process.env.HWP_DEV_PORT) || 5178,
    strictPort: true,
  },
})
