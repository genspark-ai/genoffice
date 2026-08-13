import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  root: 'src/renderer',
  plugins: [react()],
  server: {
    port: Number(process.env.MARKDOWN_DEV_PORT) || 5177,
    strictPort: true,
  },
})
