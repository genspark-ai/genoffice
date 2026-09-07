import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import { rhwpStudioPlugin } from './studio-vite-plugin'

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  preload: {
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  renderer: {
    plugins: [react(), rhwpStudioPlugin()],
    optimizeDeps: {
      exclude: ['@rhwp/editor'],
    },
    server: {
      port: Number(process.env.HWP_DEV_PORT) || 5178,
      strictPort: Boolean(process.env.HWP_DEV_PORT),
    },
  },
})
