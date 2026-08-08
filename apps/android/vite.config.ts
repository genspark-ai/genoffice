import path from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  plugins: [react()],
  resolve: {
    dedupe: ['react', 'react-dom'],
  },
  server: {
    fs: {
      allow: [path.resolve(__dirname, '../..')],
    },
  },
  build: {
    target: 'es2022',
    rollupOptions: {
      input: path.resolve(__dirname, 'index.html'),
    },
  },
})
