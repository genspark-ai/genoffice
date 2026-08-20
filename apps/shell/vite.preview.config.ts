import react from '@vitejs/plugin-react'
import { defineConfig, type Plugin } from 'vite'

// Renderer-only preview server for the KĀRYA shell Home screen (browser
// preview, no Electron — see preview-shim.ts for the bridge stub). Mirrors
// apps/docs/vite.renderer.config.ts; never used by production builds.
const bridgeShim: Plugin = {
  name: 'karya-preview-bridge-shim',
  transformIndexHtml() {
    return [
      {
        tag: 'script',
        attrs: { type: 'module', src: '/preview-shim.ts' },
        injectTo: 'head-prepend',
      },
    ]
  },
}

export default defineConfig({
  root: 'src/renderer',
  plugins: [react(), bridgeShim],
  server: {
    port: Number(process.env.SHELL_DEV_PORT) || 5199,
    strictPort: true,
    // bind the IPv4 loopback explicitly: Vite defaults to localhost (::1 on
    // this machine), but the preview host probes 127.0.0.1
    host: '127.0.0.1',
  },
})
