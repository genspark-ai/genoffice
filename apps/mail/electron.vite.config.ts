import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

export default defineConfig({
  // @genoffice/i18n, @genoffice/electron-utils, @genoffice/mail-protocols
  // and @genoffice/file-parse ship as TS source — must be bundled (same
  // policy as apps/markdown). file-parse is the H5 quarantine preview
  // extractor (Q-3); no second parser exists in the mail app.
  main: {
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/i18n',
          '@genoffice/electron-utils',
          '@genoffice/mail-protocols',
          '@genoffice/file-parse',
        ],
      }),
    ],
  },
  preload: {
    // same bundling requirement as main (see comment above); the preload
    // stays zod-free — it imports channel names only, never the schemas.
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@genoffice/i18n', '@genoffice/electron-utils', '@genoffice/mail-protocols'],
      }),
    ],
  },
  renderer: {
    plugins: [react()],
    server: {
      port: Number(process.env.MAIL_DEV_PORT) || 5179,
      strictPort: true,
    },
  },
})
