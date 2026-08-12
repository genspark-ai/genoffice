// Local packaging override: reuse the already-installed Electron runtime
// instead of re-downloading it (this machine's network can't reach the
// download host reliably). electron-builder merges this with
// electron-builder.cjs when passed via `-c`.
const { join } = require('node:path')

/** @type {import('electron-builder').Configuration} */
module.exports = {
  ...require('./electron-builder.cjs'),
  electronDist: join(__dirname, '../../node_modules/electron/dist'),
}
