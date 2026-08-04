/**
 * electron-builder configuration (moved out of package.json "build" so the
 * auto-update feed URL can be injected at build time instead of living in
 * the repo).
 *
 * GENOFFICE_UPDATE_URL — public base URL of the update channel (the generic
 * provider prefix that serves latest.yml / latest-mac.yml). Required for
 * release builds; CI provides it as a repository secret. For local release
 * builds put it in apps/shell/electron-builder.env (gitignored) — the
 * electron-builder CLI loads that file automatically.
 *
 * When the variable is unset (forks, PR smoke builds, plain local packaging)
 * the publish config is omitted: electron-builder then bakes no
 * app-update.yml into the app and in-app auto-update stays disabled.
 */

const { existsSync } = require('node:fs')
const { join } = require('node:path')

const updateUrl = process.env.GENOFFICE_UPDATE_URL

// The gsk CLI tree below is copied verbatim from node_modules. Its `commander`
// dependency may be installed either nested under @genspark/cli or hoisted to
// the top-level node_modules (npm's layout differs by version/installer), so we
// resolve whichever actually exists instead of hardcoding one layout — shipping
// the app with a broken gsk runtime is worse than failing the build early.
function gskCommanderDir() {
  const nested = join(__dirname, '../../node_modules/@genspark/cli/node_modules/commander')
  if (existsSync(nested)) return nested
  const hoisted = join(__dirname, '../../node_modules/commander')
  if (existsSync(hoisted)) return hoisted
  return null
}

const gskCommander = gskCommanderDir()
for (const rel of [
  '../../node_modules/@genspark/cli',
  gskCommander && '../../node_modules/commander',
  '../../node_modules/ws',
]) {
  if (!rel) continue
  if (!existsSync(join(__dirname, rel))) {
    throw new Error(
      `electron-builder extraResources source missing: ${rel} (npm hoisting changed?)`,
    )
  }
}
if (!gskCommander) {
  throw new Error(
    'electron-builder: cannot find the commander dependency required by @genspark/cli ' +
      '(checked node_modules/@genspark/cli/node_modules/commander and node_modules/commander)',
  )
}

/** @type {import('electron-builder').Configuration} */
const config = {
  appId: 'com.genoffice.app',
  productName: 'GenOffice',
  electronVersion: '41.7.1',
  directories: {
    output: 'release',
  },
  files: ['out/**'],
  extraResources: [
    {
      from: 'build/THIRD-PARTY-NOTICES.txt',
      to: 'THIRD-PARTY-NOTICES.txt',
    },
    {
      from: '../../node_modules/electron/dist/LICENSES.chromium.html',
      to: 'LICENSES.chromium.html',
    },
    {
      from: '../docs/out',
      to: 'modules/docs',
    },
    {
      from: '../sheets/out',
      to: 'modules/sheets',
    },
    {
      from: '../slides/out',
      to: 'modules/slides',
    },
    {
      from: '../pdf/out',
      to: 'modules/pdf',
    },
    {
      from: '../../node_modules/@genspark/cli',
      to: 'gsk/node_modules/@genspark/cli',
    },
    {
      from: gskCommander,
      to: 'gsk/node_modules/commander',
    },
    {
      from: '../../node_modules/ws',
      to: 'gsk/node_modules/ws',
    },
  ],
  fileAssociations: [
    {
      ext: 'docx',
      name: 'Word Document',
      role: 'Editor',
    },
    {
      ext: 'xlsx',
      name: 'Excel Workbook',
      role: 'Editor',
    },
    {
      ext: 'pptx',
      name: 'PowerPoint Presentation',
      role: 'Editor',
    },
    {
      ext: 'xls',
      name: 'Excel 97-2003 Workbook',
      role: 'Editor',
    },
    {
      ext: 'csv',
      name: 'CSV Document',
      role: 'Editor',
    },
    {
      ext: 'pdf',
      name: 'PDF Document',
      role: 'Editor',
    },
  ],
  npmRebuild: false,
  mac: {
    target: ['dmg', 'zip'],
    category: 'public.app-category.productivity',
    hardenedRuntime: true,
    gatekeeperAssess: false,
    entitlements: 'build/entitlements.mac.plist',
    entitlementsInherit: 'build/entitlements.mac.plist',
    notarize: true,
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        to: 'native/xlsx-sidecar',
      },
    ],
  },
  win: {
    target: [
      {
        target: 'nsis',
        arch: ['x64'],
      },
    ],
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/x86_64-pc-windows-gnu/release/xlsx-sidecar.exe',
        to: 'native/xlsx-sidecar.exe',
      },
    ],
  },
  linux: {
    target: ['AppImage', 'deb'],
    executableName: 'genoffice',
    artifactName: 'genoffice-${version}-${arch}.${ext}',
    category: 'Office',
    maintainer: 'GenOffice',
    synopsis: 'AI-native office suite (docs, sheets, slides, pdf)',
    description:
      'GenOffice is an AI-native office suite: word processor, spreadsheet, presentations, and PDF. ' +
      'The original file is the source of truth; edits are applied as narrow patches so untouched content survives.',
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        to: 'native/xlsx-sidecar',
      },
    ],
  },
  nsis: {
    oneClick: false,
    allowToChangeInstallationDirectory: true,
  },
  dmg: {
    sign: true,
  },
  afterAllArtifactBuild: 'build/notarize-dmg.js',
}

if (updateUrl) {
  config.publish = [
    {
      provider: 'generic',
      url: updateUrl.replace(/\/+$/, ''),
      channel: 'latest',
    },
  ]
}

module.exports = config
