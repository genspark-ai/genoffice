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

// The gsk CLI tree below is copied verbatim from node_modules, and the
// nested commander path depends on npm's current hoisting layout — fail the
// build with a clear message if an install ever changes it, instead of
// shipping an installer with a broken gsk runtime.
// LICENSES.chromium.html only exists after the Electron binary download —
// since Electron 42 that no longer happens during `npm ci` (the postinstall
// script was replaced by the lazy `install-electron` bin), and electron-builder
// exits 0 on a missing extraResources source, so without this check the
// installer would silently ship without the Chromium license.
for (const rel of [
  '../../node_modules/@genspark/cli',
  '../../node_modules/@genspark/cli/node_modules/commander',
  '../../node_modules/ws',
  '../../node_modules/electron/dist/LICENSES.chromium.html',
]) {
  if (!existsSync(join(__dirname, rel))) {
    throw new Error(
      `electron-builder extraResources source missing: ${rel} (npm hoisting changed?)`,
    )
  }
}

// The module trees are electron-vite outputs produced by build:all; a missing
// one means that module's build did not run or failed. electron-builder only
// logs "file source doesn't exist" for an absent extraResources source and
// still exits 0, so without this the installer launches normally and is simply
// missing that editor — it surfaces only when a user opens the tab.
//
// Runs from the beforePack hook, not at module load: gen-third-party-notices
// requires this config to read extraResources, and the dist:* scripts run
// notices before build:all, when the out dirs legitimately don't exist yet.
function assertModuleTreesPresent() {
  for (const rel of [
    '../docs/out',
    '../sheets/out',
    '../slides/out',
    '../pdf/out',
    '../markdown/out',
  ]) {
    if (!existsSync(join(__dirname, rel))) {
      throw new Error(
        `electron-builder extraResources source missing: ${rel} (run npm run build:all first)`,
      )
    }
  }
}

/** @type {import('electron-builder').Configuration} */
const config = {
  appId: 'com.genoffice.app',
  productName: 'GenOffice',
  // Resolved from the installed electron package so dependency bumps can
  // never leave a stale hard-coded pin behind (packaging would silently ship
  // the old runtime).
  electronVersion: require('electron/package.json').version,
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
      from: '../markdown/out',
      to: 'modules/markdown',
    },
    {
      from: '../../node_modules/@genspark/cli',
      to: 'gsk/node_modules/@genspark/cli',
    },
    {
      from: '../../node_modules/@genspark/cli/node_modules/commander',
      to: 'gsk/node_modules/commander',
    },
    {
      from: '../../node_modules/ws',
      to: 'gsk/node_modules/ws',
    },
  ],
  // `mimeType` is read only by the Linux target, where it becomes the
  // desktop entry's MimeType= list; associations without it are dropped
  // there. macOS and Windows ignore the field and key off `ext`.
  fileAssociations: [
    {
      ext: 'docx',
      name: 'Word Document',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    {
      ext: 'xlsx',
      name: 'Excel Workbook',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    {
      ext: 'pptx',
      name: 'PowerPoint Presentation',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    },
    {
      ext: 'xls',
      name: 'Excel 97-2003 Workbook',
      role: 'Editor',
      mimeType: 'application/vnd.ms-excel',
    },
    {
      ext: 'csv',
      name: 'CSV Document',
      role: 'Editor',
      mimeType: 'text/csv',
    },
    {
      ext: 'pdf',
      name: 'PDF Document',
      role: 'Editor',
      mimeType: 'application/pdf',
    },
    {
      ext: 'md',
      name: 'Markdown Document',
      role: 'Editor',
      mimeType: 'text/markdown',
    },
    {
      ext: 'markdown',
      name: 'Markdown Document',
      role: 'Editor',
      mimeType: 'text/markdown',
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
  // Unlike win (which cross-compiles the sidecar to an explicit target
  // triple), linux takes it from cargo's host-native target/release/ — the
  // same source mac uses. So no `arch` is pinned here: electron-builder
  // defaults to the build host's architecture, which is the only one the
  // sidecar was actually built for. Packaging arm64 on an x64 host, or the
  // reverse, needs a matching `cargo build --target` first.
  linux: {
    // AppImage only: it needs no packaging identity, whereas deb/rpm would
    // require a Debian maintainer and homepage in the repo metadata.
    target: ['AppImage'],
    category: 'Office',
    icon: 'build/icon.png',
    // mac and win name the binary from productName; linux instead derives it
    // from package.json "name", and "@genoffice/shell" sanitizes to the
    // invalid "@genofficeshell". Setting it explicitly also makes the
    // generated genoffice.desktop match the WM_CLASS Electron reports (it
    // takes that from the executable basename), so the running window links
    // back to its launcher entry.
    executableName: 'genoffice',
    // Electron takes its X11 app_id from package.json "desktopName"
    // (genoffice.desktop); syncDesktopName makes electron-builder name the
    // .desktop file and its StartupWMClass from the same value. Without it
    // StartupWMClass falls back to productName ("GenOffice"), which does not
    // match the "genoffice" WM_CLASS the window actually reports — and X11
    // compares case-sensitively, so the taskbar shows an unlinked window.
    syncDesktopName: true,
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
  beforePack: async () => {
    assertModuleTreesPresent()
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
