import { writeFileSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, BrowserWindow } from 'electron'
import { PNG } from 'pngjs'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const svgDir = join(root, 'apps/shell/src/main/assets/tab-actions/svg')
const pngDir = join(root, 'apps/shell/src/main/assets/tab-actions/png')

mkdirSync(svgDir, { recursive: true })
mkdirSync(pngDir, { recursive: true })

/**
 * 18 Action SVGs designed on a 24x24 grid with 1.75px stroke,
 * rounded caps and joins, outline style matching GenOffice.
 */
const SVGS = {
  save: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
  <polyline points="17 21 17 13 7 13 7 21"/>
  <polyline points="7 3 7 8 15 8"/>
</svg>`,

  saveAs: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M14 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v3"/>
  <polyline points="7 3 7 8 15 8"/>
  <path d="M18.5 12.5a1.77 1.77 0 0 1 2.5 2.5L15 21l-4 1 1-4 6.5-5.5z"/>
</svg>`,

  print: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M6 9V3h12v6"/>
  <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
  <rect x="6" y="14" width="12" height="7"/>
</svg>`,

  rename: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M12 20h9"/>
  <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>
</svg>`,

  duplicate: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M15 2H6a2 2 0 0 0-2 2v12"/>
  <rect x="8" y="6" width="12" height="16" rx="2"/>
</svg>`,

  splitBeside: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <rect x="3" y="4" width="18" height="16" rx="2"/>
  <line x1="12" y1="4" x2="12" y2="20"/>
</svg>`,

  openInNewWindow: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>
  <polyline points="15 3 21 3 21 9"/>
  <line x1="10" y1="14" x2="21" y2="3"/>
</svg>`,

  copyPath: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
  <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>
</svg>`,

  copyName: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <rect x="9" y="9" width="12" height="12" rx="2"/>
  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
</svg>`,

  reveal: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2z"/>
  <path d="M2 10h20"/>
</svg>`,

  openDefault: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <rect x="3" y="3" width="18" height="18" rx="2"/>
  <path d="m8 16 8-8"/>
  <path d="M11 8h5v5"/>
</svg>`,

  moveLeft: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="m12 19-7-7 7-7"/>
  <path d="M19 12H5"/>
</svg>`,

  moveRight: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="m12 5 7 7-7 7"/>
  <path d="M5 12h14"/>
</svg>`,

  close: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M18 6 6 18"/>
  <path d="m6 6 12 12"/>
</svg>`,

  closeOthers: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <rect x="3" y="8" width="10" height="12" rx="2"/>
  <path d="M13 4h6a2 2 0 0 1 2 2v10"/>
  <path d="m15 13 4 4m0-4-4 4"/>
</svg>`,

  closeRight: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <rect x="3" y="6" width="9" height="12" rx="2"/>
  <path d="M15 12h6m-3-3 3 3-3 3"/>
</svg>`,

  closeAll: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M8 4h10a2 2 0 0 1 2 2v10"/>
  <rect x="4" y="8" width="12" height="12" rx="2"/>
  <path d="m8 12 4 4m0-4-4 4"/>
</svg>`,

  reopen: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/>
  <path d="M3 3v5h5"/>
</svg>`,
}

// Light theme stroke: dark charcoal (#1f2328)
const COLOR_LIGHT = '#1f2328'
// Dark theme stroke: bright clean light gray (#f0f0f0)
const COLOR_DARK = '#f0f0f0'

app.whenReady().then(async () => {
  try {
    // 1. Write SVGs
    for (const [name, svg] of Object.entries(SVGS)) {
      writeFileSync(join(svgDir, `${name}.svg`), svg, 'utf8')
    }
    console.log(`Saved ${Object.keys(SVGS).length} SVGs to ${svgDir}`)

    // 2. Offscreen browser for rasterization
    const win = new BrowserWindow({
      width: 64,
      height: 64,
      show: false,
      transparent: true,
      frame: false,
      webPreferences: { offscreen: true },
    })

    const render = async (svgContent, size, color) => {
      const styledSvg = svgContent.replace('stroke="currentColor"', `stroke="${color}"`)
      const html = `<!DOCTYPE html><html><head><style>
        * { margin:0; padding:0; box-sizing:border-box; }
        html, body { width:${size}px; height:${size}px; background:transparent; overflow:hidden; }
      </style></head><body>
        <div style="width:${size}px;height:${size}px;display:flex;align-items:center;justify-content:center;">
          ${styledSvg.replace('<svg ', `<svg width="${size}" height="${size}" `)}
        </div>
      </body></html>`

      await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`)
      await new Promise((r) => setTimeout(r, 40))
      const img = await win.webContents.capturePage({ x: 0, y: 0, width: size, height: size })
      return img.toPNG()
    }

    let renderedCount = 0
    for (const [name, svg] of Object.entries(SVGS)) {
      // 1x Light (16x16)
      const png1xLight = await render(svg, 16, COLOR_LIGHT)
      writeFileSync(join(pngDir, `${name}.png`), png1xLight)

      // 2x Light (32x32)
      const png2xLight = await render(svg, 32, COLOR_LIGHT)
      writeFileSync(join(pngDir, `${name}@2x.png`), png2xLight)

      // 1x Dark (16x16)
      const png1xDark = await render(svg, 16, COLOR_DARK)
      writeFileSync(join(pngDir, `${name}-dark.png`), png1xDark)

      // 2x Dark (32x32)
      const png2xDark = await render(svg, 32, COLOR_DARK)
      writeFileSync(join(pngDir, `${name}-dark@2x.png`), png2xDark)

      // Validate transparency
      const p1 = PNG.sync.read(png1xLight)
      const p2 = PNG.sync.read(png2xLight)
      if (p1.data[3] !== 0 || p2.data[3] !== 0) {
        throw new Error(`Top-left corner of ${name} is not transparent!`)
      }
      renderedCount += 4
    }

    console.log(`Successfully generated and verified ${renderedCount} PNG assets in ${pngDir}`)
    app.quit()
  } catch (err) {
    console.error('Rasterization failed:', err)
    process.exit(1)
  }
})
