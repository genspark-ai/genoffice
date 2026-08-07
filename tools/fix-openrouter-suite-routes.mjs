import { readFileSync, writeFileSync } from 'node:fs'

function removeGensparkErrorFallback(path, apiName) {
  let src = readFileSync(path, 'utf8')
  const marker = '// Signed-out failures get an inline sign-in button; detected via'
  let changed = false

  while (src.includes(marker)) {
    const start = src.indexOf(marker)
    const call = `void window.${apiName}\n            .aiGskStatus()`
    const callStart = src.indexOf(call, start)
    if (callStart < 0) throw new Error(`Could not locate ${apiName}.aiGskStatus after marker in ${path}`)
    const catchEndToken = `.catch(() => {})`
    const catchEnd = src.indexOf(catchEndToken, callStart)
    if (catchEnd < 0) throw new Error(`Could not locate Genspark fallback end in ${path}`)
    let end = catchEnd + catchEndToken.length
    while (src[end] === '\r' || src[end] === '\n' || src[end] === ' ') end++
    src = src.slice(0, start) + `// OpenRouter errors are shown directly; never convert provider errors into a Genspark login prompt.\n          ` + src.slice(end)
    changed = true
  }

  // Even if an old persisted chat entry contains loginRequired, never render the obsolete login button.
  src = src.replace(/\{entry\.loginRequired && \([\s\S]*?\n\s*\)\}/g, '')

  if (changed) writeFileSync(path, src)
  console.log(`${path}: ${changed ? 'removed Genspark error fallback' : 'no runtime fallback found'}`)
}

removeGensparkErrorFallback('apps/slides/src/renderer/ai/AiPanel.tsx', 'slidesApi')

// Sheets has its own AI panel/route in upstream GenOffice. Patch it when present.
for (const candidate of [
  'apps/sheets/src/renderer/ai/AiPanel.tsx',
  'apps/sheets/src/renderer/src/ai/AiPanel.tsx',
]) {
  try {
    readFileSync(candidate, 'utf8')
    removeGensparkErrorFallback(candidate, 'sheetsApi')
    break
  } catch (err) {
    if (err?.code !== 'ENOENT') throw err
  }
}
