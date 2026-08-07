/** Print-theme CSS: mirrors the editor typography so the PDF matches the canvas */
const PRINT_CSS = `
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
  color: #1f2328;
  font-size: 12pt;
  line-height: 1.7;
  word-wrap: break-word;
}
h1, h2, h3, h4, h5, h6 { line-height: 1.3; margin: 1.3em 0 0.5em; font-weight: 650; break-after: avoid; }
h1 { font-size: 2em; padding-bottom: 0.25em; border-bottom: 1px solid #e4e7eb; }
h2 { font-size: 1.5em; }
h3 { font-size: 1.25em; }
h4 { font-size: 1.1em; }
p { margin: 0.6em 0; }
ul, ol { padding-left: 1.6em; margin: 0.6em 0; }
li > p { margin: 0.15em 0; }
blockquote { margin: 0.8em 0; padding: 0.1em 1em; border-left: 3px solid #d0d5db; color: #57606a; }
code { font-family: 'SF Mono', Menlo, Consolas, monospace; font-size: 0.875em; background: rgba(129,139,152,0.14); border-radius: 4px; padding: 0.15em 0.35em; }
pre { background: #f6f8fa; border: 1px solid #e4e7eb; border-radius: 8px; padding: 12px 16px; margin: 0.8em 0; white-space: pre-wrap; break-inside: avoid; }
pre code { background: none; padding: 0; font-size: 0.85em; line-height: 1.6; }
hr { border: none; border-top: 2px solid #e4e7eb; margin: 1.6em 0; }
a { color: #0a69da; }
img { max-width: 100%; height: auto; }
table { border-collapse: collapse; margin: 0.8em 0; break-inside: avoid; }
th, td { border: 1px solid #d0d5db; padding: 5px 10px; }
th { background: #f6f8fa; font-weight: 600; }
ul[data-type='taskList'] { list-style: none; padding-left: 0.4em; }
ul[data-type='taskList'] li { display: flex; gap: 8px; }
ul[data-type='taskList'] li > label { flex: 0 0 auto; margin-top: 0.3em; }
ul[data-type='taskList'] li[data-checked='true'] > div { color: #8b929b; text-decoration: line-through; }
.md-callout { margin: 0.8em 0; padding: 10px 14px; border-radius: 8px; border-left: 4px solid #3b82f6; background: rgba(59,130,246,0.08); break-inside: avoid; }
.md-callout[data-type='tip'] { border-left-color: #10b981; background: rgba(16,185,129,0.08); }
.md-callout[data-type='warning'] { border-left-color: #f59e0b; background: rgba(245,158,11,0.1); }
.md-callout[data-type='danger'] { border-left-color: #ef4444; background: rgba(239,68,68,0.08); }
.md-toggle { margin: 0.6em 0; }
.print-toggle-summary { font-weight: 600; margin: 0.4em 0 0.2em; }
.md-toggle-body { margin-left: 1.2em; }
`

/**
 * Build a self-contained print HTML document from the live editor DOM.
 * Toggles are expanded and their summary inputs flattened to text; editor-only
 * chrome (contenteditable, drag handles) never makes it into the clone.
 */
export function buildPrintHtml(editorRoot: HTMLElement, title: string): string {
  const clone = editorRoot.cloneNode(true) as HTMLElement
  clone.removeAttribute('contenteditable')
  for (const el of clone.querySelectorAll('[contenteditable]'))
    el.removeAttribute('contenteditable')

  // editor-only code block chrome (language picker + copy button) must not print
  for (const bar of clone.querySelectorAll('.md-codeblock-bar')) bar.remove()

  // cloneNode does not carry input value properties — walk original and clone
  // toggles in parallel (same document order) to recover the live summaries
  const liveToggles = editorRoot.querySelectorAll('.md-toggle')
  const cloneToggles = clone.querySelectorAll('.md-toggle')
  cloneToggles.forEach((toggle, i) => {
    const header = toggle.querySelector('.md-toggle-header')
    if (header) {
      const summary = document.createElement('div')
      summary.className = 'print-toggle-summary'
      summary.textContent =
        liveToggles[i]?.querySelector<HTMLInputElement>('.md-toggle-summary')?.value ?? ''
      header.replaceWith(summary)
    }
    const body = toggle.querySelector<HTMLElement>('.md-toggle-body')
    if (body) body.style.display = ''
  })

  const escapedTitle = title.replace(/&/g, '&amp;').replace(/</g, '&lt;')
  return [
    '<!doctype html>',
    '<html>',
    '<head>',
    '<meta charset="utf-8">',
    `<title>${escapedTitle}</title>`,
    `<style>${PRINT_CSS}</style>`,
    '</head>',
    `<body>${clone.innerHTML}</body>`,
    '</html>',
  ].join('\n')
}
