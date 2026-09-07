/** Shared rhwp-studio snapshot helpers for the vendor script and tests. */

export const PWA_FILES = ['sw.js', 'registerSW.js', 'manifest.webmanifest']

export const REQUIRED_RELATIVE = [
  'index.html',
  'fonts/NotoSansKR-Regular.woff2',
  'fonts/Pretendard-Regular.woff2',
]

export const REQUIRED_ASSET_EXTS = ['.js', '.wasm']

const PWA_HTML_RE =
  /<link\s+rel="manifest"[^>]*>|<script[^>]*(?:id="vite-plugin-pwa:register-sw"|src="[^"]*registerSW\.js")[^>]*><\/script>/g

export function isPwaPath(urlPath) {
  const name = urlPath.split('?')[0].split('/').pop() ?? ''
  return PWA_FILES.includes(name) || /^workbox-.*\.js$/.test(name)
}

export function stripPwaHtml(html) {
  return html.replace(PWA_HTML_RE, '')
}

/**
 * Embed mode strips File new/open/save from the registry so the host owns those
 * actions — and also skips boot-time `createNewDocument()`. Untitled tabs then
 * have no pages. Keep only `file:new-doc` registered; OA() still hides the menu.
 */
export const EMBED_NEW_DOC_MARK = '/*genoffice-embed-new-doc*/'

const EMBED_NEW_DOC_RE =
  /bA\.registerAll\(yA===`embed`\?Ev\.filter\(e=>!sD\.includes\(e\.id\)\):Ev\)/

export function keepEmbedNewDoc(js) {
  if (js.includes(EMBED_NEW_DOC_MARK)) return js
  if (!js.includes('file:new-doc') || !js.includes('registerAll')) return js
  const next = js.replace(
    EMBED_NEW_DOC_RE,
    `bA.registerAll(yA===\`embed\`?Ev.filter(e=>${EMBED_NEW_DOC_MARK}e.id===\`file:new-doc\`||!sD.includes(e.id)):Ev)`,
  )
  if (next === js) {
    throw new Error('rhwp-studio embed command filter changed — update keepEmbedNewDoc()')
  }
  return next
}

export function hasEmbedNewDoc(js) {
  return js.includes(EMBED_NEW_DOC_MARK)
}

/**
 * Studio computes paragraph SHA fences inside getSelectionContext but throws
 * them away. The public SDK cannot add fields there (exactKeys), so expose a
 * sibling prepareTextCommand for the host to bind applyTextCommand.
 */
export const PREPARE_TEXT_MARK = '/*genoffice-prepare-text*/'
export const PREPARE_TEXT_V2_MARK = '/*genoffice-prepare-text-v2*/'
export const PREPARE_TEXT_V3_MARK = '/*genoffice-prepare-text-v3*/'

const PREPARE_SNAP_RE = /try\{([A-Za-z_$][\w$]*)\(this\.deps\.wasm,i\),this\.currentFormat\(\)/
const PREPARE_CLASS_RE =
  /selectedTextSha256:([A-Za-z_$][\w$]*)\}\}async applyTextCommand\(([A-Za-z_$][\w$]*)\)\{/
const PREPARE_HANDLER_RE =
  /async getSelectionContext\(\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.getSelectionContext\(\)\},async applyTextCommand\(/
const PREPARE_ROUTE_RE =
  /case`getSelectionContext`:return ([A-Za-z_$][\w$]*)\(([A-Za-z_$][\w$]*),`getSelectionContext params`\),([A-Za-z_$][\w$]*)\.getSelectionContext\(\);case`applyTextCommand`:/
const PREPARE_V1_CLASS_RE =
  /\/\*genoffice-prepare-text\*\/prepareTextCommand\(\)\{[\s\S]*?\}\}async applyTextCommand\(([A-Za-z_$][\w$]*)\)\{/
const PREPARE_V1_ROUTE_RE =
  /case`getSelectionContext`:return ([A-Za-z_$][\w$]*)\(([A-Za-z_$][\w$]*),`getSelectionContext params`\),([A-Za-z_$][\w$]*)\.getSelectionContext\(\);case`prepareTextCommand`:return \3\.prepareTextCommand\(\);case`applyTextCommand`:/
const SET_FIELD_CLASS_RE =
  /setField\(e,t\)\{this\.syncGeneration\(\);return this\.deps\.wasm\.setFieldValueByName\([^;]+\)\}async applyTextCommand/
const SET_FIELD_HANDLER_RE =
  /async setField\(e,t\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.setField\(e,t\)\},async applyTextCommand\(/
const SET_FIELD_ROUTE_RE =
  /case`setField`:return ([A-Za-z_$][\w$]*)\.setField\(([A-Za-z_$][\w$]*)\.name,\2\.value\);case`applyTextCommand`:/

/** Per-paragraph table-control scan cap. Controls past this index are skipped. */
const TABLE_CONTROLS_PER_PARA = 8

function tableAgentMethods() {
  return `listTables(){this.syncGeneration();let e=this.deps.wasm,t=[];for(let n=0;n<e.getSectionCount();n+=1)for(let r=0;r<e.getParagraphCount(n);r+=1)for(let i=0;i<${TABLE_CONTROLS_PER_PARA};i+=1){let a;try{a=e.getTableDimensions(n,r,i);if(typeof a==\`string\`)a=JSON.parse(a)}catch{continue}if(!a||!a.rowCount)continue;let o=[],s=Number(a.cellCount||0);for(let c=0;c<s;c+=1){try{let l=e.getCellInfo(n,r,i,c);if(typeof l==\`string\`)l=JSON.parse(l);let u=e.getCellParagraphCount(n,r,i,c),d=[];for(let f=0;f<u;f+=1){let p=e.getCellParagraphLength(n,r,i,c,f);d.push(p>0?e.getTextInCell(n,r,i,c,f,0,p):\`\`)}o.push({index:c,row:l.row,col:l.col,text:d.join(\`\\n\`)})}catch{}}t.push({section:n,paragraph:r,control:i,rows:a.rowCount,cols:a.colCount,cells:o})}return t}replaceCell(e,t,n,r,i){this.syncGeneration();let a=this.deps.wasm,o=a.getCellParagraphLength(e,t,n,r,0),s=a.replaceTextInCellDeferredPagination(e,t,n,r,0,0,o,String(i??\`\`));if(typeof s==\`string\`)try{s=JSON.parse(s)}catch{}return s}`
}

function prepareAgentMethods(snap) {
  return `${PREPARE_TEXT_V3_MARK}prepareTextCommand(){let e=this.getSelectionContext(),n=this.deps.input.getSelection(),r=null,i=null;if(e.target&&n&&n.start&&n.end&&n.start.sectionIndex===e.target.section&&n.start.paragraphIndex===e.target.paragraph&&n.end.sectionIndex===e.target.section&&n.end.paragraphIndex===e.target.paragraph&&n.end.charOffset>n.start.charOffset){r=n.start.charOffset,i=n.end.charOffset}if(!e.editable||!e.target)return{editable:!1,reason:\`not_editable\`,target:e.target,text:null,textSha256:null,formatSha256:null,adjacentContextSha256:null,selectionStart:r,selectionEnd:i};try{let t=${snap}(this.deps.wasm,e.target);return{editable:!0,reason:null,target:e.target,text:t.text,textSha256:t.textSha256,formatSha256:t.formatSha256,adjacentContextSha256:t.adjacentContextSha256,selectionStart:r,selectionEnd:i}}catch(a){return{editable:!1,reason:String(a&&a.message||a),target:e.target,text:null,textSha256:null,formatSha256:null,adjacentContextSha256:null,selectionStart:r,selectionEnd:i}}listBodyParagraphs(){this.syncGeneration();let e=this.deps.wasm,t=[];for(let n=0;n<e.getSectionCount();n+=1)for(let r=0;r<e.getParagraphCount(n);r+=1){let i=e.getParagraphLength(n,r),a={kind:\`body_paragraph\`,section:n,paragraph:r,charOffset:0,length:i};try{let o=${snap}(e,a);t.push({editable:!0,reason:null,target:a,text:o.text,textSha256:o.textSha256,formatSha256:o.formatSha256,adjacentContextSha256:o.adjacentContextSha256})}catch(s){t.push({editable:!1,reason:String(s&&s.message||s),target:a,text:null,textSha256:null,formatSha256:null,adjacentContextSha256:null})}}return t}listFields(){this.syncGeneration();let e=this.deps.wasm.getFieldList();if(typeof e==\`string\`)try{e=JSON.parse(e)}catch{e=[]}if(!Array.isArray(e))return[];return e.map(t=>{let n=t&&(t.name||t.fieldName||t.fieldId||t.id)||\`\`,r=\`\`;if(n)try{let i=this.deps.wasm.getFieldValueByName(n);r=typeof i==\`string\`?i:i==null?\`\`:String(i)}catch{}return{name:n,value:r,type:t&&t.fieldType||null}}).filter(t=>t.name)}setField(e,t){this.syncGeneration();return this.deps.wasm.setFieldValueByName(String(e??\`\`),String(t??\`\`))}${tableAgentMethods()}`
}

function prepareAgentHandlers(ready, agent) {
  return `async getSelectionContext(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.getSelectionContext()},async prepareTextCommand(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.prepareTextCommand()},async listBodyParagraphs(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listBodyParagraphs()},async listFields(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listFields()},async setField(e,t){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.setField(e,t)},async listTables(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listTables()},async replaceCell(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.replaceCell(e,t,n,r,i)},async applyTextCommand(`
}

function prepareAgentRoutes(guard, params, host) {
  return `case\`getSelectionContext\`:return ${guard}(${params},\`getSelectionContext params\`),${host}.getSelectionContext();case\`prepareTextCommand\`:return ${host}.prepareTextCommand();case\`listBodyParagraphs\`:return ${host}.listBodyParagraphs();case\`listFields\`:return ${host}.listFields();case\`setField\`:return ${host}.setField(${params}.name,${params}.value);case\`listTables\`:return ${host}.listTables();case\`replaceCell\`:return ${host}.replaceCell(${params}.section,${params}.paragraph,${params}.control,${params}.cellIndex,${params}.text);case\`applyTextCommand\`:`
}

function stripClassMethodCommas(js) {
  return js
    .replace(/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/g, 'listBodyParagraphs(){this.syncGeneration()')
    .replace(/,listFields\(\)\{this\.syncGeneration\(\)/g, 'listFields(){this.syncGeneration()')
    .replace(/,setField\(e,t\)\{this\.syncGeneration\(\)/g, 'setField(e,t){this.syncGeneration()')
    .replace(/,listTables\(\)\{this\.syncGeneration\(\)/g, 'listTables(){this.syncGeneration()')
    .replace(/,replaceCell\(e,t,n,r,i\)\{this\.syncGeneration\(\)/g, 'replaceCell(e,t,n,r,i){this.syncGeneration()')
}

function attachTableSurface(js) {
  let next = stripClassMethodCommas(js.replace(PREPARE_TEXT_V2_MARK, PREPARE_TEXT_V3_MARK))
  next = next.replace(
    SET_FIELD_CLASS_RE,
    `setField(e,t){this.syncGeneration();return this.deps.wasm.setFieldValueByName(String(e??\`\`),String(t??\`\`))}${tableAgentMethods()}async applyTextCommand`,
  )
  next = next.replace(SET_FIELD_HANDLER_RE, (_, ready, agent) => {
    return `async setField(e,t){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.setField(e,t)},async listTables(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listTables()},async replaceCell(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.replaceCell(e,t,n,r,i)},async applyTextCommand(`
  })
  next = next.replace(
    SET_FIELD_ROUTE_RE,
    'case`setField`:return $1.setField($2.name,$2.value);case`listTables`:return $1.listTables();case`replaceCell`:return $1.replaceCell($2.section,$2.paragraph,$2.control,$2.cellIndex,$2.text);case`applyTextCommand`:',
  )
  return next
}

export function exposePrepareTextCommand(js) {
  const classComplete =
    js.includes(PREPARE_TEXT_V3_MARK) &&
    js.includes('listTables(){this.syncGeneration()') &&
    !/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/.test(js)
  if (classComplete) return js
  if (js.includes(PREPARE_TEXT_V3_MARK) || js.includes(PREPARE_TEXT_V2_MARK)) {
    const upgraded = attachTableSurface(js)
    if (!upgraded.includes(PREPARE_TEXT_V3_MARK) || !upgraded.includes('case`listTables`')) {
      throw new Error('rhwp-studio prepareTextCommand surface changed — update exposePrepareTextCommand()')
    }
    return upgraded
  }
  const snap = js.match(PREPARE_SNAP_RE)
  if (!snap) throw new Error('rhwp-studio paragraph snapshot helper changed — update exposePrepareTextCommand()')
  const methods = prepareAgentMethods(snap[1])
  let next = js
  if (PREPARE_V1_CLASS_RE.test(next)) {
    next = next.replace(PREPARE_V1_CLASS_RE, `${methods}async applyTextCommand($1){`)
  } else {
    next = next.replace(PREPARE_CLASS_RE, `selectedTextSha256:$1}}${methods}async applyTextCommand($2){`)
  }
  next = next.replace(PREPARE_HANDLER_RE, (_, ready, agent) => prepareAgentHandlers(ready, agent))
  if (next.includes('async prepareTextCommand(){if(await') && !next.includes('async listBodyParagraphs(){if(await')) {
    next = next.replace(
      /async prepareTextCommand\(\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.prepareTextCommand\(\)\},async applyTextCommand\(/,
      (_, ready, agent) =>
        `async prepareTextCommand(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.prepareTextCommand()},async listBodyParagraphs(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listBodyParagraphs()},async listFields(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listFields()},async setField(e,t){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.setField(e,t)},async listTables(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listTables()},async replaceCell(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.replaceCell(e,t,n,r,i)},async applyTextCommand(`,
    )
  }
  next = next.replace(PREPARE_ROUTE_RE, (_, guard, params, host) => prepareAgentRoutes(guard, params, host))
  if (PREPARE_V1_ROUTE_RE.test(next) && !next.includes('case`listBodyParagraphs`')) {
    next = next.replace(PREPARE_V1_ROUTE_RE, (_, guard, params, host) => prepareAgentRoutes(guard, params, host))
  }
  if (!next.includes(PREPARE_TEXT_V3_MARK) || !next.includes('case`listTables`') || !next.includes('case`setField`')) {
    throw new Error('rhwp-studio prepareTextCommand surface changed — update exposePrepareTextCommand()')
  }
  return next
}

export function hasPrepareTextCommand(js) {
  return (
    js.includes(PREPARE_TEXT_V3_MARK) ||
    js.includes(PREPARE_TEXT_V2_MARK) ||
    js.includes(PREPARE_TEXT_MARK)
  )
}

/**
 * Drop abandoned page-turn experiments from a local snapshot. Stock studio
 * behavior is restored; only `file:new-doc` stays patched.
 */
export function stripAbandonedStudioPatches(js) {
  let next = js
  next = next.replace(
    /\/\*genoffice-eager-prefetch\*\/n\(\)/g,
    'if(typeof r.requestIdleCallback==`function`){this.deferredPrefetchTask={kind:`idle`,id:r.requestIdleCallback(n,{timeout:1e3})};return}this.deferredPrefetchTask={kind:`timeout`,id:window.setTimeout(n,250)}',
  )
  next = next.replace(
    /\/\*genoffice-prefetch-overscan\*\/for\(let e of\[s-2,s-1,c\+1,c\+2\]\)/g,
    'for(let e of[s-1,c+1])',
  )
  next = next.replace(
    /\/\*genoffice-page-align\*\/n>0\?([A-Za-z_$][\w$]*)\(e,r,o,a\):([A-Za-z_$][\w$]*)\(e,r,o\)/g,
    'n>0?Math.min($1(e,r,o,a),r+i):Math.max($2(e,r,o),r-i)',
  )
  next = next.replace(
    /flushDeferredPaginationIfNeeded\(`before-navigation`,\/\*genoffice-nav-pagination\*\/!0\)/g,
    'flushDeferredPaginationIfNeeded(`before-navigation`,!1)',
  )
  next = next.replace(
    /e!==`document-agent-rendered`&&\/\*genoffice-sync-layout\*\/\(typeof e==`string`&&e\.includes\(`deferred-pagination-flush`\)\?this\.refreshPages\(\):this\.refreshPagesForMutation\(\)\)/g,
    'e!==`document-agent-rendered`&&this.refreshPagesForMutation()',
  )
  next = next.replace(
    /this\.recalcLayout\(\),\/\*genoffice-clamp-scroll\*\/\(\(\)=>\{let e=this\.viewportManager\.getViewportSize\(\),t=Math\.max\(0,this\.virtualScroll\.getTotalHeight\(\)-e\.height\);this\.viewportManager\.getScrollY\(\)>t&&this\.viewportManager\.setScrollTop\(t\)\}\)\(\),this\.cancelPendingTextEditRefresh\(\)/g,
    'this.recalcLayout(),this.cancelPendingTextEditRefresh()',
  )
  next = next.replace(
    /vp\.call\(this,e\.key===`PageUp`\?-1:1,e\.shiftKey\),\/\*genoffice-caret-after-page\*\/this\.updateCaret\(\);return\}/g,
    'vp.call(this,e.key===`PageUp`?-1:1,e.shiftKey);return}',
  )
  return next
}
