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
export const PREPARE_TEXT_V4_MARK = '/*genoffice-prepare-text-v4*/'
export const PREPARE_TEXT_V5_MARK = '/*genoffice-prepare-text-v5*/'
export const PREPARE_TEXT_V6_MARK = '/*genoffice-prepare-text-v6*/'
export const PREPARE_TEXT_V7_MARK = '/*genoffice-prepare-text-v7*/'

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

const INSERT_SPLIT_NEEDLE =
  'insertBodyParagraphs(e,t,n){this.syncGeneration();let r=this.deps.wasm,i=Number(n);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||i<1)throw Error(`insert count must be a positive integer`);let s=r.getParagraphCount(e);'

function insertBodyMethod() {
  // WasmBridge has splitParagraph (Enter) but not insertParagraph.
  // Split the predecessor at its end so a locked first paragraph is not cloned as N locked rows.
  return `insertBodyParagraphs(e,t,n){this.syncGeneration();let r=this.deps.wasm,i=Number(n);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||i<1)throw Error(\`insert count must be a positive integer\`);let s=r.getParagraphCount(e);if(!s)throw Error(\`문서가 로드되지 않았습니다\`);if(t<1){for(let a=0;a<i;a+=1){let o=r.splitParagraph(e,0,0);if(typeof o==\`string\`)try{o=JSON.parse(o)}catch{}if(o&&o.ok===!1)throw Error(String(o.error||o.message||\`splitParagraph failed\`))}}else{let p=Math.min(Math.max(t,1),s)-1,l=r.getParagraphLength(e,p);for(let a=0;a<i;a+=1){let o=r.splitParagraph(e,p,l);if(typeof o==\`string\`)try{o=JSON.parse(o)}catch{}if(o&&o.ok===!1)throw Error(String(o.error||o.message||\`splitParagraph failed\`))}}return{section:e,index:t,count:i}}`
}

function insertFilledMethod() {
  // One RPC: insert empties, then applyTextCommand each line (re-list before every apply).
  return `async insertFilledParagraphs(e,t,n){if(!Array.isArray(n)||n.length<1)throw Error(\`insert texts must be a non-empty array\`);this.insertBodyParagraphs(e,t,n.length);let r=0,i=t,s=0;while(r<n.length){this.syncGeneration();let a=this.listBodyParagraphs(),o=null;for(let c=0;c<a.length;c+=1){let l=a[c];if(l&&l.editable&&l.target&&l.target.section===e&&l.target.paragraph>=i){o=l;break}}if(!o){s+=1;if(s>n.length+4)throw Error(\`paragraph is not editable\`);let u=a[a.length-1];this.insertBodyParagraphs(u&&u.target?u.target.section:e,u&&u.target?u.target.paragraph+1:t,n.length-r);continue}let d=this.getDocumentState();await this.applyTextCommand({schemaVersion:1,commandId:crypto.randomUUID(),expectedDocumentEpoch:d.documentEpoch,expectedChangeSeq:d.changeSeq,expectedDocumentSha256:d.documentSha256,target:o.target,expectedBeforeSha256:o.textSha256,expectedFormatSha256:o.formatSha256,expectedAdjacentContextSha256:o.adjacentContextSha256,replacement:String(n[r]??\`\`)});r+=1;i=o.target.paragraph+1;s=0}return{section:e,index:t,count:n.length}}`
}

function charFormatMethod() {
  return `applyBodyCharFormat(e,t,n,r,i){this.syncGeneration();let a=i&&typeof i==\`object\`?Object.assign({},i):{};if(a.fontName){let o=this.deps.wasm.findOrCreateFontId(String(a.fontName));if(!(o>=0))throw Error(\`font not found\`);a.fontId=o;delete a.fontName}let s=this.deps.wasm.applyCharFormat(e,t,n,r,JSON.stringify(a));if(typeof s==\`string\`)try{s=JSON.parse(s)}catch{}if(s&&s.ok===!1)throw Error(String(s.error||s.message||\`applyCharFormat failed\`));return s}`
}

function formatAgentMethods() {
  // Dialog-free table + char/para format. Wasm bridge already wraps createTable / apply*.
  return `insertTable(e,t,n,r){this.syncGeneration();let i=Number(n),s=Number(r);if(!Number.isInteger(e)||!Number.isInteger(t)||!Number.isInteger(i)||!Number.isInteger(s)||i<1||s<1)throw Error(\`table size must be positive integers\`);if(i>20||s>10)throw Error(\`table is too large\`);let o=0;try{o=Number(this.deps.wasm.getParagraphLength(e,t))||0}catch{o=0}let a=this.deps.wasm.createTable(e,t,o,i,s);if(typeof a==\`string\`)try{a=JSON.parse(a)}catch{}if(a&&a.ok===!1)throw Error(String(a.error||a.message||\`createTable failed\`));return{section:e,paragraph:Number(a?.paraIdx??t),control:Number(a?.controlIdx??0),rows:i,cols:s}}${charFormatMethod()}applyBodyParaFormat(e,t,n){this.syncGeneration();let r=n&&typeof n==\`object\`?Object.assign({},n):{};if(r.headType===\`Bullet\`){r.numberingId=this.deps.wasm.ensureDefaultBullet(r.bulletChar||\`●\`);r.paraLevel=0;delete r.bulletChar}else if(r.headType===\`Number\`){r.numberingId=this.deps.wasm.ensureDefaultNumbering();r.paraLevel=0}let i=this.deps.wasm.applyParaFormat(e,t,JSON.stringify(r));if(typeof i==\`string\`)try{i=JSON.parse(i)}catch{}if(i&&i.ok===!1)throw Error(String(i.error||i.message||\`applyParaFormat failed\`));return i}`
}

function insertAgentMethod() {
  return `${insertBodyMethod()}${insertFilledMethod()}${formatAgentMethods()}`
}

function replaceCellMethod() {
  // Deferred cell replace patches the page tree; a burst of fills traps WASM.
  return `replaceCell(e,t,n,r,i){this.syncGeneration();let a=this.deps.wasm,o=0;try{o=a.getCellParagraphLength(e,t,n,r,0)}catch{o=0}o=Number(o)||0;if(o>0){let d=a.deleteTextInCell(e,t,n,r,0,0,o);if(typeof d==\`string\`)try{d=JSON.parse(d)}catch{}if(d&&d.ok===!1)throw Error(String(d.error||d.message||\`deleteTextInCell failed\`))}let x=String(i??\`\`);if(!x)return{ok:!0};let s=a.insertTextInCell(e,t,n,r,0,0,x);if(typeof s==\`string\`)try{s=JSON.parse(s)}catch{}if(s&&s.ok===!1)throw Error(String(s.error||s.message||\`insertTextInCell failed\`));return s}`
}

function tableAgentMethods() {
  return `listTables(){this.syncGeneration();let e=this.deps.wasm,t=[];for(let n=0;n<e.getSectionCount();n+=1)for(let r=0;r<e.getParagraphCount(n);r+=1)for(let i=0;i<${TABLE_CONTROLS_PER_PARA};i+=1){let a;try{a=e.getTableDimensions(n,r,i);if(typeof a==\`string\`)a=JSON.parse(a)}catch{continue}if(!a||!a.rowCount)continue;let o=[],s=Number(a.cellCount||0);for(let c=0;c<s;c+=1){try{let l=e.getCellInfo(n,r,i,c);if(typeof l==\`string\`)l=JSON.parse(l);let u=e.getCellParagraphCount(n,r,i,c),d=[];for(let f=0;f<u;f+=1){let p=e.getCellParagraphLength(n,r,i,c,f);d.push(p>0?e.getTextInCell(n,r,i,c,f,0,p):\`\`)}o.push({index:c,row:l.row,col:l.col,text:d.join(\`\\n\`)})}catch{}}t.push({section:n,paragraph:r,control:i,rows:a.rowCount,cols:a.colCount,cells:o})}return t}${replaceCellMethod()}${insertAgentMethod()}`
}

/** A missing `}` here leaves listBodyParagraphs inside prepareTextCommand — blank Hangul page. */
const UNCLOSED_PREPARE_RE =
  /selectionEnd:i\}\}listBodyParagraphs\(\)\{this\.syncGeneration\(\)/

function closePrepareTextCommand(js) {
  return js.replace(
    UNCLOSED_PREPARE_RE,
    'selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration()',
  )
}

/** `a&&a.x??y` is a SyntaxError — blanks the Hangul iframe and blocks AI send. */
function repairIllegalNullishMix(js) {
  return js
    .replace(/Number\(a&&a\.paraIdx\?\?t\)/g, 'Number(a?.paraIdx??t)')
    .replace(/Number\(a&&a\.controlIdx\?\?0\)/g, 'Number(a?.controlIdx??0)')
}

/**
 * stripClassMethodCommas used to also eat the object-literal comma before the
 * insertFilled handler (`}async insertFilled…{if(await`). That is
 * `Unexpected token 'async'` — blank Hangul page, AI never becomes ready.
 */
const STRIPPED_FILLED_HANDLER_RE =
  /\}async insertFilledParagraphs\(e,t,n\)\{if\(await /g

function repairStrippedFilledHandlerComma(js) {
  return js.replace(
    STRIPPED_FILLED_HANDLER_RE,
    '},async insertFilledParagraphs(e,t,n){if(await ',
  )
}

/** WASM apply*Format encodes a JSON string; a raw object traps as OOB. */
function repairFormatJsonStringify(js) {
  return js
    .replace(
      /this\.deps\.wasm\.applyCharFormat\(e,t,n,r,a\)/g,
      'this.deps.wasm.applyCharFormat(e,t,n,r,JSON.stringify(a))',
    )
    .replace(
      /this\.deps\.wasm\.applyParaFormat\(e,t,r\)/g,
      'this.deps.wasm.applyParaFormat(e,t,JSON.stringify(r))',
    )
}

const DEFERRED_REPLACE_CELL_RE =
  /replaceCell\(e,t,n,r,i\)\{this\.syncGeneration\(\);let a=this\.deps\.wasm,o=a\.getCellParagraphLength\(e,t,n,r,0\),s=a\.replaceTextInCellDeferredPagination\(e,t,n,r,0,0,o,String\(i\?\?\`\`\)\);if\(typeof s==`string`\)try\{s=JSON\.parse\(s\)\}catch\{\}return s\}/

function repairDeferredCellWrite(js) {
  return js.replace(DEFERRED_REPLACE_CELL_RE, replaceCellMethod())
}

function repairTableInsertOffset(js) {
  if (js.includes('createTable(e,t,o,i,s)')) return js
  return js.replace(
    /let a=this\.deps\.wasm\.createTable\(e,t,0,i,s\)/,
    'let o=0;try{o=Number(this.deps.wasm.getParagraphLength(e,t))||0}catch{o=0}let a=this.deps.wasm.createTable(e,t,o,i,s)',
  )
}

const BROKEN_INSERT_HANDLER_RE =
  /async insertBodyParagraphs\(e\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.insertBodyParagraphs\(e\.section,e\.index,e\.count\)\}/

const INSERT_CLASS_METHOD_RE =
  /insertBodyParagraphs\(e,t,n\)\{this\.syncGeneration\(\);[\s\S]*?return\{section:e,index:t,count:i\}\}/

function repairInsertWasmCall(js) {
  if (js.includes(INSERT_SPLIT_NEEDLE)) return js
  if (!INSERT_CLASS_METHOD_RE.test(js)) return js
  return js.replace(INSERT_CLASS_METHOD_RE, insertAgentMethod())
}

function repairInsertHandler(js) {
  return js.replace(
    BROKEN_INSERT_HANDLER_RE,
    (_, ready, agent) =>
      `async insertBodyParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertBodyParagraphs(e,t,n)}`,
  )
}

function insertHandler(ready, agent) {
  return `async insertBodyParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertBodyParagraphs(e,t,n)},async insertFilledParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertFilledParagraphs(e,t,n)},${formatHandler(ready, agent)}`
}

function formatHandler(ready, agent) {
  return `async insertTable(e,t,n,r){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertTable(e,t,n,r)},async applyBodyCharFormat(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.applyBodyCharFormat(e,t,n,r,i)},async applyBodyParaFormat(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.applyBodyParaFormat(e,t,n)}`
}

function prepareSurfaceComplete(js) {
  return (
    js.includes(PREPARE_TEXT_V7_MARK) &&
    js.includes('findOrCreateFontId(String(a.fontName))') &&
    js.includes(INSERT_SPLIT_NEEDLE) &&
    js.includes('async insertFilledParagraphs(e,t,n){') &&
    js.includes('case`insertFilledParagraphs`') &&
    js.includes('case`insertTable`') &&
    js.includes('case`applyBodyCharFormat`') &&
    js.includes('insertTable(e,t,n,r){') &&
    js.includes('Number(a?.paraIdx??t)') &&
    !js.includes('a&&a.paraIdx??t') &&
    js.includes('applyCharFormat(e,t,n,r,JSON.stringify(a))') &&
    js.includes('applyParaFormat(e,t,JSON.stringify(r))') &&
    js.includes('insertTextInCell(e,t,n,r,0,0,x)') &&
    js.includes('createTable(e,t,o,i,s)') &&
    !js.includes('replaceTextInCellDeferredPagination(e,t,n,r,0,0,o,String') &&
    js.includes('r.splitParagraph(e,') &&
    !js.includes('r.insertParagraph(e,t+a)') &&
    !js.includes('insertBodyParagraphs(e.section,e.index,e.count)') &&
    js.includes('},async insertFilledParagraphs(e,t,n){if(await') &&
    !js.includes('}async insertFilledParagraphs(e,t,n){if(await') &&
    !/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/.test(js) &&
    !/,insertBodyParagraphs\(e,t,n\)\{this\.syncGeneration\(\)/.test(js) &&
    !/,insertTable\(e,t,n,r\)\{/.test(js) &&
    !UNCLOSED_PREPARE_RE.test(js)
  )
}

function prepareAgentMethods(snap) {
  return `${PREPARE_TEXT_V7_MARK}prepareTextCommand(){let e=this.getSelectionContext(),n=this.deps.input.getSelection(),r=null,i=null;if(e.target&&n&&n.start&&n.end&&n.start.sectionIndex===e.target.section&&n.start.paragraphIndex===e.target.paragraph&&n.end.sectionIndex===e.target.section&&n.end.paragraphIndex===e.target.paragraph&&n.end.charOffset>n.start.charOffset){r=n.start.charOffset,i=n.end.charOffset}if(!e.editable||!e.target)return{editable:!1,reason:\`not_editable\`,target:e.target,text:null,textSha256:null,formatSha256:null,adjacentContextSha256:null,selectionStart:r,selectionEnd:i};try{let t=${snap}(this.deps.wasm,e.target);return{editable:!0,reason:null,target:e.target,text:t.text,textSha256:t.textSha256,formatSha256:t.formatSha256,adjacentContextSha256:t.adjacentContextSha256,selectionStart:r,selectionEnd:i}}catch(a){return{editable:!1,reason:String(a&&a.message||a),target:e.target,text:null,textSha256:null,formatSha256:null,adjacentContextSha256:null,selectionStart:r,selectionEnd:i}}}listBodyParagraphs(){this.syncGeneration();let e=this.deps.wasm,t=[];for(let n=0;n<e.getSectionCount();n+=1)for(let r=0;r<e.getParagraphCount(n);r+=1){let i=e.getParagraphLength(n,r),a={kind:\`body_paragraph\`,section:n,paragraph:r,charOffset:0,length:i};try{let o=${snap}(e,a);t.push({editable:!0,reason:null,target:a,text:o.text,textSha256:o.textSha256,formatSha256:o.formatSha256,adjacentContextSha256:o.adjacentContextSha256})}catch(s){t.push({editable:!1,reason:String(s&&s.message||s),target:a,text:null,textSha256:null,formatSha256:null,adjacentContextSha256:null})}}return t}listFields(){this.syncGeneration();let e=this.deps.wasm.getFieldList();if(typeof e==\`string\`)try{e=JSON.parse(e)}catch{e=[]}if(!Array.isArray(e))return[];return e.map(t=>{let n=t&&(t.name||t.fieldName||t.fieldId||t.id)||\`\`,r=\`\`;if(n)try{let i=this.deps.wasm.getFieldValueByName(n);r=typeof i==\`string\`?i:i==null?\`\`:String(i)}catch{}return{name:n,value:r,type:t&&t.fieldType||null}}).filter(t=>t.name)}setField(e,t){this.syncGeneration();return this.deps.wasm.setFieldValueByName(String(e??\`\`),String(t??\`\`))}${tableAgentMethods()}`
}

function prepareAgentHandlers(ready, agent) {
  return `async getSelectionContext(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.getSelectionContext()},async prepareTextCommand(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.prepareTextCommand()},async listBodyParagraphs(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listBodyParagraphs()},async listFields(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listFields()},async setField(e,t){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.setField(e,t)},async listTables(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listTables()},async replaceCell(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.replaceCell(e,t,n,r,i)},${insertHandler(ready, agent)},async applyTextCommand(`
}

function prepareAgentRoutes(guard, params, host) {
  return `case\`getSelectionContext\`:return ${guard}(${params},\`getSelectionContext params\`),${host}.getSelectionContext();case\`prepareTextCommand\`:return ${host}.prepareTextCommand();case\`listBodyParagraphs\`:return ${host}.listBodyParagraphs();case\`listFields\`:return ${host}.listFields();case\`setField\`:return ${host}.setField(${params}.name,${params}.value);case\`listTables\`:return ${host}.listTables();case\`replaceCell\`:return ${host}.replaceCell(${params}.section,${params}.paragraph,${params}.control,${params}.cellIndex,${params}.text);case\`insertBodyParagraphs\`:return ${host}.insertBodyParagraphs(${params}.section,${params}.index,${params}.count);case\`insertFilledParagraphs\`:return ${host}.insertFilledParagraphs(${params}.section,${params}.index,${params}.texts);case\`insertTable\`:return ${host}.insertTable(${params}.section,${params}.index,${params}.rows,${params}.cols);case\`applyBodyCharFormat\`:return ${host}.applyBodyCharFormat(${params}.section,${params}.paragraph,${params}.start,${params}.end,${params}.format);case\`applyBodyParaFormat\`:return ${host}.applyBodyParaFormat(${params}.section,${params}.paragraph,${params}.format);case\`applyTextCommand\`:`
}

function stripClassMethodCommas(js) {
  return js
    .replace(/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/g, 'listBodyParagraphs(){this.syncGeneration()')
    .replace(/,listFields\(\)\{this\.syncGeneration\(\)/g, 'listFields(){this.syncGeneration()')
    .replace(/,setField\(e,t\)\{this\.syncGeneration\(\)/g, 'setField(e,t){this.syncGeneration()')
    .replace(/,listTables\(\)\{this\.syncGeneration\(\)/g, 'listTables(){this.syncGeneration()')
    .replace(/,replaceCell\(e,t,n,r,i\)\{this\.syncGeneration\(\)/g, 'replaceCell(e,t,n,r,i){this.syncGeneration()')
    .replace(/,insertBodyParagraphs\(e,t,n\)\{this\.syncGeneration\(\)/g, 'insertBodyParagraphs(e,t,n){this.syncGeneration()')
    .replace(
      /,async insertFilledParagraphs\(e,t,n\)\{if\(!Array\.isArray/g,
      'async insertFilledParagraphs(e,t,n){if(!Array.isArray',
    )
    .replace(/,insertTable\(e,t,n,r\)\{/g, 'insertTable(e,t,n,r){')
    .replace(/,applyBodyCharFormat\(e,t,n,r,i\)\{/g, 'applyBodyCharFormat(e,t,n,r,i){')
    .replace(/,applyBodyParaFormat\(e,t,n\)\{/g, 'applyBodyParaFormat(e,t,n){')
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

const REPLACE_CELL_CLASS_RE =
  /replaceCell\(e,t,n,r,i\)\{this\.syncGeneration\(\);[\s\S]*?return s\}async applyTextCommand/
const REPLACE_CELL_HANDLER_RE =
  /async replaceCell\(e,t,n,r,i\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.replaceCell\(e,t,n,r,i\)\},async applyTextCommand\(/
const REPLACE_CELL_ROUTE_RE =
  /case`replaceCell`:return ([A-Za-z_$][\w$]*)\.replaceCell\(([A-Za-z_$][\w$]*)\.section,\2\.paragraph,\2\.control,\2\.cellIndex,\2\.text\);case`applyTextCommand`:/

function attachInsertSurface(js) {
  let next = stripClassMethodCommas(js.replace(PREPARE_TEXT_V3_MARK, PREPARE_TEXT_V4_MARK))
  next = next.replace(
    REPLACE_CELL_CLASS_RE,
    `${replaceCellMethod()}${insertBodyMethod()}async applyTextCommand`,
  )
  next = next.replace(REPLACE_CELL_HANDLER_RE, (_, ready, agent) => {
    return `async replaceCell(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.replaceCell(e,t,n,r,i)},async insertBodyParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertBodyParagraphs(e,t,n)},async applyTextCommand(`
  })
  next = next.replace(
    REPLACE_CELL_ROUTE_RE,
    'case`replaceCell`:return $1.replaceCell($2.section,$2.paragraph,$2.control,$2.cellIndex,$2.text);case`insertBodyParagraphs`:return $1.insertBodyParagraphs($2.section,$2.index,$2.count);case`applyTextCommand`:',
  )
  return next
}

const INSERT_BODY_CLASS_END_RE =
  /insertBodyParagraphs\(e,t,n\)\{this\.syncGeneration\(\);[\s\S]*?return\{section:e,index:t,count:i\}\}async applyTextCommand/
const INSERT_BODY_HANDLER_END_RE =
  /async insertBodyParagraphs\(e,t,n\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.insertBodyParagraphs\(e,t,n\)\},async applyTextCommand\(/
const INSERT_BODY_ROUTE_END_RE =
  /case`insertBodyParagraphs`:return ([A-Za-z_$][\w$]*)\.insertBodyParagraphs\(([A-Za-z_$][\w$]*)\.section,\2\.index,\2\.count\);case`applyTextCommand`:/

function attachFillSurface(js) {
  let next = stripClassMethodCommas(js.replace(PREPARE_TEXT_V4_MARK, PREPARE_TEXT_V5_MARK))
  next = next.replace(INSERT_BODY_CLASS_END_RE, `${insertBodyMethod()}${insertFilledMethod()}async applyTextCommand`)
  next = next.replace(INSERT_BODY_HANDLER_END_RE, (_, ready, agent) => {
    return `async insertBodyParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertBodyParagraphs(e,t,n)},async insertFilledParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertFilledParagraphs(e,t,n)},async applyTextCommand(`
  })
  next = next.replace(
    INSERT_BODY_ROUTE_END_RE,
    'case`insertBodyParagraphs`:return $1.insertBodyParagraphs($2.section,$2.index,$2.count);case`insertFilledParagraphs`:return $1.insertFilledParagraphs($2.section,$2.index,$2.texts);case`applyTextCommand`:',
  )
  return next
}

const INSERT_FILLED_CLASS_END_RE =
  /async insertFilledParagraphs\(e,t,n\)\{[\s\S]*?return\{section:e,index:t,count:n\.length\}\}async applyTextCommand/
const INSERT_FILLED_HANDLER_END_RE =
  /async insertFilledParagraphs\(e,t,n\)\{if\(await ([A-Za-z_$][\w$]*),!([A-Za-z_$][\w$]*)\)throw Error\(`Document agent is not initialized`\);return \2\.insertFilledParagraphs\(e,t,n\)\},async applyTextCommand\(/
const INSERT_FILLED_ROUTE_END_RE =
  /case`insertFilledParagraphs`:return ([A-Za-z_$][\w$]*)\.insertFilledParagraphs\(([A-Za-z_$][\w$]*)\.section,\2\.index,\2\.texts\);case`applyTextCommand`:/

function attachFormatSurface(js) {
  let next = stripClassMethodCommas(js.replace(PREPARE_TEXT_V5_MARK, PREPARE_TEXT_V6_MARK))
  next = next.replace(
    INSERT_FILLED_CLASS_END_RE,
    `${insertFilledMethod()}${formatAgentMethods()}async applyTextCommand`,
  )
  next = next.replace(INSERT_FILLED_HANDLER_END_RE, (_, ready, agent) => {
    return `async insertFilledParagraphs(e,t,n){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.insertFilledParagraphs(e,t,n)},${formatHandler(ready, agent)},async applyTextCommand(`
  })
  next = next.replace(
    INSERT_FILLED_ROUTE_END_RE,
    'case`insertFilledParagraphs`:return $1.insertFilledParagraphs($2.section,$2.index,$2.texts);case`insertTable`:return $1.insertTable($2.section,$2.index,$2.rows,$2.cols);case`applyBodyCharFormat`:return $1.applyBodyCharFormat($2.section,$2.paragraph,$2.start,$2.end,$2.format);case`applyBodyParaFormat`:return $1.applyBodyParaFormat($2.section,$2.paragraph,$2.format);case`applyTextCommand`:',
  )
  return next
}

const CHAR_FORMAT_V6_RE =
  /applyBodyCharFormat\(e,t,n,r,i\)\{this\.syncGeneration\(\);let a=this\.deps\.wasm\.applyCharFormat\(e,t,n,r,i\);if\(typeof a==`string`\)try\{a=JSON\.parse\(a\)\}catch\{\}if\(a&&a\.ok===!1\)throw Error\(String\(a\.error\|\|a\.message\|\|`applyCharFormat failed`\)\);return a\}/

function attachFontSurface(js) {
  let next = stripClassMethodCommas(js.replace(PREPARE_TEXT_V6_MARK, PREPARE_TEXT_V7_MARK))
  next = next.replace(CHAR_FORMAT_V6_RE, charFormatMethod())
  return next
}

export function exposePrepareTextCommand(js) {
  const closed = repairTableInsertOffset(
    repairDeferredCellWrite(
      repairFormatJsonStringify(
        repairStrippedFilledHandlerComma(
          repairIllegalNullishMix(
            closePrepareTextCommand(repairInsertHandler(repairInsertWasmCall(js))),
          ),
        ),
      ),
    ),
  )
  if (prepareSurfaceComplete(closed)) return closed
  if (
    closed.includes(PREPARE_TEXT_V7_MARK) ||
    closed.includes(PREPARE_TEXT_V6_MARK) ||
    closed.includes(PREPARE_TEXT_V5_MARK) ||
    closed.includes(PREPARE_TEXT_V4_MARK) ||
    closed.includes(PREPARE_TEXT_V3_MARK) ||
    closed.includes(PREPARE_TEXT_V2_MARK)
  ) {
    const upgraded = repairTableInsertOffset(
      repairDeferredCellWrite(
        repairFormatJsonStringify(
          repairStrippedFilledHandlerComma(
            repairIllegalNullishMix(
              attachFontSurface(
                attachFormatSurface(attachFillSurface(attachInsertSurface(attachTableSurface(closed)))),
              ),
            ),
          ),
        ),
      ),
    )
    if (
      !upgraded.includes(PREPARE_TEXT_V7_MARK) ||
      !upgraded.includes('findOrCreateFontId(String(a.fontName))') ||
      !upgraded.includes('applyCharFormat(e,t,n,r,JSON.stringify(a))') ||
      upgraded.includes('a&&a.paraIdx??t') ||
      upgraded.includes('}async insertFilledParagraphs(e,t,n){if(await')
    ) {
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
        `async prepareTextCommand(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.prepareTextCommand()},async listBodyParagraphs(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listBodyParagraphs()},async listFields(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listFields()},async setField(e,t){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.setField(e,t)},async listTables(){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.listTables()},async replaceCell(e,t,n,r,i){if(await ${ready},!${agent})throw Error(\`Document agent is not initialized\`);return ${agent}.replaceCell(e,t,n,r,i)},${insertHandler(ready, agent)},async applyTextCommand(`,
    )
  }
  next = next.replace(PREPARE_ROUTE_RE, (_, guard, params, host) => prepareAgentRoutes(guard, params, host))
  if (PREPARE_V1_ROUTE_RE.test(next) && !next.includes('case`listBodyParagraphs`')) {
    next = next.replace(PREPARE_V1_ROUTE_RE, (_, guard, params, host) => prepareAgentRoutes(guard, params, host))
  }
  if (
    !next.includes(PREPARE_TEXT_V7_MARK) ||
    !next.includes('findOrCreateFontId(String(a.fontName))') ||
    !next.includes('case`listTables`') ||
    !next.includes('case`setField`') ||
    !next.includes('case`insertBodyParagraphs`') ||
    !next.includes('case`insertFilledParagraphs`') ||
    !next.includes('case`insertTable`') ||
    !next.includes('Number(a?.paraIdx??t)') ||
    !next.includes('applyCharFormat(e,t,n,r,JSON.stringify(a))') ||
    !next.includes('insertTextInCell(e,t,n,r,0,0,x)') ||
    !next.includes('createTable(e,t,o,i,s)') ||
    next.includes('a&&a.paraIdx??t') ||
    next.includes('}async insertFilledParagraphs(e,t,n){if(await')
  ) {
    throw new Error('rhwp-studio prepareTextCommand surface changed — update exposePrepareTextCommand()')
  }
  return next
}

export function hasPrepareTextCommand(js) {
  return (
    js.includes(PREPARE_TEXT_V7_MARK) ||
    js.includes(PREPARE_TEXT_V6_MARK) ||
    js.includes(PREPARE_TEXT_V5_MARK) ||
    js.includes(PREPARE_TEXT_V4_MARK) ||
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
