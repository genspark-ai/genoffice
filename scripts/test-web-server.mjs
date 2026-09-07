/**
 * GenOffice Web Server 功能测试脚本
 * 
 * 对标 Electron IPC 功能，验证 Web Server 的功能完整性
 * 
 * 运行方式:
 *   node scripts/test-web-server.mjs
 */

import { writeFileSync } from 'node:fs'

const BASE_URL = process.env.WEB_SERVER_URL || 'http://localhost:8080'

// Electron IPC 通道分类
const ELECTRON_CHANNELS = {
  // 基础应用通道
  app: [
    'app:get-language',
    'app:get-version',
    'app:get-platform',
    'app:get-theme',
  ],
  
  // AI 功能通道
  ai: [
    'ai:chat',
    'ai:get-settings',
    'ai:set-settings',
    'ai:stream',
    'ai:stream-cancel',
    'ai:web-search',
    'ai:image-search',
    'ai:gsk-login',
    'ai:log-run-failure',
  ],
  
  // 协作功能通道
  collab: [
    'collab:join',
    'collab:leave',
    'collab:sync',
  ],
  
  // 文档功能通道
  docs: [
    'docs:open',
    'docs:open-path',
    'docs:read-path',
    'docs:save-new',
    'docs:print',
    'docs:pick-image',
    'docs:font-metrics',
    'docs:recent',
    'docs:consume-new-blank',
    'docs:consume-pending-open',
    'docs:consume-ai-doc-content',
    'docs:write-recovery',
    'docs:password-intent-revision',
  ],
  
  // 表格功能通道
  sheets: [
    'sheets:new-blank',
    'sheets:has-queued-workbook',
    'workbook:open-path',
  ],
  
  // 幻灯片功能通道
  slides: [
    'slides:new-blank',
    'slides:open',
    'slides:open-path',
    'slides:save',
    'slides:save-as',
    'slides:export-pdf',
    'slides:recent',
    'slides:consume-pending-open',
    'slides:add-blank-slide',
    'slides:add-slide',
    'slides:add-chart',
    'slides:add-image-bytes',
    'slides:add-table',
    'slides:add-text',
    'slides:add-element',
    'slides:edit-text',
    'slides:delete-element',
    'slides:undo',
    'slides:redo',
    'slides:get-render-slides',
  ],
  
  // PDF 功能通道
  pdf: [
    'pdf:open-path',
  ],
  
  // 项目功能通道
  project: [
    'project:list',
    'project:create',
    'project:files',
    'project:rename',
    'project:delete',
    'project:moveFile',
    'project:timeline',
  ],
  
  // 文件功能通道
  files: [
    'files:pick',
    'files:add',
    'files:read-image',
  ],
  
  // 窗口功能通道
  win: [
    'win:new',
    'win:list',
    'win:focus',
  ],
  
  // 剪贴板功能通道
  clipboard: [
    'clipboard:copy',
    'clipboard:cut',
    'clipboard:paste',
  ],
}

// Web Server 已实现的通道
const WEB_SERVER_IMPLEMENTED = [
  // App
  'app:get-language',
  'app:get-version',
  'app:get-platform',
  'app:get-theme',
  // AI
  'ai:chat',
  'ai:get-settings',
  'ai:set-settings',
  'ai:gsk-login',
  'ai:stream',
  'ai:stream-cancel',
  'ai:web-search',
  'ai:image-search',
  // Docs
  'docs:get-settings',
  'docs:save-settings',
  'docs:recent',
  'docs:font-metrics',
  'docs:pick-image',
  'docs:open',
  'docs:open-path',
  'docs:read-path',
  'docs:save-new',
  'docs:print',
  'docs:consume-new-blank',
  'docs:consume-pending-open',
  'docs:consume-ai-doc-content',
  'docs:write-recovery',
  'docs:password-intent-revision',
  // Sheets
  'sheets:new-blank',
  'sheets:has-queued-workbook',
  'sheets:consume-new-blank',
  'workbook:open-path',
  // Slides (基础)
  'slides:new-blank',
  'slides:recent',
  'slides:open',
  'slides:open-path',
  'slides:save',
  'slides:save-as',
  'slides:export-pdf',
  'slides:consume-pending-open',
  'slides:add-blank-slide',
  'slides:add-slide',
  'slides:add-chart',
  'slides:add-image-bytes',
  'slides:add-table',
  'slides:add-text',
  'slides:add-element',
  'slides:edit-text',
  'slides:delete-element',
  'slides:undo',
  'slides:redo',
  'slides:get-render-slides',
  // Slides (扩展)
  'slides:get-animations',
  'slides:get-chart-data',
  'slides:get-comments',
  'slides:get-header-footer',
  'slides:get-layouts',
  'slides:get-link',
  'slides:get-notes',
  'slides:get-sections',
  'slides:get-shape-keys',
  'slides:get-slide-links',
  'slides:get-slide-size',
  'slides:get-transition',
  'slides:apply-edit-script',
  'slides:apply-header-footer',
  'slides:apply-theme',
  'slides:apply-txn',
  'slides:batch-edit-transform',
  'slides:chart-color-schemes',
  'slides:clipboard-external',
  'slides:clipboard-probe',
  'slides:copy-elements',
  'slides:copy-slide',
  'slides:delete-comment',
  'slides:delete-slide',
  'slides:duplicate-elements',
  'slides:edit-background',
  'slides:edit-chart',
  'slides:edit-connector-endpoints',
  'slides:edit-fill',
  'slides:edit-image-fill',
  'slides:edit-picture-opacity',
  'slides:edit-picture-src-rect',
  'slides:edit-stroke',
  'slides:edit-table-cell',
  'slides:edit-table-style',
  'slides:edit-transform',
  'slides:find-replace',
  'slides:flip-elements',
  'slides:font-catalog',
  'slides:font-missing',
  'slides:get-run-links',
  'slides:group-elements',
  'slides:has-slide-clipboard',
  'slides:history-batch-begin',
  'slides:history-batch-end',
  'slides:insert-image',
  'slides:is-dirty',
  'slides:master-close',
  'slides:master-delete-element',
  'slides:master-edit-fill',
  'slides:master-edit-stroke',
  'slides:master-edit-text',
  'slides:master-edit-transform',
  'slides:master-enter',
  'slides:master-open',
  'slides:media-data',
  'slides:move-section',
  'slides:move-slide',
  'slides:native-clipboard',
  'slides:paste-elements',
  'slides:paste-slide',
  'slides:presenter-end',
  'slides:presenter-start',
  'slides:presenter-swap',
  'slides:remove-section',
  'slides:rename-section',
  'slides:reorder-element',
  'slides:replace-picture-bytes',
  'slides:set-advance-times',
  'slides:set-animations',
  'slides:set-element-font',
  'slides:set-element-paragraph-format',
  'slides:set-hidden',
  'slides:set-link',
  'slides:set-notes',
  'slides:set-sections',
  'slides:set-slide-layout',
  'slides:set-slide-size',
  'slides:set-table-cell-anchor',
  'slides:set-table-col-width',
  'slides:set-table-row-height',
  'slides:set-transition',
  'slides:show-fullscreen',
  'slides:table-merge',
  'slides:table-structure',
  'slides:ungroup-element',
  'slides:add-comment',
  'slides:add-ink',
  'slides:add-media-bytes',
  'slides:add-section',
  'slides:add-slide-with-layout',
  'slides:add-smartart',
  'slides:ai-snapshot-restore',
  'slides:audience-ready',
  'slides:cloud-gen-status',
  'slides:files-add',
  'slides:files-pick',
  'slides:files-read-image',
  'slides:pick-export-dir',
  'slides:pick-export-pdf-path',
  'slides:private-font-data',
  'slides:private-font-faces',
  'slides:repaste-slide',
  // PDF
  'pdf:open-path',
  // Project
  'project:list',
  'project:create',
  'project:files',
  'project:rename',
  'project:delete',
  'project:moveFile',
  'project:timeline',
  // Files
  'files:pick',
  'files:add',
  'files:read-image',
  // Collaboration
  'collab:join',
  'collab:leave',
  'collab:sync',
  // Clipboard (原生)
  'copy',
  'cut',
  'paste',
  'clipboard:copy',
  'clipboard:cut',
  'clipboard:paste',
  // Win
  'win:new',
  'win:list',
  'win:focus',
  // Markdown
  'md-asset',
  // Slides 字体
  'slides:font-download',
  'slides:font-install-local',
  'slides:insert-model3d',
  // AI
  'ai:log-run-failure',
  // AnyDoc
  'anydoc:get-config',
  'anydoc:set-config',
  'anydoc:recognize',
  'anydoc:convert',
  'anydoc:extract-text',
  'anydoc:extract-tables',
  'anydoc:extract-images',
  'anydoc:render-preview',
  // Web
  'web:write-temp-file',
  'web:read-file-bytes',
  'web:make-temp-dir',
  'web:save-file',
]

// 测试结果收集
const results = {
  passed: [],
  failed: [],
  skipped: [],
  unimplemented: [],
}

async function testChannel(channel) {
  // 根据通道类型构造参数
  let body = { args: [] }
  if (channel === 'project:create') {
    body = { args: [{ name: 'Test Project' }] }
  } else if (channel === 'project:files' || channel === 'project:rename' || channel === 'project:delete' || channel === 'project:timeline' || channel === 'project:moveFile') {
    body = { args: [{ id: 'test-id' }] }
  } else if (channel === 'docs:font-metrics') {
    body = { args: ['sans-serif'] }
  } else if (channel === 'ai:chat') {
    body = { args: [{ message: 'Hello' }] }
  } else if (channel === 'ai:stream') {
    body = { args: [{ message: 'Hello', sessionId: `test-${Date.now()}` }] }
  } else if (channel === 'files:add') {
    body = { args: [] }
  } else if (channel === 'collab:join' || channel === 'collab:leave' || channel === 'collab:sync') {
    body = { args: [{ docId: 'test-doc', userId: 'test-user' }] }
  } else if (channel === 'docs:open-path' || channel === 'docs:read-path') {
    // 创建临时测试文件
    const docxPath = '/tmp/genoffice-test.docx'
    writeFileSync(docxPath, Buffer.from('PK\x03\x04')) // 最小 DOCX 头
    body = { args: [docxPath] }
  } else if (channel === 'workbook:open-path') {
    const xlsxPath = '/tmp/genoffice-test.xlsx'
    writeFileSync(xlsxPath, Buffer.from('PK\x03\x04'))
    body = { args: [xlsxPath] }
  } else if (channel === 'slides:open-path') {
    const pptxPath = '/tmp/genoffice-test.pptx'
    writeFileSync(pptxPath, Buffer.from('PK\x03\x04'))
    body = { args: [pptxPath] }
  } else if (channel === 'pdf:open-path') {
    const pdfPath = '/tmp/genoffice-test.pdf'
    writeFileSync(pdfPath, Buffer.from('%PDF-1.4'))
    body = { args: [pdfPath] }
  } else if (channel === 'docs:save-new' || channel === 'slides:save' || channel === 'slides:save-as') {
    body = { args: [{ defaultName: 'test.docx' }] }
  }
  
  try {
    const response = await fetch(`${BASE_URL}/api/ipc/${channel}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })
    
    if (response.ok) {
      return { status: 'passed', data: await response.json() }
    } else if (response.status === 404) {
      return { status: 'unimplemented' }
    } else {
      return { status: 'failed', error: `HTTP ${response.status}` }
    }
  } catch (error) {
    return { status: 'failed', error: error.message }
  }
}

async function runTests() {
  console.log(`
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║   GenOffice Web Server 功能测试                               ║
║   对标 Electron IPC 功能                                       ║
║                                                               ║
║   Server: ${BASE_URL.padEnd(50)}║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
`)

  // 首先测试基础连接
  console.log('\n📡 基础连接测试...')
  try {
    const health = await fetch(`${BASE_URL}/health`)
    const healthData = await health.json()
    console.log(`   ✅ Health: ${JSON.stringify(healthData)}`)
  } catch (error) {
    console.log(`   ❌ 服务未运行: ${error.message}`)
    console.log('\n请先启动 Web Server: npm run start:web')
    process.exit(1)
  }

  // 按类别测试通道
  for (const [category, channels] of Object.entries(ELECTRON_CHANNELS)) {
    console.log(`\n📂 ${category.toUpperCase()} 功能测试`)
    console.log('─'.repeat(60))
    
    for (const channel of channels) {
      const implemented = WEB_SERVER_IMPLEMENTED.includes(channel)
      if (!implemented) {
        results.unimplemented.push({ category, channel })
        console.log(`   ⏭️  ${channel.padEnd(40)} (未实现)`)
        continue
      }
      
      const result = await testChannel(channel)
      
      if (result.status === 'passed') {
        results.passed.push({ category, channel, data: result.data })
        console.log(`   ✅ ${channel.padEnd(40)} → ${JSON.stringify(result.data).slice(0, 40)}`)
      } else if (result.status === 'unimplemented') {
        results.unimplemented.push({ category, channel })
        console.log(`   ⏭️  ${channel.padEnd(40)} (未实现)`)
      } else {
        results.failed.push({ category, channel, error: result.error })
        console.log(`   ❌ ${channel.padEnd(40)} → ${result.error}`)
      }
    }
  }

  // 输出总结
  console.log(`
╔═══════════════════════════════════════════════════════════════╗
║                       测试结果总结                             ║
╚═══════════════════════════════════════════════════════════════╝
`)

  console.log(`\n✅ 通过: ${results.passed.length}`)
  console.log(`❌ 失败: ${results.failed.length}`)
  console.log(`⏭️  未实现: ${results.unimplemented.length}`)

  if (results.unimplemented.length > 0) {
    console.log(`
📋 需要实现的通道 (按类别分组):

${Object.entries(
  results.unimplemented.reduce((acc, { category, channel }) => {
    if (!acc[category]) acc[category] = []
    acc[category].push(channel)
    return acc
  }, {})
).map(([cat, channels]) => `
【${cat}】
${channels.map(ch => `  - ${ch}`).join('\n')}
`).join('\n')}
`)

    // 计算实现度
    const totalElectron = Object.values(ELECTRON_CHANNELS).flat().length
    const implementedCount = totalElectron - results.unimplemented.length
    const percentage = ((implementedCount / totalElectron) * 100).toFixed(1)
    
    console.log(`\n📊 Electron 功能实现度: ${implementedCount}/${totalElectron} (${percentage}%)`)
  }

  if (results.failed.length > 0) {
    console.log(`
❌ 失败的测试:
${results.failed.map(({ channel, error }) => `  - ${channel}: ${error}`).join('\n')}
`)
  }

  // 详细报告
  if (results.passed.length > 0) {
    console.log(`
✅ 已实现的功能详情:
${results.passed.map(({ channel, data }) => `  - ${channel}: ${JSON.stringify(data)}`).join('\n')}
`)
  }

  return results
}

// 运行测试
runTests().then((results) => {
  const exitCode = results.failed.length > 0 ? 1 : 0
  process.exit(exitCode)
}).catch((error) => {
  console.error('测试失败:', error)
  process.exit(1)
})
