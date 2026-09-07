# GenOffice Web 版本实现指南

## 研究结论

### 1. Web File System Access API

**推荐方案**: 使用原生 API + `file-system-access` 降级库

```typescript
// 核心 API
const handle = await window.showOpenFilePicker({
  types: [{
    description: 'Documents',
    accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
  }]
});
const file = await handle.getFile();
const content = await file.arrayBuffer();
```

**降级方案**: File Input + Blob 下载
```typescript
// 对于不支持 File System Access API 的浏览器
<input type="file" accept=".docx" onchange="handleFile(e)">
// 使用 file-saver 下载
import { saveAs } from 'file-saver';
saveAs(blob, 'document.docx');
```

### 2. AI 服务集成

**推荐方案**: HTTP API + SSE 流式响应

```typescript
// 流式 AI 响应
async function* streamAI(prompt: string) {
  const response = await fetch('/api/ai/stream', {
    method: 'POST',
    body: JSON.stringify({ prompt }),
    headers: { 'Content-Type': 'application/json' }
  });
  
  const reader = response.body?.getReader();
  const decoder = new TextDecoder();
  
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    yield decoder.decode(value);
  }
}
```

**库推荐**:
- `@vercel/ai` - 流式响应处理
- `ai` (LangChain) - AI 应用框架

### 3. 协作文档编辑

**推荐方案**: Yjs + WebRTC/WebSocket

```typescript
import * as Y from 'yjs';
import { WebrtcProvider } from 'y-webrtc';

// 创建文档
const ydoc = new Y.Doc();
const provider = new WebrtcProvider('room-id', ydoc);

// 获取文本
const ytext = ydoc.getText('content');
```

**替代方案**:
- `TipTap` - 基于 ProseMirror 的富文本编辑器
- `Lexical` - Meta 出品的编辑器框架

### 4. 离线支持

**推荐方案**: Workbox + IndexedDB

```typescript
// sw.js
import { precacheAndRoute } from 'workbox-precaching';
precacheAndRoute(self.__WB_MANIFEST);

// 缓存策略
import { StaleWhileRevalidate } from 'workbox-strategies';
registerRoute(
  ({ url }) => url.origin === 'https://api.example.com',
  new StaleWhileRevalidate({ cacheName: 'api-cache' })
);
```

### 5. 打印功能

**推荐方案**: `print-js` + 浏览器原生打印

```typescript
import print from 'print-js';

// 打印 HTML 内容
print({
  printable: elementId,
  type: 'html',
  documentTitle: 'My Document'
});

// 导出 PDF
import jsPDF from 'jspdf';
```

## 技术栈选择

| 功能 | 推荐方案 | 备选方案 |
|------|----------|----------|
| 文件选择 | File System Access API | `<input type="file">` |
| 文件保存 | File System Access API | `file-saver` |
| AI 对话 | Server-Sent Events | WebSocket |
| 协作编辑 | Yjs + y-webrtc | TipTap + Hocuspocus |
| 离线缓存 | Workbox | Service Worker |
| 打印 | print-js | 浏览器原生 |

## 实现优先级

### Phase 1: 核心功能
1. ✅ Web Server IPC 基础 (已完成 36.4%)
2. 文件打开/保存
3. AI 聊天接口

### Phase 2: 增强功能
4. 项目管理完整功能
5. 文件上传/下载
6. AI 流式对话

### Phase 3: 高级功能
7. 协作文档编辑
8. 离线支持
9. 打印/导出

## 依赖建议

```json
{
  "dependencies": {
    "file-saver": "^2.0.5",
    "yjs": "^13.6.0",
    "y-webrtc": "^10.2.5",
    "@tiptap/core": "^2.0.0",
    "print-js": "^1.6.0",
    "workbox-precaching": "^7.0.0"
  }
}
```
