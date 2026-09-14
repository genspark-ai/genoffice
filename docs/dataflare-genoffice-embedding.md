# Dataflare Work 集成 GenOffice AI 文档

## 推荐架构

采用“Dataflare 宿主 + GenOffice Web 微前端 + Dataflare AI Gateway”的三层架构：

```text
Dataflare Vue 宿主
  ├─ 登录、租户、权限、客户/订单/邮件上下文
  ├─ OfficeWorkspaceView（iframe 容器）
  └─ host ↔ editor postMessage 协议
       ↓ 同源反向代理 /office-engine
GenOffice Web Server + Docs/Sheets/Slides Renderer
  ├─ Office 文档解析、渲染、编辑、撤销、格式保真
  ├─ AI 面板与翻译预览
  └─ Translation Adapter → Dataflare Translation Gateway
       ↓
Dataflare Work Backend
  ├─ 多租户 Provider/模型配置
  ├─ 翻译记忆、术语库、知识库/RAG
  ├─ 任务、计费、审计、质量校验
  └─ CRM/邮件/订单/客户业务工具
```

不要把 Electron preload、Node IPC 或 ProseMirror 编辑状态直接嵌入 Vue。Electron 是桌面壳；Dataflare 的集成边界应是 GenOffice Web Server。

## 为什么先使用 iframe

1. 隔离 React/TipTap/文档 CSS，避免与 Vue、Element Plus、Tailwind 冲突。
2. 复用 GenOffice 已有 Web Server、IPC bridge 和多应用路由，不重复实现编辑器。
3. 先统一鉴权、保存、AI 和事件协议，后续再切换到 Module Federation 或 Web Component。
4. 编辑器崩溃时只重载 iframe，不影响 Dataflare 工作台。

生产环境建议同源部署：

```text
https://work.example.com/              → Dataflare frontend/backend
https://work.example.com/office-engine → GenOffice web-server
```

不要把 `Manager-Token`、API Key 或长期 JWT 放进 iframe URL/query string。认证应由 Dataflare 会话和反向代理完成。

## 已落地桥接协议

协议名：`genoffice-dataflare/v1`。

Dataflare → GenOffice：`init`（租户、用户、文档、业务对象、语言、主题、只读状态）、`open-document`、`set-readonly`、`focus-ai`、`dispose`。

GenOffice → Dataflare：`ready`、`document-dirty`、`document-saved`、`ai-progress`、`error`。

消息必须检查 `event.source`、`event.origin` 和协议字段。生产环境应将父源改为固定白名单，不接受任意 query 参数作为可信来源。

## 文档数据与保存

iframe 只负责 Office 内容编辑，不直接访问 CRM 数据库。正式版本增加 `DocumentAdapter`：

```text
Dataflare docs_document / crm_knowledge
       ↓ GET /api/v1/office/documents/{id}/content
Host bridge init/open-document
       ↓
GenOffice loads bytes/content into editor
       ↓ save command + revision + optimistic lock
Dataflare persists and audits the revision
```

第一阶段可以用临时文件/下载上传验证交互；正式版本必须使用 `documentId + revision + idempotencyKey`，避免 iframe 重载或协作编辑覆盖数据。

## AI 翻译集成

编辑器发送结构化 `TranslationRequest`，不直接读取 Dataflare API Key：

1. GenOffice 提取 paragraph/run/cell/shape 单元。
2. Dataflare Gateway 校验租户权限和幂等键。
3. Gateway 检索术语库和翻译记忆。
4. Koog/Provider 批量翻译并通过 SSE v2 发出 `unit_started/unit_completed`。
5. Gateway 做数字、术语、格式安全校验。
6. GenOffice 展示预览，用户确认后按稳定 `unitId/path` 应用。
7. 用户确认后再写入翻译记忆。

这样能复用 Docs、Sheets、Slides、PDF 的编辑器能力，同时避免 Dataflare 后端依赖 ProseMirror 或 PPTX 内部结构。

## 分阶段路线

### Phase 1：Web iframe MVP

- `/office` 宿主路由。
- 同源反向代理。
- `init/focus-ai/set-readonly` 桥接。
- Dataflare Provider/权限治理。
- Docs 选区翻译和 AI 面板。

### Phase 2：云文档适配

- `DocumentAdapter` 和 `docs_document` revision API。
- 云文档打开、保存、自动保存、冲突提示。
- 业务对象上下文注入 CRM/邮件/订单 Agent。

### Phase 3：企业 AI 文档

- 结构化翻译任务、术语库、TM、质量校验。
- 文档抽取、摘要、字段识别、合同风险、邮件生成。
- 审计、计费、敏感信息脱敏、异步任务恢复。

### Phase 4：直挂载优化

协议、CSS 隔离、保存 API 稳定后，再考虑 Module Federation 或 React Web Component。直挂载减少 iframe 层级，但会增加版本、样式、运行时和故障隔离成本。

## 部署检查

- 配置 `VITE_GENOFFICE_URL=/office-engine`。
- 反向代理 `/office-engine/*` 到 GenOffice 静态资源与 API。
- 只允许 Dataflare 实际 origin，禁止生产 `Access-Control-Allow-Origin: *`。
- Cookie 使用 `Secure`、`HttpOnly`、`SameSite=Lax/Strict`。
- 开启 CSP：`frame-src 'self'`，并将 `frame-ancestors` 限制为 Dataflare origin。
- iframe 仅开启 `clipboard-read; clipboard-write`，不授予 camera/microphone/geolocation。
