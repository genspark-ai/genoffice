# GenOffice Web Server

> Standalone Web Server for GenOffice - no Electron required

## 特性

- **338+ IPC 通道** - 与 Electron 版本完全兼容
- **MiniMax AI 集成** - 开箱即用的 AI 能力
- **SSE 流式响应** - 实时 AI 对话
- **跨平台** - Linux / macOS / Windows
- **容器化部署** - Docker 支持

## 快速开始

### 方式一: 直接运行

```bash
# 安装依赖
npm install

# 构建
npm run build

# 运行
npm start
# 或指定端口
PORT=3000 npm start
```

### 方式二: Docker

```bash
# 构建镜像
docker build -t genoffice-web -f Dockerfile ..

# 运行容器
docker run -p 8080:8080 genoffice-web

# 或使用 docker-compose
docker-compose up -d
```

### 方式三: 使用 MiniMax AI

```bash
# 设置 API Key
MINIMAX_API_KEY=your-api-key npm start
```

## API 端点

### 健康检查

```bash
curl http://localhost:8080/health
```

### IPC 通道调用

```bash
# 调用 AI 对话
curl -X POST http://localhost:8080/api/ipc/ai:chat \
  -H "Content-Type: application/json" \
  -d '{"args":[{"message":"Hello"}]}'

# 打开文档
curl -X POST http://localhost:8080/api/ipc/docs:open \
  -H "Content-Type: application/json" \
  -d '{"args":[]}'

# 获取主题设置
curl -X POST http://localhost:8080/api/ipc/home:get-theme \
  -H "Content-Type: application/json" \
  -d '{"args":[]}'
```

### 通道列表

```bash
curl http://localhost:8080/api/channels
```

### AI 流式响应 (SSE)

```bash
curl -X POST http://localhost:8080/api/ai/stream \
  -H "Content-Type: application/json" \
  -d '{"messages":[{"role":"user","content":"Hello"}]}'
```

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `HOST` | `0.0.0.0` | 绑定地址 |
| `PORT` | `8080` | 端口 |
| `MINIMAX_API_KEY` | - | MiniMax API Key |
| `DATA_DIR` | `/tmp/genoffice-data` | 数据存储目录 |

## 通道统计

| 类别 | 数量 |
|------|------|
| AI | 41 |
| Slides | 126 |
| Docs | 15 |
| Home | 32 |
| Tabs | 8 |
| Update | 5 |
| PDF | 4 |
| Files | 7 |
| 其他 | 100+ |
| **总计** | **338** |

## 开发

```bash
# 开发模式 (热重载)
npm run dev

# 构建
npm run build

# 类型检查
npm run typecheck
```

## 打包发布

```bash
# 安装 pkg
npm install -g pkg

# 打包所有平台
npm run pkg:all

# 单独打包
npm run pkg:linux
npm run pkg:win
npm run pkg:mac
```

## 许可

Apache-2.0
