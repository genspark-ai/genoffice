# Capability: web-launch(真实启动 web 版本验证)

## 1. 概述

genoffice 的 web 版本可被真实启动:执行文档中的启动命令后,Electron 主进程(带 HTTP 桥)与 Vite 开发服务器运行,真实浏览器可打开 web 地址并完整使用核心流程。

## 2. 真实启动

- `npm run dev -w @genoffice/docs` 启动 docs 的 Electron 主进程(桥端口 5273)与 Vite(5173);`npm run dev -w @genoffice/markdown` 同理(桥 5277 / Vite 5177)。
- 真实浏览器(Chromium/Chrome)打开 `http://localhost:<dev端口>/`,页面完整渲染,编辑器可见。
- web 端 `window.*` API 由浏览器 bootstrap 以 HTTP 传输构造,调用经 `POST /api/ipc/:channel` 与 SSE。

## 3. 核心流程

- markdown:编辑→保存→重开,内容持久化,全程经 HTTP 桥。
- docs:新建→编辑→保存→重开,内容持久化,全程经 HTTP 桥。

## 4. 降级

- web 端调用原生专属通道(如原生打开对话框)收到结构化"仅桌面版支持"错误(`WEB_UNSUPPORTED` 语义)。

## 5. 文档

- 启动命令与真实启动验证说明已更新,命令实测可跑。
