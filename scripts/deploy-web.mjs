/**
 * GenOffice Web 部署脚本
 * 
 * 功能：
 * 1. 构建所有应用的渲染器
 * 2. 将静态文件复制到统一目录
 * 3. 配置 Electron 主进程服务静态文件
 * 4. 创建启动脚本
 */

import { execSync } from 'node:child_process'
import { copyFileSync, cpSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

const ROOT = resolve(import.meta.dirname, '..')
const DEPLOY_DIR = join(ROOT, 'deploy')
const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'shell']

function log(msg) {
  console.log(`[deploy] ${msg}`)
}

function run(cmd, opts = {}) {
  log(`Running: ${cmd}`)
  execSync(cmd, { cwd: ROOT, stdio: 'inherit', ...opts })
}

// 1. 确保 Rust 环境
if (!existsSync(join(process.env.HOME || '/root', '.cargo/bin/cargo'))) {
  log('Rust not found. Installing...')
  run(`curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y`)
}

// 2. 安装依赖
log('Installing dependencies...')
run('npm install')

// 3. 构建所有应用
log('Building all apps...')
run('npm run build:all')

// 4. 创建部署目录
log('Creating deployment directory...')
mkdirSync(DEPLOY_DIR, { recursive: true })
mkdirSync(join(DEPLOY_DIR, 'static'), { recursive: true })

// 5. 复制静态文件
for (const app of APPS) {
  const srcDir = join(ROOT, 'apps', app, 'out', 'renderer')
  const destDir = join(DEPLOY_DIR, 'static', app)
  
  if (existsSync(srcDir)) {
    log(`Copying ${app} renderer to static...`)
    cpSync(srcDir, destDir, { recursive: true })
  }
}

// 6. 复制主进程文件
log('Copying main process...')
cpSync(join(ROOT, 'apps', 'shell', 'out', 'main'), join(DEPLOY_DIR, 'main'), { recursive: true })
cpSync(join(ROOT, 'apps', 'shell', 'out', 'preload'), join(DEPLOY_DIR, 'preload'), { recursive: true })

// 7. 创建启动脚本
const startScript = `#!/bin/bash
# GenOffice Web 启动脚本
# 
# 使用方式：
#   ./start.sh              # 前台运行
#   nohup ./start.sh &     # 后台运行

SCRIPT_DIR="$(cd "$(dirname "\${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

export NODE_ENV=production
export SHELL_IPC_PORT=5299
export SHELL_DEV_PORT=8080

# 复制模块文件
cp -r apps/docs/out/* . 2>/dev/null || true
cp -r apps/sheets/out/* . 2>/dev/null || true
cp -r apps/slides/out/* . 2>/dev/null || true
cp -r apps/pdf/out/* . 2>/dev/null || true
cp -r apps/markdown/out/* . 2>/dev/null || true

# 启动 Electron 主进程（带 HTTP IPC Bridge）
# 注意：Electron 需要图形界面
# 无头环境请使用 Xvfb: xvfb-run --auto-servernum node main/index.js
electron main/index.js
`

writeFileSync(join(DEPLOY_DIR, 'start.sh'), startScript)
copyFileSync(join(ROOT, 'package.json'), join(DEPLOY_DIR, 'package.json'))

// 8. 创建 Docker 部署文件
const dockerfile = `FROM ubuntu:24.04

# 安装依赖
RUN apt-get update && apt-get install -y \\
    curl \\
    gcc \\
    g++ \\
    make \\
    git \\
    xvfb \\
    libgtk-3-0 \\
    libnotify-dev \\
    libnss3 \\
    libxss1 \\
    libasound2 \\
    libxtst6 \\
    xauth \\
    && rm -rf /var/lib/apt/lists/*

# 安装 Node.js
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \\
    && apt-get install -y nodejs

# 安装 Rust
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
ENV PATH="/root/.cargo/bin:$PATH"

# 复制应用
COPY . /app
WORKDIR /app

# 构建
RUN npm install && npm run build:all

# 启动（使用 Xvfb 虚拟显示）
CMD ["xvfb-run", "--auto-servernum", "--server-args=-screen 0 1920x1080x24", "node", "main/index.js"]
`

writeFileSync(join(DEPLOY_DIR, 'Dockerfile'), dockerfile)

// 9. 创建 docker-compose.yml
const dockerCompose = `version: '3.8'
services:
  genoffice-web:
    build: .
    ports:
      - "8080:8080"
      - "5299:5299"
    environment:
      - DISPLAY=:99
    volumes:
      - ./data:/app/data
    restart: unless-stopped
`

writeFileSync(join(DEPLOY_DIR, 'docker-compose.yml'), dockerCompose)

log('Deployment files created in: ' + DEPLOY_DIR)
log('')
log('部署选项：')
log('1. 直接运行: cd deploy && ./start.sh')
log('2. Docker 部署: cd deploy && docker-compose up -d')
log('3. 手动启动: xvfb-run --auto-servernum node main/index.js')
