#!/usr/bin/env node
/**
 * GenOffice Web Server Release Script
 * 
 * 构建并发布 Web Server
 */

import { execSync } from 'child_process'
import { existsSync, mkdirSync, writeFileSync } from 'fs'
import { join, dirname } from 'path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const ROOT = join(__dirname, '..')

// 颜色输出
const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  blue: '\x1b[34m',
  yellow: '\x1b[33m',
  red: '\x1b[31m',
}

function log(color, ...args) {
  console.log(`${color}[GenOffice Release]${colors.reset}`, ...args)
}

function run(cmd, options = {}) {
  log(colors.blue, `Running: ${cmd}`)
  try {
    execSync(cmd, {
      cwd: ROOT,
      stdio: 'inherit',
      ...options,
    })
    return true
  } catch (error) {
    log(colors.red, `Failed: ${cmd}`)
    return false
  }
}

async function main() {
  const version = process.argv[2] || '0.8.0'
  const channel = process.argv[3] || 'feature/collaboration'
  
  log(colors.green, `=== GenOffice Web Server Release v${version} ===`)
  log(colors.yellow, `Channel: ${channel}`)
  
  // 1. 确保在正确的分支
  log(colors.blue, '\n[1/6] Checking git status...')
  if (!run('git status --porcelain')) {
    log(colors.red, 'Git status check failed')
    process.exit(1)
  }
  
  // 2. 安装依赖
  log(colors.blue, '\n[2/6] Installing dependencies...')
  if (!run('npm install')) {
    log(colors.red, 'npm install failed')
    process.exit(1)
  }
  
  // 3. 构建所有应用
  log(colors.blue, '\n[3/6] Building all apps...')
  if (!run('npm run build:all')) {
    log(colors.red, 'Build failed')
    process.exit(1)
  }
  
  // 4. 构建 Web Server
  log(colors.blue, '\n[4/6] Building Web Server...')
  if (!run('npm run build -w @genoffice/web-server')) {
    log(colors.red, 'Web Server build failed')
    process.exit(1)
  }
  
  // 5. 创建发布标签
  log(colors.blue, '\n[5/6] Creating release tag...')
  const tag = `web-server/v${version}`
  
  // 检查远程
  run('git fetch origin')
  run(`git fetch gitcode`)
  
  // 创建标签
  if (!run(`git tag -a ${tag} -m "GenOffice Web Server ${version}"`)) {
    log(colors.red, 'Tag creation failed')
    process.exit(1)
  }
  
  // 6. 推送
  log(colors.blue, '\n[6/6] Pushing to remotes...')
  
  // 推送到 GitHub
  log(colors.blue, 'Pushing to GitHub...')
  run(`git push origin ${channel} ${tag}`)
  
  // 推送到 GitCode
  log(colors.blue, 'Pushing to GitCode...')
  run(`git push gitcode ${channel} ${tag}`)
  
  log(colors.green, '\n=== Release Complete ===')
  log(colors.green, `Tag: ${tag}`)
  log(colors.green, `Branch: ${channel}`)
  log(colors.green, '\nTo create a release:')
  log(colors.blue, `  https://github.com/louloulin/genoffice/releases/new?tag=${tag}`)
}

main().catch(console.error)
