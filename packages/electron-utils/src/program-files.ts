import { extname } from 'node:path'

/** Executable, script, shortcut, and launcher extensions that should not be launched via default app */
const PROGRAM_EXTENSIONS = new Set([
  // Windows executables, scripts, shortcuts, and installers
  '.exe',
  '.com',
  '.scr',
  '.msi',
  '.bat',
  '.cmd',
  '.ps1',
  '.vbs',
  '.vbe',
  '.js',
  '.jse',
  '.wsf',
  '.hta',
  '.lnk',
  '.url',
  '.reg',
  '.cpl',
  '.pif',
  '.gadget',
  '.msc',
  '.inf',
  '.ins',
  '.isp',
  // macOS applications, scripts, workflows, and packages
  '.app',
  '.command',
  '.action',
  '.workflow',
  '.scpt',
  '.applescript',
  '.terminal',
  '.pkg',
  '.mpkg',
  // Linux launchers, binaries, and shell scripts
  '.desktop',
  '.sh',
  '.bash',
  '.zsh',
  '.csh',
  '.ksh',
  '.bin',
  '.run',
  '.out',
  // Java archives
  '.jar',
])

export function isProgramFile(path: string): boolean {
  return PROGRAM_EXTENSIONS.has(extname(path).toLowerCase())
}

const LINK_SOURCE_EXTENSIONS = new Set([
  '.docx',
  '.doc',
  '.dotx',
  '.rtf',
  '.txt',
  '.md',
  '.pdf',
  '.xlsx',
  '.xls',
  '.csv',
  '.pptx',
  '.ppt',
  '.odt',
  '.ods',
  '.odp',
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
])

/** A file a document's link may open in its default app; a document names the path, so anything else is only shown in its folder */
export function isOpenableLinkSource(path: string): boolean {
  return LINK_SOURCE_EXTENSIONS.has(extname(path).toLowerCase())
}
