/** Folder-tree data shared by the shell home screen and the editors' Files pane. */

export interface FolderRoot {
  path: string
  /** folder name shown on the root row */
  name: string
  /** false when the folder does not exist and cannot be created, or is read-only */
  usable: boolean
}

export interface FolderEntry {
  path: string
  name: string
  mtimeMs: number
  /** whether it contains at least one visible sub-folder (drives the expand chevron) */
  hasSubfolders: boolean
}

/** a document file listed by the tree (same shape as the home recents rows) */
export interface FileEntry {
  path: string
  name: string
  /** lowercased extension without the dot */
  ext: string
  mtimeMs: number
  sizeBytes: number
  starred: boolean
  /** the path failed to stat */
  missing?: boolean
}

export interface FolderListing {
  dir: string
  folders: FolderEntry[]
  /** supported document files directly inside `dir`, newest first */
  files: FileEntry[]
  /** the directory could not be read (deleted or moved outside the app) */
  missing?: boolean
}

/** what to do when a moved item's name already exists in the target */
export type MoveConflictPolicy = 'ask' | 'replace' | 'keepBoth' | 'skip'

export interface MoveResult {
  /** old path → new path for everything that moved */
  moved: Array<{ from: string; to: string }>
  /** items skipped because the name exists in the target (policy 'ask'/'skip') */
  conflicts: string[]
  /** items that failed for another reason */
  failed: Array<{ path: string; error: string }>
}

export interface FileOpResult {
  ok: boolean
  /** the new absolute path when ok */
  path?: string
  error?: string
}

/** what an editor's preload has to expose for the Files pane (all served by the shell main process) */
export interface FilesPaneApi {
  folderRoot(): Promise<FolderRoot>
  listFolder(dir: string): Promise<FolderListing>
  createFolder(parent: string, name: string): Promise<FileOpResult>
  renameFolder(dir: string, newName: string): Promise<FileOpResult>
  renameFile(path: string, newName: string): Promise<FileOpResult>
  movePaths(paths: string[], targetDir: string, onConflict: MoveConflictPolicy): Promise<MoveResult>
  deleteFolder(dir: string): Promise<void>
  deleteFiles(paths: string[]): Promise<void>
  /** open a document in the shell (new tab, or activates the tab already showing it) */
  openPath(path: string): Promise<void>
  revealPath(path: string): Promise<void>
  onFolderChanged(handler: (dirs: string[]) => void): () => void
}
