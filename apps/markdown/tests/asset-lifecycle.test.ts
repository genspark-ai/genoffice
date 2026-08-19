import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  discardPendingOwnedAssets,
  extractMarkdownImageSources,
  prepareAssetsForSaveAs,
  readOwnedAssetManifest,
  reconcileOwnedAssets,
  rewriteMarkdownImageSources,
  rollbackPreparedSaveAsAssets,
  writeImageIntoOwnedAssets,
} from '../src/main/asset-lifecycle'

const tempDirectories: string[] = []

async function temporaryDirectory(prefix: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), prefix))
  tempDirectories.push(directory)
  return directory
}

afterEach(async () => {
  await Promise.all(
    tempDirectories.splice(0).map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('Markdown owned asset lifecycle', () => {
  it("deletes only pending owned files when edits are explicitly discarded with Don't Save", async () => {
    const directory = await temporaryDirectory('markdown-assets-discard-')
    const documentPath = join(directory, 'note.md')
    const authored = await writeImageIntoOwnedAssets(
      documentPath,
      'pending.png',
      Buffer.from('pending image'),
    )

    const result = await discardPendingOwnedAssets(documentPath)

    expect(result.errors).toEqual([])
    expect(result.deleted).toEqual(['pending.png'])
    expect(existsSync(join(directory, authored))).toBe(false)
    expect(await readOwnedAssetManifest(documentPath)).toMatchObject({
      files: [],
      pending: {},
    })
  })

  it('preserves saved references and relinquishes modified pending files instead of deleting them', async () => {
    const directory = await temporaryDirectory('markdown-assets-discard-safe-')
    const documentPath = join(directory, 'note.md')
    const saved = await writeImageIntoOwnedAssets(
      documentPath,
      'saved.png',
      Buffer.from('saved image'),
    )
    await reconcileOwnedAssets(documentPath, [saved])
    const modified = await writeImageIntoOwnedAssets(
      documentPath,
      'modified.png',
      Buffer.from('original pending image'),
    )
    await writeFile(join(directory, modified), 'user modified image')

    const result = await discardPendingOwnedAssets(documentPath)

    expect(result.deleted).toEqual([])
    expect(result.preserved).toEqual(['modified.png'])
    expect(await readFile(join(directory, saved), 'utf8')).toBe('saved image')
    expect(await readFile(join(directory, modified), 'utf8')).toBe('user modified image')
    expect((await readOwnedAssetManifest(documentPath))?.files.map((file) => file.name)).toEqual([
      'saved.png',
    ])
  })

  it('keeps a deleted insertion undoable until save, then collects it', async () => {
    const directory = await temporaryDirectory('markdown-assets-gc-')
    const documentPath = join(directory, 'note.md')
    const authored = await writeImageIntoOwnedAssets(
      documentPath,
      'pasted.png',
      Buffer.from('image bytes'),
    )
    const assetPath = join(directory, authored)

    expect(existsSync(assetPath)).toBe(true)
    // Deleting the image node changes only editor state. The on-disk file stays
    // available if the user invokes Undo before saving.
    expect(existsSync(assetPath)).toBe(true)

    const result = await reconcileOwnedAssets(documentPath, [])
    expect(result.errors).toEqual([])
    expect(result.deleted).toEqual(['pasted.png'])
    expect(existsSync(assetPath)).toBe(false)
    expect((await readOwnedAssetManifest(documentPath))?.files).toEqual([])
  })

  it('deduplicates repeated references and protects assets shared by saved sibling documents', async () => {
    const directory = await temporaryDirectory('markdown-assets-shared-')
    const firstDocument = join(directory, 'first.md')
    const secondDocument = join(directory, 'second.md')
    const authored = await writeImageIntoOwnedAssets(
      firstDocument,
      'shared.png',
      Buffer.from('shared image'),
    )
    const assetPath = join(directory, authored)

    // A save in a sibling document must not collect first.md's still-undoable insertion.
    await reconcileOwnedAssets(secondDocument, [])
    expect(existsSync(assetPath)).toBe(true)

    await reconcileOwnedAssets(firstDocument, [authored, authored])
    await reconcileOwnedAssets(secondDocument, [authored])
    await reconcileOwnedAssets(firstDocument, [])
    expect(existsSync(assetPath)).toBe(true)

    await reconcileOwnedAssets(secondDocument, [authored])
    expect(existsSync(assetPath)).toBe(true)
    await reconcileOwnedAssets(secondDocument, [])
    expect(existsSync(assetPath)).toBe(false)
  })

  it('conservatively preserves an owned asset referenced by a sibling Markdown file', async () => {
    const directory = await temporaryDirectory('markdown-assets-sibling-')
    const documentPath = join(directory, 'owner.md')
    const siblingPath = join(directory, 'sibling.md')
    const authored = await writeImageIntoOwnedAssets(
      documentPath,
      'shared.png',
      Buffer.from('shared image'),
    )
    const assetPath = join(directory, authored)
    await writeFile(siblingPath, `![shared](${authored})`)

    await reconcileOwnedAssets(documentPath, [])
    expect(existsSync(assetPath)).toBe(true)

    await writeFile(siblingPath, '# No image now')
    await reconcileOwnedAssets(documentPath, [])
    expect(existsSync(assetPath)).toBe(false)
  })

  it('preserves an HTML-only sibling reference until that reference is removed', async () => {
    const directory = await temporaryDirectory('markdown-assets-html-sibling-')
    const documentPath = join(directory, 'owner.md')
    const siblingPath = join(directory, 'sibling.markdown')
    const authored = await writeImageIntoOwnedAssets(
      documentPath,
      'shared.png',
      Buffer.from('shared image'),
    )
    const assetPath = join(directory, authored)
    await writeFile(
      siblingPath,
      `<IMG data-src="ignored.png" SRC='assets&#47;shared.png' srcset="ignored-2x.png 2x">`,
    )

    await reconcileOwnedAssets(documentPath, [])
    expect(existsSync(assetPath)).toBe(true)

    await writeFile(siblingPath, '<p>No image now</p>')
    await reconcileOwnedAssets(documentPath, [])
    expect(existsSync(assetPath)).toBe(false)
  })

  it('retains owned assets when a sibling HTML image tag is malformed', async () => {
    const directory = await temporaryDirectory('markdown-assets-html-ambiguous-')
    const documentPath = join(directory, 'owner.md')
    const siblingPath = join(directory, 'sibling.md')
    const authored = await writeImageIntoOwnedAssets(
      documentPath,
      'shared.png',
      Buffer.from('shared image'),
    )
    await writeFile(siblingPath, '<img src="assets/shared.png>')

    const result = await reconcileOwnedAssets(documentPath, [])

    expect(result.deleted).toEqual([])
    expect(existsSync(join(directory, authored))).toBe(true)
  })

  it('retains owned assets when a sibling HTML image tag has duplicate src attributes', async () => {
    const directory = await temporaryDirectory('markdown-assets-html-duplicate-src-')
    const documentPath = join(directory, 'owner.md')
    const siblingPath = join(directory, 'sibling.md')
    const authored = await writeImageIntoOwnedAssets(
      documentPath,
      'shared.png',
      Buffer.from('shared image'),
    )
    await writeFile(siblingPath, `<img src="${authored}" SRC="assets/other.png">`)

    const result = await reconcileOwnedAssets(documentPath, [])

    expect(result.deleted).toEqual([])
    expect(existsSync(join(directory, authored))).toBe(true)
  })

  it('never deletes untracked or user-modified files', async () => {
    const directory = await temporaryDirectory('markdown-assets-user-')
    const documentPath = join(directory, 'note.md')
    const owned = await writeImageIntoOwnedAssets(
      documentPath,
      'owned.png',
      Buffer.from('original'),
    )
    const assetsDirectory = join(directory, 'assets')
    const manualPath = join(assetsDirectory, 'manual.png')
    await writeFile(manualPath, 'manual')
    await writeFile(join(directory, owned), 'user replaced this file')

    await reconcileOwnedAssets(documentPath, [])

    expect(await readFile(manualPath, 'utf8')).toBe('manual')
    expect(await readFile(join(directory, owned), 'utf8')).toBe('user replaced this file')
    expect((await readOwnedAssetManifest(documentPath))?.files).toEqual([])
  })

  it('copies every safe local image on Save As without touching the old document assets', async () => {
    const sourceDirectory = await temporaryDirectory('markdown-assets-source-')
    const targetDirectory = await temporaryDirectory('markdown-assets-target-')
    const sourceDocument = join(sourceDirectory, 'source.md')
    const targetDocument = join(targetDirectory, 'copy.md')
    const owned = await writeImageIntoOwnedAssets(
      sourceDocument,
      'owned.png',
      Buffer.from('owned source'),
    )
    await mkdir(join(sourceDirectory, 'images'))
    await writeFile(join(sourceDirectory, 'images', 'manual.png'), 'manual source')

    // A different user-managed file at the preferred destination must not be overwritten.
    await mkdir(join(targetDirectory, 'assets'))
    await writeFile(join(targetDirectory, 'assets', 'owned.png'), 'target user file')
    const text = [
      `![owned](${owned})`,
      '![manual](images/manual.png)',
      `![duplicate](${owned})`,
      '![unsafe](../outside.png)',
    ].join('\n')
    const imageSources = [owned, 'images/manual.png', owned, '../outside.png']

    const prepared = await prepareAssetsForSaveAs(
      sourceDocument,
      targetDocument,
      text,
      imageSources,
    )
    await writeFile(targetDocument, prepared.text)
    await reconcileOwnedAssets(targetDocument, prepared.imageSources)

    expect(prepared.rewrites).toEqual(
      expect.arrayContaining([
        { from: owned, to: 'assets/owned-1.png' },
        { from: 'images/manual.png', to: 'assets/manual.png' },
      ]),
    )
    expect(prepared.text).toContain('![owned](assets/owned-1.png)')
    expect(prepared.text).toContain('![duplicate](assets/owned-1.png)')
    expect(prepared.text).toContain('![manual](assets/manual.png)')
    expect(prepared.text).toContain('![unsafe](../outside.png)')
    expect(await readFile(join(targetDirectory, 'assets', 'owned.png'), 'utf8')).toBe(
      'target user file',
    )
    expect(await readFile(join(targetDirectory, 'assets', 'owned-1.png'), 'utf8')).toBe(
      'owned source',
    )
    expect(await readFile(join(targetDirectory, 'assets', 'manual.png'), 'utf8')).toBe(
      'manual source',
    )
    expect(await readFile(join(sourceDirectory, owned), 'utf8')).toBe('owned source')
    expect(await readFile(join(sourceDirectory, 'images', 'manual.png'), 'utf8')).toBe(
      'manual source',
    )
    expect((await readOwnedAssetManifest(targetDocument))?.files.map((file) => file.name)).toEqual([
      'manual.png',
      'owned-1.png',
    ])
  })

  it('copies and rewrites raw HTML image sources on Save As', async () => {
    const sourceDirectory = await temporaryDirectory('markdown-assets-html-source-')
    const targetDirectory = await temporaryDirectory('markdown-assets-html-target-')
    const sourceDocument = join(sourceDirectory, 'source.md')
    const targetDocument = join(targetDirectory, 'copy.md')
    await mkdir(join(sourceDirectory, 'images'))
    await writeFile(join(sourceDirectory, 'images', 'html.png'), 'html image')
    const text =
      `<figure data-kind="keep"><IMG alt="preview" SRC='images&#47;html.png'></figure>\n` +
      '<img src=images/html.png />'

    const prepared = await prepareAssetsForSaveAs(sourceDocument, targetDocument, text, [])

    expect(prepared.rewrites).toEqual([{ from: 'images/html.png', to: 'assets/html.png' }])
    expect(prepared.text).toBe(
      `<figure data-kind="keep"><IMG alt="preview" SRC='assets/html.png'></figure>\n` +
        '<img src=assets/html.png />',
    )
    expect(prepared.imageSources).toEqual(['assets/html.png'])
    expect(await readFile(join(targetDirectory, 'assets', 'html.png'), 'utf8')).toBe('html image')
  })

  it('rejects traversal image sources during Save As', async () => {
    const root = await temporaryDirectory('markdown-assets-traversal-')
    const sourceDirectory = join(root, 'source')
    const targetDirectory = join(root, 'target')
    await mkdir(sourceDirectory)
    await mkdir(targetDirectory)
    await writeFile(join(root, 'outside.png'), 'outside image')
    const sourceDocument = join(sourceDirectory, 'source.md')
    const targetDocument = join(targetDirectory, 'copy.md')
    const text = '<img src="../outside.png">'

    const prepared = await prepareAssetsForSaveAs(sourceDocument, targetDocument, text, [])

    expect(prepared.rewrites).toEqual([])
    expect(prepared.created).toEqual([])
    expect(prepared.text).toBe(text)
  })

  it.skipIf(process.platform === 'win32')(
    'rejects image symlinks that escape the source document directory',
    async () => {
      const root = await temporaryDirectory('markdown-assets-symlink-')
      const sourceDirectory = join(root, 'source')
      const targetDirectory = join(root, 'target')
      await mkdir(sourceDirectory)
      await mkdir(targetDirectory)
      const outsidePath = join(root, 'outside.png')
      await writeFile(outsidePath, 'outside image')
      await symlink(outsidePath, join(sourceDirectory, 'linked.png'))
      const sourceDocument = join(sourceDirectory, 'source.md')
      const targetDocument = join(targetDirectory, 'copy.md')
      const text = '<img src="linked.png">'

      const prepared = await prepareAssetsForSaveAs(sourceDocument, targetDocument, text, [])

      expect(prepared.rewrites).toEqual([])
      expect(prepared.created).toEqual([])
      expect(prepared.text).toBe(text)
    },
  )

  it('persists pending ownership removal when rollback cannot delete a shared copied file', async () => {
    const sourceDirectory = await temporaryDirectory('markdown-assets-rollback-source-')
    const targetDirectory = await temporaryDirectory('markdown-assets-rollback-target-')
    const sourceDocument = join(sourceDirectory, 'source.md')
    const targetDocument = join(targetDirectory, 'copy.md')
    await writeFile(sourceDocument, '# Source')
    await writeFile(join(sourceDirectory, 'local.png'), 'source image')
    const prepared = await prepareAssetsForSaveAs(
      sourceDocument,
      targetDocument,
      '![local](local.png)',
      ['local.png'],
    )
    const copied = prepared.imageSources[0]!
    const siblingDocument = join(targetDirectory, 'sibling.md')
    await writeFile(siblingDocument, `![shared](${copied})`)
    await reconcileOwnedAssets(siblingDocument, [copied])

    await rollbackPreparedSaveAsAssets(prepared)

    expect(existsSync(join(targetDirectory, copied))).toBe(true)
    expect(await readOwnedAssetManifest(targetDocument)).toMatchObject({
      files: [{ name: 'local.png' }],
      documents: { 'sibling.md': ['local.png'] },
      pending: {},
    })
  })

  it('persists pending ownership removal when a copied file fails hash validation', async () => {
    const sourceDirectory = await temporaryDirectory('markdown-assets-hash-source-')
    const targetDirectory = await temporaryDirectory('markdown-assets-hash-target-')
    const sourceDocument = join(sourceDirectory, 'source.md')
    const targetDocument = join(targetDirectory, 'copy.md')
    await writeFile(sourceDocument, '# Source')
    await writeFile(join(sourceDirectory, 'local.png'), 'source image')
    const prepared = await prepareAssetsForSaveAs(
      sourceDocument,
      targetDocument,
      '![local](local.png)',
      ['local.png'],
    )
    const copied = prepared.imageSources[0]!
    await writeFile(join(targetDirectory, copied), 'user changed the copied file')

    await rollbackPreparedSaveAsAssets(prepared)

    expect(await readFile(join(targetDirectory, copied), 'utf8')).toBe(
      'user changed the copied file',
    )
    expect(await readOwnedAssetManifest(targetDocument)).toMatchObject({
      files: [{ name: 'local.png' }],
      pending: {},
    })
  })
})

describe('serialized Markdown image references', () => {
  it('extracts destinations with titles and duplicate references', () => {
    expect(
      extractMarkdownImageSources(
        '![one](assets/a.png "title")\n![two](<assets/b image.png>)\n![again](assets/a.png)',
      ),
    ).toEqual(['assets/a.png', 'assets/b image.png', 'assets/a.png'])
  })

  it('extracts and rewrites mixed Markdown and HTML image sources without touching lookalikes', () => {
    const text = [
      '![markdown](assets/a.png "title")',
      '<IMG data-src="ignored.png" SRC = "assets&#47;b&amp;c.png" srcset="ignored-2x.png 2x">',
      "<img class=x src='assets/c.png'>",
      '<img src=assets/d.png />',
    ].join('\n')
    const rewritten = rewriteMarkdownImageSources(
      text,
      new Map([
        ['assets/a.png', 'assets/a-new.png'],
        ['assets/b&c.png', 'assets/b&new.png'],
        ['assets/c.png', "assets/c's.png"],
        ['assets/d.png', 'assets/d new.png'],
        ['ignored.png', 'assets/must-not-change.png'],
      ]),
    )

    expect(extractMarkdownImageSources(text)).toEqual([
      'assets/a.png',
      'assets/b&c.png',
      'assets/c.png',
      'assets/d.png',
    ])
    expect(rewritten).toBe(
      [
        '![markdown](assets/a-new.png "title")',
        '<IMG data-src="ignored.png" SRC = "assets/b&amp;new.png" srcset="ignored-2x.png 2x">',
        "<img class=x src='assets/c&#39;s.png'>",
        '<img src=assets/d&#32;new.png />',
      ].join('\n'),
    )
    expect(extractMarkdownImageSources(rewritten)).toEqual([
      'assets/a-new.png',
      'assets/b&new.png',
      "assets/c's.png",
      'assets/d new.png',
    ])
  })

  it('does not treat image examples inside Markdown code as rendered assets', () => {
    const text = [
      '`<img src="assets/inline.png">`',
      '```html',
      '<img src="assets/fenced.png">',
      '```',
      '~~~markdown',
      '![example](assets/fenced-markdown.png)',
      '~~~',
      '    <img src="assets/indented.png">',
      '<img src="assets/rendered.png">',
      '![rendered](assets/rendered-markdown.png)',
    ].join('\n')
    const rewritten = rewriteMarkdownImageSources(
      text,
      new Map([
        ['assets/inline.png', 'assets/inline-new.png'],
        ['assets/fenced.png', 'assets/fenced-new.png'],
        ['assets/fenced-markdown.png', 'assets/fenced-markdown-new.png'],
        ['assets/indented.png', 'assets/indented-new.png'],
        ['assets/rendered.png', 'assets/rendered-new.png'],
        ['assets/rendered-markdown.png', 'assets/rendered-markdown-new.png'],
      ]),
    )

    expect(extractMarkdownImageSources(text)).toEqual([
      'assets/rendered.png',
      'assets/rendered-markdown.png',
    ])
    expect(rewritten).toContain('`<img src="assets/inline.png">`')
    expect(rewritten).toContain('<img src="assets/fenced.png">')
    expect(rewritten).toContain('![example](assets/fenced-markdown.png)')
    expect(rewritten).toContain('    <img src="assets/indented.png">')
    expect(rewritten).toContain('<img src="assets/rendered-new.png">')
    expect(rewritten).toContain('![rendered](assets/rendered-markdown-new.png)')
  })
})
