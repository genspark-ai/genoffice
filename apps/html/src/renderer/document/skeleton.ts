import { htmlLang, type Lang } from '@genoffice/i18n'

/**
 * A minimal, standards-mode document, for a blank one to start from.
 *
 * Each part earns its place: without the doctype the preview runs in quirks
 * mode, where box sizing and table layout follow different rules than the
 * author expects; without `lang` a screen reader has no language to read the
 * page in and the browser picks a font fallback and spell-checker for the wrong
 * one; without the charset a page of non-Latin text can decode as mojibake.
 *
 * A viewport meta is deliberately absent: this renders in a desktop pane with
 * no mobile viewport for it to affect.
 *
 * The document language follows the UI, so the skeleton a user inserts is the
 * one their tooling is already set up for.
 */
export function documentSkeleton(lang: Lang): string {
  return `<!DOCTYPE html>
<html lang="${htmlLang(lang)}">
  <head>
    <meta charset="UTF-8">
    <title></title>
  </head>
  <body></body>
</html>
`
}
