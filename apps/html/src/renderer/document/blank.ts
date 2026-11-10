// no head-boundary guessing (`</head>` and `<body>` are optional): drop only elements whose text is never rendered
const STRIP_RE =
  /<!--[\s\S]*?-->|<(title|script|style|template|noscript)(?:\s[^>]*)?>[\s\S]*?<\/\1\s*>/gi
const MEDIA_RE = /<(?:img|svg|video|audio|canvas|iframe|picture|object|embed)\b/i

/** True when the page has nothing visible yet: no body text and no media, so the AI panel offers generation instead of edits. */
export function isDocEmpty(text: string): boolean {
  const body = text.replace(STRIP_RE, '')
  if (MEDIA_RE.test(body)) return false
  return (
    body
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;|&#160;/gi, ' ')
      .trim() === ''
  )
}
