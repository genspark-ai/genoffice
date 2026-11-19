import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'

/** A small physical page makes each discussion-capacity boundary independently observable. */
export async function capacityDocument({
  height,
  count = 2,
  replyLines = 0,
  secondPage = false,
  bottom = false,
  rootLines = 10,
  widths = [9000, 12000],
}: {
  height: number
  count?: number
  replyLines?: number
  secondPage?: boolean
  bottom?: boolean
  rootLines?: number
  widths?: readonly [number, number]
}): Promise<Buffer> {
  const body = (lines: number) =>
    Array.from({ length: lines }, (_, i) => `${i ? '<w:br/>' : ''}<w:t>Review ${i + 1}</w:t>`).join(
      '',
    )
  const comments: string[] = [],
    extended: string[] = []
  for (let i = 0; i < count + (secondPage ? 1 : 0); i++) {
    const para = (i + 1).toString(16).padStart(8, '0')
    comments.push(
      `<w:comment w:id="${i}" w:author="Reviewer"><w:p w14:paraId="${para}"><w:r>${body(rootLines)}</w:r></w:p></w:comment>`,
    )
    extended.push(`<w15:commentEx w15:paraId="${para}"/>`)
    if (replyLines && i < count) {
      const replyPara = (i + 100).toString(16).padStart(8, '0')
      comments.push(
        `<w:comment w:id="${i + 100}" w:author="Reply reviewer"><w:p w14:paraId="${replyPara}"><w:r>${body(replyLines)}</w:r></w:p></w:comment>`,
      )
      extended.push(`<w15:commentEx w15:paraId="${replyPara}" w15:paraIdParent="${para}"/>`)
    }
  }
  const sect = (width: number) =>
    `<w:sectPr><w:pgSz w:w="${width}" w:h="${height}"/><w:pgMar w:top="240" w:bottom="240" w:left="720" w:right="720"/></w:sectPr>`
  const anchor = (id: number) =>
    `<w:commentRangeStart w:id="${id}"/><w:r><w:t>Target${id} </w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`
  const filler = bottom
    ? Array.from(
        { length: 25 },
        () =>
          '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>Filler</w:t></w:r></w:p>',
      ).join('')
    : ''
  return Buffer.from(
    await buildDocx({
      bodyXml: `${filler}<w:p>${Array.from({ length: count }, (_, i) => anchor(i)).join('')}</w:p>${secondPage ? `<w:p><w:pPr>${sect(widths[0])}</w:pPr></w:p><w:p>${anchor(count)}</w:p>` : ''}`,
      sectPrXml: sect(secondPage ? widths[1] : widths[0]),
      extraRels:
        '<Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${comments.join('')}</w:comments>`,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
        {
          path: 'word/commentsExtended.xml',
          xml: `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${extended.join('')}</w15:commentsEx>`,
          contentType: 'application/vnd.ms-word.commentsExtended+xml',
        },
      ],
    }),
  )
}
