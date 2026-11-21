import { Fragment } from 'react'

/** a run's text with a soft-hyphen span between its pieces (editor/hyphenation-static hyphenPieces) */
export function HyphenatedText({ pieces }: { pieces: readonly string[] }) {
  return (
    <>
      {pieces.map((piece, i) => (
        <Fragment key={i}>
          {i > 0 && <span className="doc-shy" />}
          {piece}
        </Fragment>
      ))}
    </>
  )
}
