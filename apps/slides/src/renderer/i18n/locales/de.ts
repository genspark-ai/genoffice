import { de as app } from '../app/de'
import { de as ribbon } from '../ribbon/de'
import { de as panes } from '../panes/de'
import { de as ai } from '../ai/de'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const de: Dict = { ...app, ...ribbon, ...panes, ...ai }
