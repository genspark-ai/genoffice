import { th as app } from '../app/th'
import { th as ribbon } from '../ribbon/th'
import { th as panes } from '../panes/th'
import { th as ai } from '../ai/th'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const th: Dict = { ...app, ...ribbon, ...panes, ...ai }
