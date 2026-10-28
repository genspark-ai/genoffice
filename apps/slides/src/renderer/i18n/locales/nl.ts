import { nl as app } from '../app/nl'
import { nl as ribbon } from '../ribbon/nl'
import { nl as panes } from '../panes/nl'
import { nl as ai } from '../ai/nl'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const nl: Dict = { ...app, ...ribbon, ...panes, ...ai }
