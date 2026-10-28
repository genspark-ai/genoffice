import { fr as app } from '../app/fr'
import { fr as ribbon } from '../ribbon/fr'
import { fr as panes } from '../panes/fr'
import { fr as ai } from '../ai/fr'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const fr: Dict = { ...app, ...ribbon, ...panes, ...ai }
