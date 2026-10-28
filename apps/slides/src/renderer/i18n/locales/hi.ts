import { hi as app } from '../app/hi'
import { hi as ribbon } from '../ribbon/hi'
import { hi as panes } from '../panes/hi'
import { hi as ai } from '../ai/hi'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const hi: Dict = { ...app, ...ribbon, ...panes, ...ai }
