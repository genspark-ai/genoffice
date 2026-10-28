import { vi as app } from '../app/vi'
import { vi as ribbon } from '../ribbon/vi'
import { vi as panes } from '../panes/vi'
import { vi as ai } from '../ai/vi'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const vi: Dict = { ...app, ...ribbon, ...panes, ...ai }
