import { zhTW as app } from '../app/zh-TW'
import { zhTW as ribbon } from '../ribbon/zh-TW'
import { zhTW as panes } from '../panes/zh-TW'
import { zhTW as ai } from '../ai/zh-TW'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const zhTW: Dict = { ...app, ...ribbon, ...panes, ...ai }
