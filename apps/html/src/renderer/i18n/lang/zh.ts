/* zh dictionary: defines the key set every other locale shard must match. */
import { zh as appZh } from '../app/zh'
import { zh as aiZh } from '../ai/zh'

export const zh = {
  ...appZh,
  ...aiZh,
}
