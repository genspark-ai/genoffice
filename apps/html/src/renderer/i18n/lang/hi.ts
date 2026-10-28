import { hi as appHi } from '../app/hi'
import { hi as aiHi } from '../ai/hi'

export const hi = {
  ...appHi,
  ...aiHi,
}
