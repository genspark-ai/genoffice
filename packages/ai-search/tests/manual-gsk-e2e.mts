// Manual e2e verification script (calls gsk for real, requires login; not in CI): npx tsx tests/manual-gsk-e2e.mts
import { webSearch, imageSearch, gskGenerateImage } from '../src/index'

const w = await webSearch('PowerPoint design trends 2026', 3)
console.log('webSearch method:', w.method, '| results:', w.results.length, '| first:', w.results[0]?.title?.slice(0, 60))

const im = await imageSearch('minimalist gradient background', 3)
console.log('imageSearch method:', im.method, '| images:', im.images.length, '| first:', im.images[0]?.imageUrl?.slice(0, 70))

const g = await gskGenerateImage({
  prompt: 'a simple flat vector icon of a lightbulb, blue background',
  aspectRatio: '1:1',
  imageSize: '0.5k',
})
console.log('generateImage url:', g.url.slice(0, 90))

// Verify the generated image is anonymously downloadable via plain fetch (image-insert flow depends on this)
const resp = await fetch(g.url)
console.log('download status:', resp.status, '| content-type:', resp.headers.get('content-type'))
