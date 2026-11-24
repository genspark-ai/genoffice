/**
 * jsdom has no layout and leaves Range without the rectangle readers every
 * browser has. The editor scrolls its selection into view a frame after a
 * focus command; without them that frame throws outside any test.
 */
const none = { length: 0, item: () => null, [Symbol.iterator]: [][Symbol.iterator] }
if (typeof Range !== 'undefined') {
  Range.prototype.getClientRects ??= () => none as unknown as DOMRectList
  Range.prototype.getBoundingClientRect ??= () => new DOMRect()
}
