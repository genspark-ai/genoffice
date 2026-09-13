# Designing a deck from a spec

The same workflow the GenOffice app runs when it generates a presentation, written for an agent that does the thinking itself and lets the engine do the building. In the app every stage is a separate model call; here every stage is a separate file, and the CLI checks each file before the next stage may start. Work top-down; never start by placing elements.

Files of one deck, in a folder of their own (`deck/`):

| Stage         | File                                                         | Check                                                                                                                      |
| ------------- | ------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------- |
| 1 style sheet | `deck/style.md`                                              | read it back before every page                                                                                             |
| 2 outline     | `deck/outline.json`                                          | `genoffice slides check deck/outline.json --json`                                                                          |
| 3 pages       | `deck/pages/01.json`, `02.json`, … one file per outline page | `genoffice slides check deck/pages/NN.json --json` after writing each (also against the outline entry and the style sheet) |
| 4 build       | `deck/deck.pptx`                                             | `genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/deck.pptx --json`                   |
| 5 QC          | `deck/shots/*.png`                                           | `genoffice slides render`, `genoffice slides audit`, `genoffice slides replace`                                            |

Never put more than one page in a file, and never write a page before its outline entry exists: the build refuses to run until every outline page has its file, and a page written in one focused step with the style sheet and its brief in front of you is what keeps a 20-page deck consistent. Keep the three files together: `slides check`, `create --spec` and `slides replace` look for `outline.json` and `style.md` beside the page files and one folder up, check every page against its outline entry (file `NN.json` is entry `pages[N-1]`) and against the style sheet's colors, and say so in `detail.notes` when either file is missing.

## 0. Check what the machine can do

Run `genoffice capabilities --json` once. It reads GenOffice's own settings and reports, without a network call, whether web search, image search, image generation and media analysis are configured (a Genspark login with cloud tools on, or a key the user entered in Settings). Only when a feature is configured may the deck use it; when nothing is configured, work from the material you have and use typography, color blocks and shapes instead of photos. Never ask the user to configure a key just for a deck.

## 1. Style sheet first (one per deck)

Decide the design system before any content and write it to `deck/style.md`; re-read that file before every page. Use this shape, with concrete values:

- **Backgrounds**: main background; one per page type (cover, content, data, closing). Honor an explicit tone (dark theme, brand colors) before falling back to a light neutral. Content pages share one background.
- **Colors**: main text, primary accent, secondary accent, card background, border, written as `#RRGGBB` values. One accent system for the whole deck: even when comparing several companies or options, do not give each its own color. Never more than the two accents. The page check reports every color a page uses that the style sheet does not name (`detail.style.offPalette`); white and black are always allowed.
- **Fonts**: title font (Latin and CJK if the deck mixes scripts), body font, title size range, body size range. Use fonts that exist on the target machine; when unsure, leave `font` out and let the engine's default apply.
- **Layout library** (at least two variants per page type):
  - cover: `cover_typography_hero` (huge type, no image), `cover_dark_minimal`, `cover_split_color` (60/40 color blocks), `cover_full_image_overlay` (photo plus dark overlay), `cover_magazine`, `cover_split_image` (text left, image right)
  - content: `left_text_right_image`, `three_column_cards`, `hero_big_number`, `two_column_comparison`, `timeline_horizontal`, `full_image_text_overlay`
  - data: `kpi_cards_row`, `chart_with_insight`, `two_by_two_grid`
  - closing: `closing_cta`, `closing_thank_you`
- **Overall style**: one sentence describing the design language.

## 2. Outline, page by page

Write `deck/outline.json`:

```json
{
  "topic": "…",
  "core_hook": "one sentence with tension, a number or a counter-intuitive contrast",
  "pages": [
    {
      "title": "…",
      "type": "cover | content | data | closing",
      "layout": "one variant from the style sheet's library for that type",
      "brief": "what goes in each region, with the real facts and figures",
      "image_queries": ["English scene query or http(s) URL", "…"]
    }
  ]
}
```

Then `genoffice slides check deck/outline.json --json`: errors (unknown layout, a content layout repeated back to back, placeholder copy, missing core hook) exit 1 and must be fixed in the file; warnings are advice. Rules the check enforces and the ones it cannot:

- A core hook for the whole deck: one sentence with tension, a number or a counter-intuitive contrast.
- Content pages in one deck must not repeat the same layout variant back to back, and the deck should use at least three different variants.
- Pick the variant from the content: three parallel points → `three_column_cards`; one key number → `hero_big_number`; comparison → `two_column_comparison` or `two_by_two_grid`; a sequence → `timeline_horizontal`; metrics → `kpi_cards_row`.
- No placeholders ("XX%", "lorem"): every figure and name comes from the material you have. With web search configured, look up the facts and figures the brief needs before writing pages (`genoffice search "<query>" --json`). Without material and without search, tell the user which figures are illustrative; never present invented numbers as facts.
- Photos: with image search configured, give each photo slot an English query for a concrete scene ("harbor cranes at dawn", not "logistics") and fetch candidates with `genoffice search --images "<query>" --max 6 --json`; put a result's `imageUrl` straight into the page's image element (the builder downloads it), and keep its `sourceUrl` (the page the picture came from) for the closing page's credits: the `imageUrl` is often a search-provider proxy that says nothing about origin or license. Give the image box roughly the picture's aspect ratio; a portrait shot squeezed into a flat strip is center-cropped to a sliver and reported by `slides check`. With image generation configured, `genoffice image "<prompt>" --aspect 16:9 --out assets/p3.jpg` makes a photo or illustration in the deck's palette. With neither, plan pages that need no photo.

## 3. Write each page spec, one file at a time

For page N: re-read `deck/style.md` and the page's outline entry, resolve its `image_queries` to files or URLs, write `deck/pages/NN.json` (two-digit, one page object, see `genoffice guide slides spec`) echoing the entry's `title`, `type` and `layout` at the top of the object, and run `genoffice slides check deck/pages/NN.json --json`. It builds that page alone and returns dropped elements, unreadable images and the geometry audit (out of bounds, text taller than its box, overlaps), then checks the page against outline entry N (`detail.outline.findings`: an echoed field that disagrees or placeholder copy is an error and exits 1; a title missing from the page text or a planned photo without an image element is a warning) and against the style sheet (`detail.style.offPalette`). Fix the file until `audit`, `findings` and `offPalette` are empty, then move to page N+1. Do not batch pages: one page per step, checked, is the whole point.

Canvas is 1280 × 720 px, origin top-left, integers only; nothing may cross the edges. Elements paint in array order: background and decor shapes first, then images, text last.

**Sizing text boxes** (the boxes have zero inner padding; the top-left is where the first glyph starts):

- one line is about `sizePt × 1.8` px tall at `lineSpacingPct` 110
- a CJK character is about `sizePt × 1.35` px wide, a Latin character about `sizePt × 0.7` px
- text wraps at the box width: count the wrapped lines, make the box that tall plus one spare line
- the builder re-measures every plain text box and grows it to fit; it never shrinks, so undersized boxes push into what sits below them. Size correctly instead of relying on the fix.

**Spacing**: ≥ 8 px between text and a card edge, ≥ 20 px between a title and its subtitle, ≥ 5 px between stacked text blocks. Check every pair before output.

**Type scale (pt)**: big titles 32 to 48, subtitles 18 to 24, body 12 to 15, hero numbers up to 80.

**Fill the page**: spread content over the whole canvas; no content crammed into the top half above a blank field. Make text and images as large as the layout allows.

**Visuals**: photos only from sources you actually have (local files, search results, generated images). Without photos, work with typography, color blocks and shapes; never fake a photo with a grey box. Icon-like decoration uses the allowed shapes, at most four or five per page, content-related. No emoji. Data visuals are composed from `rect`, `donut` and `line` shapes with sizes proportional to the real values. Solid colors only (alpha allowed).

## 4. Anti-patterns (treat as defects)

- A thin vertical accent bar on the left of every card, a colored bar on top of cards, a small bar left of titles. Express hierarchy with background, weight and size contrast.
- Rainbow cards: a different accent per item.
- Decorative corner blocks or short lines; decoration that moves around from page to page.
- Every page as "shape + bold subtitle + description" list. The cover must not be a flat title plus subtitle: give it a visual anchor (a large color block, a geometric composition, a huge number, a hero image).
- A deck-wide header bar with the page title on every page. Vary the composition.
- Sibling cards whose hero numbers mix metrics (a market share next to a price next to a count): one row, one unit, one source. A number's label must say exactly what the source measured (bone-conduction share is not sports-headphone share).
- Two lines that say the same thing: a subtitle that restates the chart caption, a takeaway that repeats a bullet.

## 5. Build, look, audit, fix

1. `genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/deck.pptx --json` (`--outline` may be left out when the file sits beside the folder). Page files are taken in name order; the build refuses to run while an outline page has no file or a page file disagrees with its entry. Read `detail.issues`, `detail.imageFailures`, `detail.outline.findings` and `detail.style.offPalette`: each names the page file.
2. `genoffice slides render deck/deck.pptx --out deck/shots --json` and look at every PNG (needs GenOffice installed). Check overflow, collisions, contrast, crop, hierarchy, whitespace.
3. `genoffice slides audit deck/deck.pptx --json` for the geometry findings the eye misses. Element ids in the findings are the ones `genoffice slides read` and `genoffice slides apply` use.
4. Fix a page in its spec file and put it back with `genoffice slides replace deck/deck.pptx --slide <n> --spec deck/pages/NN.json --json` (slide n is page file n+1; the page is checked against outline entry n first); the other slides are untouched. Small nudges can also go through `genoffice slides apply` ops (`setTransform`, `setText`, `setTextStyle`, `setFill` …). At most two fix rounds; then report what remains.

Only fix objective defects and restrained polish: keep the copy, facts, palette and composition. Do not redesign a page that is already clean.
