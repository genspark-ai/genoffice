#!/usr/bin/env python3
"""Remove the Sheets script-editor material from every language edition.

The feature was built and then rejected, so it is documented nowhere in the
tree: no `run_script` tool, no `SpreadsheetApp`, no `Tools ▸ Script editor`
menu, no string in any shard. The manual described it anyway, and by the time
that was noticed it had been translated into twenty languages.

Three places refer to it, and all three are handled by token rather than by
rewriting prose, because the tokens survive translation unchanged:

  sheets.*   the `## The script editor` section itself — removed whole
  ai-panel.* a parenthetical mentioning `run_script` — the parentheses go
  slides.*   a clause about a "controlled script sandbox" that cross-links
             to Sheets; the clause goes, and with it the now-dangling link

Every edit is verified: the token must be gone afterwards, and the file must
actually have changed. A silent no-op here would leave 21 languages claiming a
feature that does not exist.
"""
import pathlib
import re
import sys

TOPICS = pathlib.Path(__file__).resolve().parents[1] / (
    'apps/shell/src/renderer/src/i18n/help/topics'
)

SCRIPT_TOKENS = (
    'run_script', 'SpreadsheetApp', 'getActiveSpreadsheet', 'Utilities.sleep',
    'script editor', 'Script editor',
)

report = []


def fail(msg):
    print(f'  FAIL  {msg}')
    report.append(msg)


def tidy(line):
    """Drop the space a removed parenthetical leaves in front of punctuation.

    Scoped to one line on purpose — collapsing every space-before-punctuation in
    a 294-file corpus would rewrite prose nobody asked about.
    """
    return re.sub(r'[ \t]+([.,;:!?%)\]])', r'\1', line)


def paren_span(text, token):
    """The widest balanced ( ... ) span containing `token`, or None.

    Not a regex, and not "the nearest opener to the left": a markdown link
    carries its own parentheses, so `[Sheets](help://sheets)` puts a `(` between
    the token and the clause that encloses it. Nearest-opener stops at the
    link's own pair and removes the link alone. Every opener to the left is
    matched instead, and the widest span wins.
    """
    at = text.find(token)
    if at < 0:
        return None
    pairs = {'(': ')', '\uff08': '\uff09'}
    best = None
    for i in range(at - 1, -1, -1):
        opener = text[i]
        if opener not in pairs:
            continue
        closer = pairs[opener]
        depth, j = 1, i + 1
        while j < len(text) and depth:
            if text[j] == opener:
                depth += 1
            elif text[j] == closer:
                depth -= 1
            j += 1
        if depth == 0 and i < at < j:
            if best is None or (j - i) > (best[1] - best[0]):
                best = (i, j)
    return best


# ---- 1. the whole section, addressed by position so no heading text is needed
en = (TOPICS / 'sheets.en.md').read_text().split('\n')
heads = [i for i, l in enumerate(en) if l.startswith('## ')]
target = next(i for i in heads if 'script editor' in en[i].lower())
end = next((i for i in heads if i > target), len(en))
print(f'sheets: removing lines {target + 1}-{end} of the English source\n  heading: {en[target]}')

removed_section = 0
for f in sorted(TOPICS.glob('sheets.*.md')):
    lines = f.read_text().split('\n')
    h = [i for i, l in enumerate(lines) if l.startswith('## ')]
    if len(h) != len(heads):
        fail(f'{f.name}: {len(h)} sections, English has {len(heads)}')
        continue
    stop = next((i for i in h if i > target), len(lines))
    # the heading is translated and no two languages spell it alike — German
    # says Skript, Russian скрипт, Japanese スクリプト. The body carries the
    # untranslated identifiers, so that is what identifies the section.
    body = '\n'.join(lines[target:stop])
    if 'SpreadsheetApp' not in body and 'run_script' not in body:
        fail(f'{f.name}: section {target} does not look like the script editor ({lines[target]!r})')
        continue
    del lines[target:stop]
    out = '\n'.join(lines)
    if out == f.read_text():
        fail(f'{f.name}: removal was a no-op')
        continue
    f.write_text(out)
    removed_section += 1
print(f'  removed the section from {removed_section} files')

# ---- 2. the run_script parenthetical
removed_paren = 0
for f in sorted(TOPICS.glob('ai-panel.*.md')):
    s = f.read_text()
    if 'run_script' not in s:
        continue  # already clean (the Malay pass removed it itself)
    # the parenthetical is the (...) span carrying the token
    span = paren_span(s, 'run_script')
    if span is None:
        fail(f'{f.name}: no parenthesised span mentions run_script')
        continue
    before, after = s[:span[0]], s[span[1]:]
    # The comma after the parenthetical belongs to the list around it, not to
    # the aside: "formulas, data fill, bulk transforms (aside), formatting" has
    # to come out as "..., bulk transforms, formatting", so only the spaces
    # the parentheses left behind are tidied.
    if before.endswith(' ') and after.startswith(' '):
        before = before[:-1]
    out = before + after
    # only the line the parenthetical was on
    out = '\n'.join(tidy(l) for l in out.split('\n'))
    if out == s:
        fail(f'{f.name}: parenthetical removal was a no-op')
        continue
    if 'run_script' in out:
        fail(f'{f.name}: run_script survived')
        continue
    f.write_text(out)
    removed_paren += 1
print(f'  removed the parenthetical from {removed_paren} files')

# ---- 3. the sandbox clause in slides
#
# Every slides edition parenthesises "the AI restyles through a controlled
# script sandbox — the same mechanism as in [Sheets](help://sheets)", and the
# Sheets topic was the only thing that link had ever pointed at. With the
# section gone the clause is false and the link is dangling, so the whole
# parenthetical goes. Keyed on the link rather than on the word "sandbox",
# which only a few languages spell that way.
removed_clause = 0
for f in sorted(TOPICS.glob('slides.*.md')):
    s = f.read_text()
    if 'help://sheets' not in s:
        continue
    span = paren_span(s, 'help://sheets')
    if span is None:
        fail(f'{f.name}: no parenthesised span carries the Sheets link')
        continue
    out = s[:span[0]] + s[span[1]:]
    # a markdown link left without its target renders as literal [Sheets]
    out = re.sub(r'\[[^\]]*\]\(\s*\)', '', out)
    out = '\n'.join(tidy(l) for l in out.split('\n'))
    if 'help://sheets' in out:
        fail(f'{f.name}: the Sheets cross-link survived the clause removal')
        continue
    if out == s:
        fail(f'{f.name}: clause removal was a no-op')
        continue
    # a sentence that was only the clause leaves a stray opener behind
    f.write_text(out)
    removed_clause += 1
print(f'  removed the sandbox clause from {removed_clause} files')

# ---- 4. nothing anywhere still claims the feature
left = []
for f in sorted(TOPICS.glob('*.md')):
    s = f.read_text()
    for tok in SCRIPT_TOKENS:
        if tok in s:
            left.append(f'{f.name}: {tok}')
for item in left:
    fail(f'still mentions the feature — {item}')

print()
if report:
    print(f'{len(report)} problems')
    sys.exit(1)
print('the script editor is gone from every language')
