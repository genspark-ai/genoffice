(function installVisualEffectClassifier(global) {
  function needsScreenshot(el, deps) {
    const { bgHex, cs, fullBorderOf } = deps;
    // A broken image's rendered box is Chrome's placeholder chrome (icon +
    // clipped alt text), never the author's design — even when transformed
    // or shadowed, let the IMG branch fall back to alt text instead.
    if (el.tagName === 'IMG' && (!el.complete || el.naturalWidth === 0)) return false;
    const style = cs(el);
    const rect = el.getBoundingClientRect();
    if (rect.width < 20 || rect.height < 12 || rect.width > 1000 || rect.height > 1000) {
      return false;
    }
    const transformed = style.transform && style.transform !== 'none';
    const shadowed = style.boxShadow && style.boxShadow !== 'none';
    const filtered =
      (style.filter && style.filter !== 'none') ||
      (style.backdropFilter && style.backdropFilter !== 'none');
    const radii = [
      style.borderTopLeftRadius,
      style.borderTopRightRadius,
      style.borderBottomRightRadius,
      style.borderBottomLeftRadius,
    ].map((value) => parseFloat(value) || 0);
    const complexRadius =
      Math.max(...radii) >= 8 &&
      (Boolean(bgHex(el)) || Boolean(fullBorderOf(el))) &&
      (new Set(radii.map((radius) => Math.round(radius))).size > 1 ||
        style.overflow === 'hidden' ||
        Math.max(...radii) >= Math.min(rect.width, rect.height) * 0.25);
    // Text-heavy cards must stay editable/searchable: soft shadows, filters
    // and rounded corners are not worth rasterizing paragraphs of content.
    // Transforms still rasterize (flowed text would land in the wrong place).
    // CJK counts double: a 60-char Japanese question carries as much content
    // as a 120-char English sentence.
    const text = (el.innerText || '').trim();
    const cjkCount = (text.match(/[\u3000-\u9fff\uf900-\ufaff\uff66-\uff9f]/g) || []).length;
    const weightedTextLength = text.length + cjkCount;
    // Many short lines (schedule cards, subject lists) are as content-heavy
    // as long prose \u2014 a 5-day timetable trapped in pixels is uneditable.
    const lineCount = text ? text.split('\n').filter((line) => line.trim()).length : 0;
    const textHeavy = weightedTextLength > 80 || (lineCount >= 4 && text.length >= 30);
    const compactRoundedCard =
      complexRadius && rect.height <= 160 && weightedTextLength <= 200 && lineCount < 4;
    return Boolean(
      transformed ||
      ((shadowed || filtered || complexRadius) && (!textHeavy || compactRoundedCard)),
    );
  }

  global.__html2docxVisualEffects = { needsScreenshot };
})(globalThis);
