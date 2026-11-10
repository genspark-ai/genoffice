(function installFormsAndMediaExtractor(global) {
  // Empty block/inline with only a bottom border is a fill-in underline
  // (consent forms, contract blanks). Without this they collapse away.
  function fillInLineNode(el, deps) {
    const { bordersOf, cs } = deps;
    if ((el.innerText || '').trim() || el.children.length) return null;
    const borders = bordersOf(el);
    // A signature rule drawn as border-top on a zero-height div is the same
    // fill-in line as a border-bottom one.
    const line =
      borders?.bottom && !borders.top && !borders.left && !borders.right
        ? borders.bottom
        : borders?.top && !borders.bottom && !borders.left && !borders.right
          ? borders.top
          : null;
    if (!line) return null;
    const rect = el.getBoundingClientRect();
    const style = cs(el);
    const display = style.display;
    const inline = display === 'inline' || display.startsWith('inline-');
    // Empty inline blanks often report ~1px height (border only) while
    // min-width still reserves the fill-in gap in the line.
    const minWidth = parseFloat(style.minWidth) || 0;
    const width = Math.max(rect.width, minWidth);
    if (inline) {
      if (width < 24 || width > 480 || rect.height > 40) return null;
    } else if (width < 40 || rect.height > 48) {
      return null;
    }
    return {
      type: 'formfield',
      mode: 'line',
      controlType: 'text',
      width: Math.round(width),
      height: Math.round(Math.max(rect.height, 16)),
      runs: [],
      borders: { bottom: line },
      inline,
    };
  }

  // Empty fully-bordered writing area (form answer boxes with min-height).
  // These are not checkboxes and must keep their authored height.
  function fillInBoxNode(el, deps) {
    const { bordersOf, cs, toHex, paintedBgHex, runStyleOf } = deps;
    if ((el.innerText || '').trim() || el.children.length) return null;
    const borders = bordersOf(el);
    if (!borders?.top || !borders?.bottom || !borders?.left || !borders?.right) {
      return null;
    }
    const style = cs(el);
    const rect = el.getBoundingClientRect();
    // An image container is not a fill-in blank: a div painted by
    // background-image (even one whose URL now 404s and shows only the
    // border ring) or shaped as a circle must stay on the screenshot path —
    // as a formfield it rendered an opaque white box over the title.
    if ((style.backgroundImage || '').includes('url(')) return null;
    const radius = String(style.borderTopLeftRadius || '');
    const radiusValue = parseFloat(radius) || 0;
    // computed border-radius keeps authored percentages ("50%") as-is
    const circle = radius.includes('%')
      ? radiusValue >= 40
      : radiusValue >= Math.min(rect.width, rect.height) * 0.4;
    if (circle) return null;
    const minH = parseFloat(style.minHeight) || 0;
    const height = Math.max(rect.height, minH);
    if (height < 28 || rect.width < 48) return null;
    // checkbox-sized squares stay on the checkbox path
    if (rect.width <= 32 && height <= 32) return null;
    const rawBg = toHex(style.backgroundColor);
    // contenteditable answer boxes show their hint via
    // :empty::before { content: attr(data-placeholder) } — resolve it so
    // the box isn't exported empty
    let placeholder = '';
    const pseudo = getComputedStyle(el, '::before');
    const pseudoContent = pseudo?.content || '';
    const attrRef = pseudoContent.match(/attr\(([-\w]+)\)/);
    if (attrRef) placeholder = el.getAttribute(attrRef[1]) || '';
    else if (pseudoContent && pseudoContent !== 'none' && pseudoContent !== 'normal') {
      placeholder = pseudoContent.replace(/^["']|["']$/g, '');
    }
    const placeholderColor = placeholder && pseudo ? toHex(pseudo.color) : null;
    return {
      type: 'formfield',
      mode: 'box',
      controlType: 'multiline',
      width: Math.round(rect.width),
      height: Math.round(height),
      value: placeholder || undefined,
      valueColor: placeholderColor || undefined,
      runs: placeholder
        ? [{
            text: placeholder,
            ...(runStyleOf ? runStyleOf(el) : {}),
            ...(placeholderColor ? { color: placeholderColor } : {}),
          }]
        : [],
      borders,
      shading: (paintedBgHex && paintedBgHex(el)) || rawBg || 'FFFFFF',
    };
  }

  function formFieldNode(el, deps) {
    const { bgHex, bordersOf, cellPaddingOf, cs, runStyleOf } = deps;
    const rect = el.getBoundingClientRect();
    cs(el);
    const value = el.value || el.getAttribute('placeholder') || '';
    const inputType = (el.getAttribute('type') || 'text').toLowerCase();
    const controlType =
      el.tagName === 'SELECT'
        ? 'dropdown'
        : el.tagName === 'TEXTAREA'
          ? 'multiline'
          : inputType === 'checkbox'
            ? 'checkbox'
            : inputType === 'radio'
              ? 'radio'
              : 'text';
    const borders = bordersOf(el);
    const hasOnlyBottomBorder =
      borders &&
      borders.bottom &&
      !borders.top &&
      !borders.left &&
      !borders.right;
    return {
      type: 'formfield',
      mode: hasOnlyBottomBorder ? 'line' : 'box',
      controlType,
      name: el.name || el.id || undefined,
      alias: el.getAttribute('aria-label') || el.name || el.id || undefined,
      value,
      checked: Boolean(el.checked),
      options:
        el.tagName === 'SELECT'
          ? [...el.options].map((option) => ({
              value: option.value,
              label: option.textContent || option.value,
              selected: option.selected,
            }))
          : undefined,
      width: Math.round(rect.width),
      height: Math.round(rect.height),
      runs: value ? [{ text: value, ...runStyleOf(el) }] : [],
      borders,
      shading: bgHex(el),
      padPx: cellPaddingOf(el),
    };
  }

  function embeddedMediaNode(el, { markForScreenshot }) {
    const rawHref =
      el.tagName === 'IFRAME'
        ? el.src
        : el.tagName === 'VIDEO' || el.tagName === 'AUDIO'
          ? el.currentSrc || el.src
          : '';
    const href = /^https?:/i.test(rawHref || '') ? rawHref : undefined;
    return {
      ...markForScreenshot(el),
      mediaType: el.tagName.toLowerCase(),
      href,
      fallbackText:
        el.getAttribute('title') ||
        el.getAttribute('aria-label') ||
        `${el.tagName === 'IFRAME' ? 'Embedded content' : el.tagName === 'VIDEO' ? 'Video' : 'Audio'}`,
    };
  }

  global.__html2docxFormsMedia = {
    embeddedMediaNode,
    fillInBoxNode,
    fillInLineNode,
    formFieldNode,
  };
})(globalThis);
