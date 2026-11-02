/** Declared-name patterns shared by line-metrics.ts and the generated line-factor table */

/** Tamil faces: Word substitutes the missing ones with cloud-font Vijaya */
export const TAMIL_FONT_RE = /tamil|vijaya|inaimathi|^tau[- ]/i

/** Segoe UI text cuts (resolve from the cloud store); the symbol/emoji/script cuts do not */
export const SEGOE_UI_TEXT_RE = /^segoe ui( (light|semilight|semibold|black|variable))?$/

export const NOTO_HAN_REGIONAL_RE =
  /^(?:noto|source han) (?:sans|serif)(?: cjk)? ?(jp|kr|sc|cn|tc|tw|hk)\b/
