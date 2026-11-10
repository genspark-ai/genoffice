/** Declared-name patterns shared by line-metrics.ts and the generated line-factor table */

/** Tamil faces: Word substitutes the missing ones with cloud-font Vijaya */
export const TAMIL_FONT_RE = /tamil|vijaya|inaimathi|^tau[- ]/i

/** Thai legacy families (Windows UPC/New faces, the SIPA TH * national fonts); Tahoma and Leelawadee are Latin-metric and stay out */
export const THAI_FONT_RE =
  /^th |sarabun|angsana|cordia|browallia|eucrosia|freesia|iris ?upc|jasmine|kodchiang|lily ?upc|dillenia|\bupc$|thai/i

/** Segoe UI text cuts (resolve from the cloud store); the symbol/emoji/script cuts do not */
export const SEGOE_UI_TEXT_RE = /^segoe ui( (light|semilight|semibold|black|variable))?$/

export const NOTO_HAN_REGIONAL_RE =
  /^(?:noto|source han) (?:sans|serif)(?: cjk)? ?(jp|kr|sc|cn|tc|tw|hk)\b/

export const MONO_FONT_RE =
  /consolas|courier|menlo|monaco|cascadia|sf mono|jetbrains mono|source code|fira code|(liberation|roboto|ubuntu|dejavu sans|andale) mono|lucida console/i

/** Arabic-script names: the naskh chains carry the bundled subsets Word's Times New Roman substitute needs */
export const ARABIC_NAME_RE =
  /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]|naskh|kufi|arabic|urdu|geeza|amiri|scheherazade|lateef|harmattan|aldhabi|andalus|nastaliq|al bayan|baghdad|damascus|diwan|farisi|mishafi|nadeem|beirut|sakkal majalla|dubai|\b(?:b|xb|ir)[ -]?(?:mitra|nazanin|titr|lotus|zar|yekan|koodak|roya|badr|homa|traffic|compset)\b|irlotus/i

/** Montserrat family names of Word's cloud-font cache (weight cuts are families of their own) */
export const MONTSERRAT_FAMILY_RE =
  /^montserrat(?: (?:thin|extralight|light|medium|semibold|extrabold|black))?$/
