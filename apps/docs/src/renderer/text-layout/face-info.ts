/**
 * What a declared font name is, answered once for every inline layout rule.
 * The old five predicates (isCjkFontName, isCjkNamedFont, isEastAsianFontName,
 * eastAsianCoverage's name half, the Korean/Japanese variants) disagreed on
 * names like Noto Sans Mono CJK SC or Han-lettered vendor faces; this table
 * is their union, and the script split is the one the line-height and
 * fallback rules key on.
 */

/** CJK vendor and system family names (zh / ja / ko faces, Noto/Source Han, metric aliases) */
const CJK_FAMILY_RE =
  /宋|黑|楷|仿|明|雅黑|等线|simsun|nsimsun|simhei|simkai|simfang|kaiti|fangsong|songti|heiti|stsong|st(zhongsong|kai|fangsong|xihei|hupo|liti|caiyun)|xingkai|lisu|youyuan|yahei|dengxian|pingfang|mingliu|jhenghei|biaukai|dfkai|kaiu|mincho|ms (ui )?p?gothic|yu gothic|yugoth|yu mincho|ゴシック|meiryo|メイリオ|hiragino|ヒラギノ|osaka|kozuka|biz ud|游|教科書体|行書体|創英|arial unicode|apple (sd )?gothic|applemyungjo|wenquanyi|source han|noto (sans|serif)( mono)? (cjk|sc|cn|jp|tc|tw|hk|kr)\b|^misans(?: (?:thin|extralight|light|normal|regular|medium|demibold|semibold|bold|heavy|black))?$|^lxgw wenkai(?: |$)|^ar pl u(?:kai|ming)(?: |$)|^droid sans fallback(?: |$)/i

/** Korean faces by name (vendor names, hangul spellings, the KR bundles) */
export const KO_FACE_RE =
  /malgun|맑은|batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서|gulim|굴림|dotum|돋움|nanum|나눔|genoffice (sans|serif) kr|(noto|source han) (sans|serif)[^,]*\bk(r|orean)?\b/i

/** Japanese faces by name */
export const JA_FACE_RE =
  /[぀-ヿ]|mincho|meiryo|hiragino|osaka|yugoth|yu (gothic|mincho)|ms (ui )?p?(gothic|mincho)|明朝|biz ud|kozuka|小塚|^hg[ps]?(教科書|行書|正楷書|創英)/i

/** any Han, hangul, kana or fullwidth letter in the name itself */
export const CJK_LETTERED_RE = /[⺀-鿿가-힣豈-﫿＀-￯]/
export const HANGUL_LETTERED_RE = /[가-힣ᄀ-ᇿ㄰-㆏]/

const nfkc = (name: string) => name.normalize('NFKC')

/** the name denotes an East Asian face (its glyph coverage is a separate question, see line-metrics) */
export function isEastAsianFaceName(name: string): boolean {
  const n = nfkc(name)
  return CJK_FAMILY_RE.test(n) || KO_FACE_RE.test(n) || CJK_LETTERED_RE.test(n)
}

/** a known Korean family; a hangul-lettered vendor name outside the table is
 *  `isHangulLetteredName` (Word routes those by fontTable, not by family) */
export function isKoreanFaceName(name: string): boolean {
  return KO_FACE_RE.test(nfkc(name))
}

export function isHangulLetteredName(name: string): boolean {
  return HANGUL_LETTERED_RE.test(nfkc(name))
}

export function isJapaneseFaceName(name: string): boolean {
  return JA_FACE_RE.test(nfkc(name))
}

/** script a face name belongs to; ko and ja by their tables, every other East Asian name zh */
export function faceScript(name: string): 'ko' | 'ja' | 'zh' | null {
  if (isKoreanFaceName(name) || isHangulLetteredName(name)) return 'ko'
  if (isJapaneseFaceName(name)) return 'ja'
  return isEastAsianFaceName(name) ? 'zh' : null
}
