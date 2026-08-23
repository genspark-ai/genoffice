/** Classify CJK script (ja/ko/traditional-zh) by font family name, for metric substitution and draw fallback-chain selection. */
export type CjkScript = 'ja' | 'ko' | 'tc'

const JA_RE =
  /[぀-ヿ]|mincho|meiryo|hiragino|osaka|yugoth|yu (gothic|mincho)|ms (ui )?p?(gothic|mincho)|明朝|biz ud|kozuka|小塚|(sans|serif) jp\b/i
const KO_RE =
  /[가-힣ᄀ-ᇿ㄰-㆏]|malgun|batang|gulim|dotum|gungsuh|myeongjo|myungjo|nanum|korean|hangul|apple (sd )?gothic|applemyungjo|(sans|serif) kr\b/i
const TC_RE =
  /jhenghei|p?mingliu|biaukai|dfkai|kaiu|正黑|細明|標楷|蘋方|儷[黑宋]|-繁|繁體|pingfang (tc|hk)|(heiti|songti|kaiti|lihei|lisong) tc|(sans|serif) (tc|hk)\b/i

export function classifyCjkScript(family: string): CjkScript | null {
  const f = family.normalize('NFKC')
  if (JA_RE.test(f)) return 'ja'
  if (KO_RE.test(f)) return 'ko'
  if (TC_RE.test(f)) return 'tc'
  return null
}
