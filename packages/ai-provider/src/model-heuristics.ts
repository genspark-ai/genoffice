/**
 * The shipped knowledge about model *families* — which ones fix their sampling,
 * which cannot take an image, which want their reasoning echoed back.
 *
 * These are guesses, and deliberately the last layer of resolution: a user's
 * endpoint profile or model declaration always wins over them (see
 * `custom-profile.ts`). They live in their own module so the registry and the
 * profile resolver can both read them without importing each other.
 */

/**
 * Model families that fix sampling and reject a temperature field, on any
 * route — vendor API, the Genspark proxy, OpenRouter's vendor-prefixed ids,
 * or a mirror behind a custom base URL. Kimi K3 answers "only 1 is allowed";
 * OpenAI's GPT-5 and GPT-6 reasoning families reject any temperature other
 * than the default outright, and the o-series reasoning models (o1/o3/o4) likewise
 * only accept the default. Google's Gemini 3 docs strongly recommend keeping
 * the default temperature of 1.0 for the whole Gemini 3 family, since lower
 * values may cause looping or degraded reasoning, so our hard-coded 0.3
 * must not be sent there either; Gemini 4 is the same thinking lineage.
 */
export function modelHasFixedSampling(model: string): boolean {
  return /(^|\/)(kimi-k3([^\w]|$)|gpt-[5-9]([^\w]|$)|gemini-[3-9]([^\w]|$)|o1(-mini|-preview)?([^\w]|$)|o3(-mini)?([^\w]|$)|o4-mini([^\w]|$))/i.test(
    model,
  )
}

/**
 * Model ids that reject image input even under a vision-capable provider.
 * DeepSeek V4 Pro and V4 Flash are text-only; V4.1 Flash and the -vision*
 * branches take images, so they fall through and receive screenshots.
 * Ling and LongCat are the same shape: only `Ling-3.0-flash-VL` and
 * `LongCat-2.5-Preview` take images; a text-only id added to either provider
 * must be added here too.
 */
export function modelLacksVision(model: string): boolean {
  // MiniMax-M2.7 remains text-only when MiniMax-M3 enables provider vision.
  if (/(^|\/)minimax-m2\.7($|-)/i.test(model)) return true
  // Atlas Cloud serves these text-only (GET api.atlascloud.ai/v1/models gives
  // them input_modalities ["text"]), including MiniMax M3, which takes images
  // on MiniMax's own API. Matched on the full namespaced id so the direct
  // GLM and MiniMax providers keep their own behaviour. Its Kimi, Qwen and
  // GLM-5.3-Flash ids do take images and fall through.
  if (/^(?:zai-org\/glm-5\.3|minimaxai\/minimax-m(?:3|2\.5))$/i.test(model)) return true
  // Runcrate serves these text-only (GET api.runcrate.ai/v1/models gives them
  // input_modalities ["text"]; zai-org/GLM-5.3 is covered above). Its DeepSeek
  // V4.1 Flash, GLM-5.3-Flash and Kimi K2.6 ids take images and fall through.
  if (/^(?:qwen\/qwen3\.8-flash|openai\/gpt-oss-120b)$/i.test(model)) return true
  return (
    /(^|\/)deep-?seek-v4-(?:pro(?:$|-)|flash(?!-vision))/i.test(model) ||
    /(^|\/)(?:ling-(?:3\.0-flash(?!-vl)|3\.0-tiny|2\.6-1t|2\.6-flash)|ring-2\.6-1t|longcat-2\.0(?:$|-))/i.test(
      model,
    )
  )
}

/**
 * Interleaved-thinking families whose vendors want the reasoning echoed back
 * on assistant messages: MiniMax documents that stripping it degrades
 * multi-turn tool use, and DeepSeek V4 rejects tool turns without it. Gated
 * per model because other vendors may reject the unknown field. Hunyuan joins
 * them because hy4-preview ships deep thinking on by default, so its first
 * turn already carries `reasoning_content`.
 */
export function modelEchoesReasoning(model: string): boolean {
  return /(^|\/)(minimax-m|deep-?seek-(v4|flash)|hy-?[34]([^\w]|$))/i.test(model)
}
