import type { Lang } from '@genoffice/i18n'

// Shared by every editor's AI header (AiModelSwitcher): the header control's
// tooltip, the empty-list line, and the word for a codex CLI model left at
// "the CLI picks its own default".
export const AI_MODEL_SWITCHER_LABELS: Record<Lang, { switch: string; empty: string; auto: string }> = {
  zh: { switch: '切换 AI 模型', empty: '尚未配置其他供应商', auto: '自动' },
  en: { switch: 'Switch AI model', empty: 'No other providers configured', auto: 'Auto' },
  ja: { switch: 'AI モデルを切り替え', empty: '他のプロバイダーが未設定です', auto: '自動' },
  ko: { switch: 'AI 모델 전환', empty: '구성된 다른 공급자가 없습니다', auto: '자동' },
  fr: { switch: "Changer de modèle d'IA", empty: 'Aucun autre fournisseur configuré', auto: 'Auto' },
  de: { switch: 'KI-Modell wechseln', empty: 'Keine weiteren Anbieter konfiguriert', auto: 'Auto' },
  es: { switch: 'Cambiar modelo de IA', empty: 'No hay otros proveedores configurados', auto: 'Auto' },
  th: { switch: 'สลับโมเดล AI', empty: 'ยังไม่ได้ตั้งค่าผู้ให้บริการอื่น', auto: 'อัตโนมัติ' },
  id: { switch: 'Ganti model AI', empty: 'Belum ada penyedia lain yang dikonfigurasi', auto: 'Otomatis' },
  ru: { switch: 'Сменить ИИ-модель', empty: 'Другие провайдеры не настроены', auto: 'Авто' },
  ar: { switch: 'تبديل نموذج الذكاء الاصطناعي', empty: 'لم يتم تكوين مزوّدين آخرين', auto: 'تلقائي' },
  pt: { switch: 'Trocar modelo de IA', empty: 'Nenhum outro provedor configurado', auto: 'Auto' },
  it: { switch: 'Cambia modello IA', empty: 'Nessun altro provider configurato', auto: 'Auto' },
  pl: { switch: 'Przełącz model AI', empty: 'Nie skonfigurowano innych dostawców', auto: 'Auto' },
  cs: { switch: 'Přepnout AI model', empty: 'Žádní další poskytovatelé nejsou nastaveni', auto: 'Auto' },
  nl: { switch: 'AI-model wisselen', empty: 'Geen andere providers geconfigureerd', auto: 'Auto' },
  ms: { switch: 'Tukar model AI', empty: 'Tiada pembekal lain dikonfigurasi', auto: 'Auto' },
  he: { switch: 'החלפת מודל בינה מלאכותית', empty: 'אין ספקים אחרים מוגדרים', auto: 'אוטומטי' },
  hi: { switch: 'AI मॉडल बदलें', empty: 'कोई अन्य प्रदाता कॉन्फ़िगर नहीं है', auto: 'स्वतः' },
  'zh-TW': { switch: '切換 AI 模型', empty: '未設定其他供應商', auto: '自動' },
  vi: { switch: 'Chuyển mô hình AI', empty: 'Chưa cấu hình nhà cung cấp khác', auto: 'Tự động' },
}
