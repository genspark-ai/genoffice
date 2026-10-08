import type { Lang } from '@genoffice/i18n'

export interface AiModelPickerStrings {
  /** chip tooltip / list heading */
  title: string
  /** chip text while no provider is usable yet */
  choose: string
  /** last row: jump to Settings › AI Model */
  manage: string
}

export const AI_MODEL_PICKER_STRINGS: Record<Lang, AiModelPickerStrings> = {
  zh: { title: '模型', choose: '选择模型', manage: '管理模型…' },
  en: { title: 'Model', choose: 'Choose model', manage: 'Manage models…' },
  ja: { title: 'モデル', choose: 'モデルを選択', manage: 'モデルを管理…' },
  ko: { title: '모델', choose: '모델 선택', manage: '모델 관리…' },
  fr: { title: 'Modèle', choose: 'Choisir un modèle', manage: 'Gérer les modèles…' },
  de: { title: 'Modell', choose: 'Modell wählen', manage: 'Modelle verwalten…' },
  es: { title: 'Modelo', choose: 'Elegir modelo', manage: 'Gestionar modelos…' },
  th: { title: 'โมเดล', choose: 'เลือกโมเดล', manage: 'จัดการโมเดล…' },
  id: { title: 'Model', choose: 'Pilih model', manage: 'Kelola model…' },
  ru: { title: 'Модель', choose: 'Выбрать модель', manage: 'Управление моделями…' },
  ar: { title: 'النموذج', choose: 'اختيار النموذج', manage: 'إدارة النماذج…' },
  pt: { title: 'Modelo', choose: 'Escolher modelo', manage: 'Gerenciar modelos…' },
  it: { title: 'Modello', choose: 'Scegli modello', manage: 'Gestisci modelli…' },
  pl: { title: 'Model', choose: 'Wybierz model', manage: 'Zarządzaj modelami…' },
  cs: { title: 'Model', choose: 'Vybrat model', manage: 'Spravovat modely…' },
  nl: { title: 'Model', choose: 'Model kiezen', manage: 'Modellen beheren…' },
  ms: { title: 'Model', choose: 'Pilih model', manage: 'Urus model…' },
  he: { title: 'מודל', choose: 'בחירת מודל', manage: 'ניהול מודלים…' },
  hi: { title: 'मॉडल', choose: 'मॉडल चुनें', manage: 'मॉडल प्रबंधित करें…' },
  'zh-TW': { title: '模型', choose: '選擇模型', manage: '管理模型…' },
  vi: { title: 'Mô hình', choose: 'Chọn mô hình', manage: 'Quản lý mô hình…' },
}
