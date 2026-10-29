import type { Lang } from '@genoffice/i18n'

/** Labels for the composer slash palette (skill picker above the textarea). */
export interface AiSlashLabels {
  /** the palette's aria-label */
  title: string
  /** palette body when nothing is installed */
  noSkills: string
  /** palette body when the query matches nothing */
  noMatch: string
}

export const AI_SLASH_LABELS: Record<Lang, AiSlashLabels> = {
  zh: {
    title: '附加技能',
    noSkills: '还没有安装技能',
    noMatch: '没有匹配的技能',
  },
  en: {
    title: 'Attach a skill',
    noSkills: 'No skills installed yet',
    noMatch: 'No matching skills',
  },
  ja: {
    title: 'スキルを添付',
    noSkills: 'スキルはまだインストールされていません',
    noMatch: '一致するスキルがありません',
  },
  ko: {
    title: '스킬 첨부',
    noSkills: '설치된 스킬이 없습니다',
    noMatch: '일치하는 스킬이 없습니다',
  },
  fr: {
    title: 'Joindre une compétence',
    noSkills: 'Aucune compétence installée',
    noMatch: 'Aucune compétence correspondante',
  },
  de: {
    title: 'Skill anhängen',
    noSkills: 'Noch keine Skills installiert',
    noMatch: 'Keine passenden Skills',
  },
  es: {
    title: 'Adjuntar una habilidad',
    noSkills: 'Aún no hay habilidades instaladas',
    noMatch: 'Ninguna habilidad coincide',
  },
  th: {
    title: 'แนบทักษะ',
    noSkills: 'ยังไม่ได้ติดตั้งทักษะ',
    noMatch: 'ไม่พบทักษะที่ตรงกัน',
  },
  id: {
    title: 'Lampirkan keterampilan',
    noSkills: 'Belum ada keterampilan terpasang',
    noMatch: 'Tidak ada keterampilan yang cocok',
  },
  ru: {
    title: 'Приложить навык',
    noSkills: 'Навыки ещё не установлены',
    noMatch: 'Подходящих навыков нет',
  },
  ar: {
    title: 'إرفاق مهارة',
    noSkills: 'لم تُثبَّت أي مهارات بعد',
    noMatch: 'لا توجد مهارات مطابقة',
  },
  pt: {
    title: 'Anexar uma habilidade',
    noSkills: 'Nenhuma habilidade instalada',
    noMatch: 'Nenhuma habilidade corresponde',
  },
  it: {
    title: 'Allega una competenza',
    noSkills: 'Nessuna competenza installata',
    noMatch: 'Nessuna competenza corrispondente',
  },
  pl: {
    title: 'Dołącz umiejętność',
    noSkills: 'Brak zainstalowanych umiejętności',
    noMatch: 'Brak pasujących umiejętności',
  },
  cs: {
    title: 'Připojit dovednost',
    noSkills: 'Zatím žádné dovednosti',
    noMatch: 'Žádná odpovídající dovednost',
  },
  nl: {
    title: 'Vaardigheid toevoegen',
    noSkills: 'Nog geen vaardigheden geïnstalleerd',
    noMatch: 'Geen passende vaardigheden',
  },
  ms: {
    title: 'Lampirkan kemahiran',
    noSkills: 'Belum ada kemahiran dipasang',
    noMatch: 'Tiada kemahiran sepadan',
  },
  he: {
    title: 'לצרף מיומנות',
    noSkills: 'עדיין לא הותקנו מיומנויות',
    noMatch: 'אין מיומנויות תואמות',
  },
  hi: {
    title: 'कौशल जोड़ें',
    noSkills: 'अभी कोई कौशल इंस्टॉल नहीं',
    noMatch: 'कोई मेल खाता कौशल नहीं',
  },
  'zh-TW': {
    title: '附加技能',
    noSkills: '尚未安裝技能',
    noMatch: '沒有符合的技能',
  },
  vi: {
    title: 'Đính kèm kỹ năng',
    noSkills: 'Chưa cài kỹ năng nào',
    noMatch: 'Không có kỹ năng phù hợp',
  },
}
