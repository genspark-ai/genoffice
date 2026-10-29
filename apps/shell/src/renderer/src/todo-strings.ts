import type { Lang } from '@genoffice/i18n'

/**
 * The todo widget's words, owned by the widget itself — one entry per UI
 * language, the same shape as the manual's title tables (i18n/help/help-titles.ts)
 * rather than the shared string dictionaries, which are keyed by StringKey and
 * split per file. zh/zh-TW/en are natural copy; the rest are best-effort.
 */

export interface TodoStrings {
  /** the sidebar card's title */
  title: string
  /** the inline add input's placeholder */
  placeholder: string
  /** the one friendly line shown when the list is empty and loaded */
  empty: string
  /** the quiet action that drops finished items */
  clearDone: string
  /** the hover ✕'s accessible name */
  delete: string
}

export const TODO_STRINGS: Record<Lang, TodoStrings> = {
  zh: {
    title: '待办',
    placeholder: '添加一条待办，回车保存',
    empty: '还没有待办，想到什么直接写下来',
    clearDone: '清除已完成',
    delete: '删除',
  },
  'zh-TW': {
    title: '待辦',
    placeholder: '新增一項待辦，按 Enter 儲存',
    empty: '還沒有待辦，想到什麼直接寫下來',
    clearDone: '清除已完成',
    delete: '刪除',
  },
  en: {
    title: 'To-dos',
    placeholder: 'Add a to-do and press Enter',
    empty: 'Nothing here yet — jot down a task and press Enter',
    clearDone: 'Clear completed',
    delete: 'Delete',
  },
  ja: {
    title: 'やること',
    placeholder: 'やることを入力して Enter',
    empty: 'まだ何もありません。思いついたことを書きましょう',
    clearDone: '完了済みを消す',
    delete: '削除',
  },
  ko: {
    title: '할 일',
    placeholder: '할 일을 입력하고 Enter 키',
    empty: '아직 없어요. 떠오르는 일을 적어 두세요',
    clearDone: '완료 항목 지우기',
    delete: '삭제',
  },
  fr: {
    title: 'À faire',
    placeholder: 'Ajouter une tâche et appuyer sur Entrée',
    empty: 'Rien pour l’instant — notez une tâche et appuyez sur Entrée',
    clearDone: 'Effacer les terminées',
    delete: 'Supprimer',
  },
  de: {
    title: 'To-dos',
    placeholder: 'Aufgabe notieren und Eingabe drücken',
    empty: 'Noch nichts hier — notiere eine Aufgabe und drücke die Eingabetaste',
    clearDone: 'Erledigte entfernen',
    delete: 'Löschen',
  },
  es: {
    title: 'Pendientes',
    placeholder: 'Añade una tarea y pulsa Intro',
    empty: 'Aún no hay nada: anota una tarea y pulsa Intro',
    clearDone: 'Borrar completadas',
    delete: 'Eliminar',
  },
  th: {
    title: 'สิ่งที่ต้องทำ',
    placeholder: 'พิมพ์สิ่งที่ต้องทำแล้วกด Enter',
    empty: 'ยังไม่มีรายการ จดสิ่งที่ต้องทำแล้วกด Enter',
    clearDone: 'ล้างรายการที่เสร็จแล้ว',
    delete: 'ลบ',
  },
  id: {
    title: 'To-do',
    placeholder: 'Tulis tugas lalu tekan Enter',
    empty: 'Belum ada apa pun — catat tugas lalu tekan Enter',
    clearDone: 'Hapus yang selesai',
    delete: 'Hapus',
  },
  ru: {
    title: 'Задачи',
    placeholder: 'Добавьте задачу и нажмите Enter',
    empty: 'Пока ничего нет — запишите задачу и нажмите Enter',
    clearDone: 'Убрать выполненные',
    delete: 'Удалить',
  },
  ar: {
    title: 'مهام',
    placeholder: 'أضف مهمة واضغط Enter',
    empty: 'لا شيء بعد — دوّن مهمة واضغط Enter',
    clearDone: 'مسح المنجزة',
    delete: 'حذف',
  },
  pt: {
    title: 'Pendências',
    placeholder: 'Anote uma tarefa e pressione Enter',
    empty: 'Nada por aqui — anote uma tarefa e pressione Enter',
    clearDone: 'Limpar concluídas',
    delete: 'Excluir',
  },
  it: {
    title: 'Da fare',
    placeholder: 'Aggiungi un impegno e premi Invio',
    empty: 'Ancora niente: annota un impegno e premi Invio',
    clearDone: 'Cancella completati',
    delete: 'Elimina',
  },
  pl: {
    title: 'Do zrobienia',
    placeholder: 'Dodaj zadanie i naciśnij Enter',
    empty: 'Jeszcze nic tu nie ma — zapisz zadanie i naciśnij Enter',
    clearDone: 'Wyczyść ukończone',
    delete: 'Usuń',
  },
  cs: {
    title: 'Úkoly',
    placeholder: 'Přidejte úkol a stiskněte Enter',
    empty: 'Zatím nic — poznamenejte si úkol a stiskněte Enter',
    clearDone: 'Vymazat hotové',
    delete: 'Smazat',
  },
  nl: {
    title: "To-do's",
    placeholder: 'Voer een taak in en druk op Enter',
    empty: 'Nog niets — noteer een taak en druk op Enter',
    clearDone: 'Voltooide wissen',
    delete: 'Verwijderen',
  },
  ms: {
    title: 'To-do',
    placeholder: 'Taip tugasan dan tekan Enter',
    empty: 'Belum ada apa-apa — catat tugasan dan tekan Enter',
    clearDone: 'Kosongkan yang selesai',
    delete: 'Buang',
  },
  he: {
    title: 'מטלות',
    placeholder: 'כתבו מטלה ולחצו Enter',
    empty: 'עוד אין כלום — רשמו מטלה ולחצו Enter',
    clearDone: 'נקה את המבוצעות',
    delete: 'מחק',
  },
  hi: {
    title: 'कार्य',
    placeholder: 'कार्य लिखें और Enter दबाएँ',
    empty: 'अभी कुछ नहीं — कोई कार्य लिखें और Enter दबाएँ',
    clearDone: 'पूरे हुए हटाएँ',
    delete: 'हटाएँ',
  },
  vi: {
    title: 'Việc cần làm',
    placeholder: 'Thêm một việc rồi nhấn Enter',
    empty: 'Chưa có gì — ghi nhanh một việc rồi nhấn Enter',
    clearDone: 'Xóa việc đã xong',
    delete: 'Xóa',
  },
}

/** the widget's words in the reader's language; English if a language is missing */
export function todoStrings(lang: Lang): TodoStrings {
  return TODO_STRINGS[lang] ?? TODO_STRINGS.en
}
