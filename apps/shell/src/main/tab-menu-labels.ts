/** labels for the tab right-click menu; falls back to English for locales without a translation */
export interface TabMenuLabels {
  close: string
  closeOthers: string
  closeRight: string
  reveal: string
}

const LABELS: Record<string, TabMenuLabels> = {
  en: {
    close: 'Close tab',
    closeOthers: 'Close other tabs',
    closeRight: 'Close tabs to the right',
    reveal: 'Show in folder',
  },
  zh: {
    close: '关闭标签页',
    closeOthers: '关闭其他标签页',
    closeRight: '关闭右侧标签页',
    reveal: '在文件夹中显示',
  },
  'zh-TW': {
    close: '關閉分頁',
    closeOthers: '關閉其他分頁',
    closeRight: '關閉右側分頁',
    reveal: '在資料夾中顯示',
  },
}

export function tabMenuLabels(lang: string): TabMenuLabels {
  return LABELS[lang] ?? LABELS.en!
}
