import type { Lang } from '@genoffice/i18n'

// Accessible name for the full-window picture viewer, so the modal is announced
// by name in the editor's own language. Same pattern as the AI panel side labels.
export const IMAGE_VIEWER_TITLES: Record<Lang, string> = {
  zh: '图片查看器',
  en: 'Image viewer',
  ja: '画像ビューアー',
  ko: '이미지 뷰어',
  fr: 'Visionneuse d’images',
  de: 'Bildansicht',
  es: 'Visor de imágenes',
  th: 'ตัวดูภาพ',
  id: 'Penampil gambar',
  ru: 'Просмотр изображений',
  ar: 'عارض الصور',
  pt: 'Visualizador de imagens',
  it: 'Visualizzatore immagini',
  pl: 'Podgląd obrazów',
  cs: 'Prohlížeč obrázků',
  nl: 'Afbeeldingviewer',
  ms: 'Pemapar imej',
  he: 'מציג תמונות',
  hi: 'चित्र व्यूअर',
  'zh-TW': '圖片檢視器',
  vi: 'Trình xem ảnh',
}
