import { getLang } from './locale'

const selectValues = {
  en: 'Clear Select All, then check the values to keep. Search covers the entire column.',
  ru: 'Снимите «Выбрать всё», затем отметьте нужные значения. Поиск работает по всему столбцу.',
}

export function largeFilterSelectValuesMessage(): string {
  return getLang() === 'ru' ? selectValues.ru : selectValues.en
}
