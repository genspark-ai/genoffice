import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { mkdtemp, open, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, join, relative, resolve, sep } from 'node:path'
import { pathToFileURL } from 'node:url'
import {
  BrowserWindow,
  WebContentsView,
  app,
  dialog,
  ipcMain,
  net,
  protocol,
  shell,
  webContents,
} from 'electron'
import type { WebContents } from 'electron'
import {
  configuredDefaultSaveDir,
  contextMenuLabels,
  fetchRemoteImage,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang } from '@genoffice/i18n'
import { generateImageTool } from '@genoffice/ai-search'
import { parseFileToText } from '@genoffice/file-parse'
import { convertHtmlToDocx } from '../../../../packages/html2docx/src'
import { atomicWriteFile } from './atomic-write'
import { ElectronBrowserDriver } from './html2docx-driver'
import {
  copyImageIntoOwnedAssets,
  discardPendingOwnedAssets,
  extractHtmlImageSources,
  pendingOwnedAssetsForDocument,
  prepareAssetsForSaveAs,
  reconcileOwnedAssets,
  renameOwnedAssetDocument,
  resolveSafeRelativeImagePath,
  resolveSourcePendingAfterSaveAs,
  rollbackPreparedSaveAsAssets,
  writeImageIntoOwnedAssets,
} from './asset-lifecycle'
import {
  ASSET_SNIFF_BYTES,
  PREVIEW_ASSET_EXTS,
  editableImageMime,
  extensionlessAssetMime,
} from './asset-mime'
import { buildPreviewDocument } from './preview-document'
import {
  assetBaseHref,
  previewUrlFor,
  registerHtmlSchemes,
  registerPreviewProtocol,
} from './preview-protocol'
import { ATTACHMENT_IMAGE_EXTS, HTML_CHANNELS } from '../shared/ipc'
import type {
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentMeta,
  AttachmentReadResult,
  ExportDocxRequest,
  ExportFormat,
  ExportPdfRequest,
  ExportResult,
  ImageData,
  SaveHtmlRequest,
  SaveHtmlResult,
  SaveMode,
} from '../shared/ipc'

const tDlg = createI18n({
  zh: {
    dlgSaveTitle: '保存 HTML 文档',
    filterHtml: 'HTML 文档',
    dlgPickImage: '选择图片',
    filterImages: '图片',
    untitledFile: '未命名文档',
    closeUnsavedMsg: '此文档有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    btnSave: '保存',
    btnDontSave: '不保存',
    btnCancel: '取消',
    dlgAddAttachment: '添加附件',
    filterSupported: '支持的文件',
    filterAll: '所有文件',
    errUnsupportedExt: '暂不支持 .{ext} 类型',
    errNotFile: '不是文件',
    errTooLarge: '超过 {mb}MB 上限',
    errImageTooLarge: '图片超过 5MB 上限',
    errUnreadable: '无法读取',
    errFileTooLarge: '文件超过大小上限',
    errParseFailed: '文件解析失败',
    errImageNoText: '图片附件不提供文本,已作为图像随用户消息发送,直接看图即可',
    errNotImage: '不是支持的图片类型',
  },
  en: {
    dlgSaveTitle: 'Save HTML Document',
    filterHtml: 'HTML Documents',
    dlgPickImage: 'Choose an Image',
    filterImages: 'Images',
    untitledFile: 'Untitled',
    closeUnsavedMsg: 'This document has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    btnSave: 'Save',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
    dlgAddAttachment: 'Add Attachments',
    filterSupported: 'Supported Files',
    filterAll: 'All Files',
    errUnsupportedExt: '.{ext} files are not supported',
    errNotFile: 'not a file',
    errTooLarge: 'exceeds the {mb}MB limit',
    errImageTooLarge: 'image exceeds the 5MB limit',
    errUnreadable: 'cannot be read',
    errFileTooLarge: 'File exceeds the size limit',
    errParseFailed: 'Failed to parse file',
    errImageNoText: 'Image attachments have no text; the image is sent along with the user message',
    errNotImage: 'not a supported image type',
  },
  ja: {
    dlgSaveTitle: 'HTML ドキュメントを保存',
    filterHtml: 'HTML ドキュメント',
    dlgPickImage: '画像を選択',
    filterImages: '画像',
    untitledFile: '無題',
    closeUnsavedMsg: 'このドキュメントに未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    btnSave: '保存',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
    dlgAddAttachment: '添付ファイルの追加',
    filterSupported: 'サポートされているファイル',
    filterAll: 'すべてのファイル',
    errUnsupportedExt: '.{ext} 形式には対応していません',
    errNotFile: 'ファイルではありません',
    errTooLarge: '{mb}MB の上限を超えています',
    errImageTooLarge: '画像が 5MB の上限を超えています',
    errUnreadable: '読み取れません',
    errFileTooLarge: 'ファイルがサイズ上限を超えています',
    errParseFailed: 'ファイルの解析に失敗しました',
    errImageNoText:
      '画像の添付ファイルはテキストを提供しません。画像としてユーザーメッセージと一緒に送信されるため、そのまま画像をご確認ください',
    errNotImage: 'サポートされていない画像形式です',
  },
  ko: {
    dlgSaveTitle: 'HTML 문서 저장',
    filterHtml: 'HTML 문서',
    dlgPickImage: '이미지 선택',
    filterImages: '이미지',
    untitledFile: '제목 없음',
    closeUnsavedMsg: '이 문서에 저장하지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    btnSave: '저장',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
    dlgAddAttachment: '첨부 파일 추가',
    filterSupported: '지원되는 파일',
    filterAll: '모든 파일',
    errUnsupportedExt: '.{ext} 형식은 지원되지 않습니다',
    errNotFile: '파일이 아닙니다',
    errTooLarge: '{mb}MB 제한을 초과했습니다',
    errImageTooLarge: '이미지가 5MB 제한을 초과했습니다',
    errUnreadable: '읽을 수 없습니다',
    errFileTooLarge: '파일이 크기 제한을 초과했습니다',
    errParseFailed: '파일을 분석하지 못했습니다',
    errImageNoText:
      '이미지 첨부 파일은 텍스트를 제공하지 않으며, 이미지 형태로 사용자 메시지와 함께 전송되므로 이미지를 직접 확인하면 됩니다',
    errNotImage: '지원되지 않는 이미지 형식입니다',
  },
  fr: {
    dlgSaveTitle: 'Enregistrer le document HTML',
    filterHtml: 'Documents HTML',
    dlgPickImage: 'Choisir une image',
    filterImages: 'Images',
    untitledFile: 'Sans titre',
    closeUnsavedMsg: 'Ce document contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    btnSave: 'Enregistrer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
    dlgAddAttachment: 'Ajouter des pièces jointes',
    filterSupported: 'Fichiers pris en charge',
    filterAll: 'Tous les fichiers',
    errUnsupportedExt: 'les fichiers .{ext} ne sont pas pris en charge',
    errNotFile: "n'est pas un fichier",
    errTooLarge: 'dépasse la limite de {mb} Mo',
    errImageTooLarge: "l'image dépasse la limite de 5 Mo",
    errUnreadable: 'lecture impossible',
    errFileTooLarge: 'Le fichier dépasse la taille maximale',
    errParseFailed: "Échec de l'analyse du fichier",
    errImageNoText:
      "Les pièces jointes image ne fournissent pas de texte ; l'image est envoyée avec le message de l'utilisateur, consultez-la directement",
    errNotImage: "type d'image non pris en charge",
  },
  de: {
    dlgSaveTitle: 'HTML-Dokument speichern',
    filterHtml: 'HTML-Dokumente',
    dlgPickImage: 'Bild auswählen',
    filterImages: 'Bilder',
    untitledFile: 'Unbenannt',
    closeUnsavedMsg: 'Dieses Dokument enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    btnSave: 'Speichern',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
    dlgAddAttachment: 'Anlagen hinzufügen',
    filterSupported: 'Unterstützte Dateien',
    filterAll: 'Alle Dateien',
    errUnsupportedExt: '.{ext}-Dateien werden nicht unterstützt',
    errNotFile: 'keine Datei',
    errTooLarge: 'überschreitet das Limit von {mb} MB',
    errImageTooLarge: 'Bild überschreitet das Limit von 5 MB',
    errUnreadable: 'kann nicht gelesen werden',
    errFileTooLarge: 'Datei überschreitet die maximale Größe',
    errParseFailed: 'Datei konnte nicht analysiert werden',
    errImageNoText:
      'Bildanlagen liefern keinen Text; das Bild wird mit der Benutzernachricht gesendet und kann direkt betrachtet werden',
    errNotImage: 'kein unterstütztes Bildformat',
  },
  es: {
    dlgSaveTitle: 'Guardar documento HTML',
    filterHtml: 'Documentos HTML',
    dlgPickImage: 'Elegir imagen',
    filterImages: 'Imágenes',
    untitledFile: 'Sin título',
    closeUnsavedMsg: 'Este documento tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Quieres guardarlos antes de cerrar?',
    btnSave: 'Guardar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
    dlgAddAttachment: 'Agregar datos adjuntos',
    filterSupported: 'Archivos compatibles',
    filterAll: 'Todos los archivos',
    errUnsupportedExt: 'los archivos .{ext} no son compatibles',
    errNotFile: 'no es un archivo',
    errTooLarge: 'supera el límite de {mb} MB',
    errImageTooLarge: 'la imagen supera el límite de 5 MB',
    errUnreadable: 'no se puede leer',
    errFileTooLarge: 'El archivo supera el tamaño máximo',
    errParseFailed: 'No se pudo analizar el archivo',
    errImageNoText:
      'Las imágenes adjuntas no proporcionan texto; la imagen se envía junto con el mensaje del usuario, puedes verla directamente',
    errNotImage: 'no es un tipo de imagen compatible',
  },
  th: {
    dlgSaveTitle: 'บันทึกเอกสาร HTML',
    filterHtml: 'เอกสาร HTML',
    dlgPickImage: 'เลือกรูปภาพ',
    filterImages: 'รูปภาพ',
    untitledFile: 'ไม่มีชื่อ',
    closeUnsavedMsg: 'เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    btnSave: 'บันทึก',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
    dlgAddAttachment: 'เพิ่มสิ่งที่แนบ',
    filterSupported: 'ไฟล์ที่รองรับ',
    filterAll: 'ไฟล์ทั้งหมด',
    errUnsupportedExt: 'ไม่รองรับไฟล์ .{ext}',
    errNotFile: 'ไม่ใช่ไฟล์',
    errTooLarge: 'เกินขีดจำกัด {mb}MB',
    errImageTooLarge: 'รูปภาพเกินขีดจำกัด 5MB',
    errUnreadable: 'ไม่สามารถอ่านได้',
    errFileTooLarge: 'ไฟล์เกินขนาดสูงสุด',
    errParseFailed: 'แยกวิเคราะห์ไฟล์ไม่สำเร็จ',
    errImageNoText:
      'สิ่งที่แนบเป็นรูปภาพไม่มีข้อความ รูปจะถูกส่งไปพร้อมข้อความของผู้ใช้ ดูรูปได้โดยตรง',
    errNotImage: 'ไม่ใช่ชนิดรูปภาพที่รองรับ',
  },
  id: {
    dlgSaveTitle: 'Simpan dokumen HTML',
    filterHtml: 'Dokumen HTML',
    dlgPickImage: 'Pilih gambar',
    filterImages: 'Gambar',
    untitledFile: 'Tanpa judul',
    closeUnsavedMsg: 'Dokumen ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    dlgAddAttachment: 'Tambahkan Lampiran',
    filterSupported: 'File yang Didukung',
    filterAll: 'Semua File',
    errUnsupportedExt: 'file .{ext} tidak didukung',
    errNotFile: 'bukan file',
    errTooLarge: 'melebihi batas {mb}MB',
    errImageTooLarge: 'gambar melebihi batas 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'File melebihi batas ukuran',
    errParseFailed: 'Gagal mengurai file',
    errImageNoText:
      'Lampiran gambar tidak menyediakan teks; gambar dikirim bersama pesan pengguna dan dapat dilihat langsung',
    errNotImage: 'bukan jenis gambar yang didukung',
  },
  ru: {
    dlgSaveTitle: 'Сохранить документ HTML',
    filterHtml: 'Документы HTML',
    dlgPickImage: 'Выберите изображение',
    filterImages: 'Изображения',
    untitledFile: 'Без названия',
    closeUnsavedMsg: 'В этом документе есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    btnSave: 'Сохранить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
    dlgAddAttachment: 'Добавить вложения',
    filterSupported: 'Поддерживаемые файлы',
    filterAll: 'Все файлы',
    errUnsupportedExt: 'файлы .{ext} не поддерживаются',
    errNotFile: 'не является файлом',
    errTooLarge: 'превышает лимит {mb} МБ',
    errImageTooLarge: 'изображение превышает лимит 5 МБ',
    errUnreadable: 'не удается прочитать',
    errFileTooLarge: 'Файл превышает максимальный размер',
    errParseFailed: 'Не удалось разобрать файл',
    errImageNoText:
      'Вложенные изображения не содержат текста; изображение отправляется вместе с сообщением пользователя, смотрите его напрямую',
    errNotImage: 'неподдерживаемый тип изображения',
  },
  ar: {
    dlgSaveTitle: 'حفظ مستند HTML',
    filterHtml: 'مستندات HTML',
    dlgPickImage: 'اختر صورة',
    filterImages: 'صور',
    untitledFile: 'بدون عنوان',
    closeUnsavedMsg: 'يحتوي هذا المستند على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    btnSave: 'حفظ',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
    dlgAddAttachment: 'إضافة مرفقات',
    filterSupported: 'الملفات المدعومة',
    filterAll: 'كل الملفات',
    errUnsupportedExt: 'ملفات .{ext} غير مدعومة',
    errNotFile: 'ليس ملفًا',
    errTooLarge: 'يتجاوز الحد {mb}MB',
    errImageTooLarge: 'الصورة تتجاوز حد 5MB',
    errUnreadable: 'تعذرت القراءة',
    errFileTooLarge: 'الملف يتجاوز الحد الأقصى للحجم',
    errParseFailed: 'فشل تحليل الملف',
    errImageNoText:
      'مرفقات الصور لا توفر نصًا؛ تُرسل الصورة مع رسالة المستخدم ويمكن الاطلاع عليها مباشرة',
    errNotImage: 'ليس نوع صورة مدعومًا',
  },
  pt: {
    dlgSaveTitle: 'Salvar documento HTML',
    filterHtml: 'Documentos HTML',
    dlgPickImage: 'Escolher imagem',
    filterImages: 'Imagens',
    untitledFile: 'Sem título',
    closeUnsavedMsg: 'Este documento tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    btnSave: 'Salvar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
    dlgAddAttachment: 'Adicionar Anexos',
    filterSupported: 'Arquivos Compatíveis',
    filterAll: 'Todos os Arquivos',
    errUnsupportedExt: 'arquivos .{ext} não são suportados',
    errNotFile: 'não é um arquivo',
    errTooLarge: 'excede o limite de {mb}MB',
    errImageTooLarge: 'a imagem excede o limite de 5MB',
    errUnreadable: 'não é possível ler',
    errFileTooLarge: 'O arquivo excede o limite de tamanho',
    errParseFailed: 'Falha ao analisar o arquivo',
    errImageNoText:
      'Anexos de imagem não fornecem texto; a imagem é enviada junto com a mensagem do usuário, basta vê-la diretamente',
    errNotImage: 'não é um tipo de imagem suportado',
  },
  it: {
    dlgSaveTitle: 'Salva documento HTML',
    filterHtml: 'Documenti HTML',
    dlgPickImage: 'Scegli immagine',
    filterImages: 'Immagini',
    untitledFile: 'Senza titolo',
    closeUnsavedMsg: 'Questo documento contiene modifiche non salvate.',
    closeUnsavedDetail: 'Vuoi salvarle prima di chiudere?',
    btnSave: 'Salva',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
    dlgAddAttachment: 'Aggiungi allegati',
    filterSupported: 'File supportati',
    filterAll: 'Tutti i file',
    errUnsupportedExt: 'i file .{ext} non sono supportati',
    errNotFile: 'non è un file',
    errTooLarge: 'supera il limite di {mb} MB',
    errImageTooLarge: "l'immagine supera il limite di 5 MB",
    errUnreadable: 'impossibile leggere',
    errFileTooLarge: 'Il file supera il limite di dimensione',
    errParseFailed: 'Impossibile analizzare il file',
    errImageNoText:
      "Gli allegati immagine non forniscono testo; l'immagine viene inviata insieme al messaggio dell'utente, basta guardarla direttamente",
    errNotImage: 'tipo di immagine non supportato',
  },
  pl: {
    dlgSaveTitle: 'Zapisz dokument HTML',
    filterHtml: 'Dokumenty HTML',
    dlgPickImage: 'Wybierz obraz',
    filterImages: 'Obrazy',
    untitledFile: 'Bez tytułu',
    closeUnsavedMsg: 'Ten dokument ma niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    btnSave: 'Zapisz',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
    dlgAddAttachment: 'Dodaj załączniki',
    filterSupported: 'Obsługiwane pliki',
    filterAll: 'Wszystkie pliki',
    errUnsupportedExt: 'pliki .{ext} nie są obsługiwane',
    errNotFile: 'to nie jest plik',
    errTooLarge: 'przekracza limit {mb} MB',
    errImageTooLarge: 'obraz przekracza limit 5 MB',
    errUnreadable: 'nie można odczytać',
    errFileTooLarge: 'Plik przekracza limit rozmiaru',
    errParseFailed: 'Nie udało się przeanalizować pliku',
    errImageNoText:
      'Załączniki graficzne nie zawierają tekstu; obraz jest wysyłany razem z wiadomością użytkownika, wystarczy na niego spojrzeć',
    errNotImage: 'nieobsługiwany typ obrazu',
  },
  cs: {
    dlgSaveTitle: 'Uložit dokument HTML',
    filterHtml: 'Dokumenty HTML',
    dlgPickImage: 'Vyberte obrázek',
    filterImages: 'Obrázky',
    untitledFile: 'Bez názvu',
    closeUnsavedMsg: 'Tento dokument obsahuje neuložené změny.',
    closeUnsavedDetail: 'Chcete je před zavřením uložit?',
    btnSave: 'Uložit',
    btnDontSave: 'Neukládat',
    btnCancel: 'Zrušit',
    dlgAddAttachment: 'Přidat přílohy',
    filterSupported: 'Podporované soubory',
    filterAll: 'Všechny soubory',
    errUnsupportedExt: 'soubory .{ext} nejsou podporovány',
    errNotFile: 'není soubor',
    errTooLarge: 'překračuje limit {mb} MB',
    errImageTooLarge: 'obrázek překračuje limit 5 MB',
    errUnreadable: 'nelze přečíst',
    errFileTooLarge: 'Soubor překračuje limit velikosti',
    errParseFailed: 'Soubor se nepodařilo zpracovat',
    errImageNoText:
      'Obrázkové přílohy neobsahují text; obrázek se odesílá spolu se zprávou uživatele',
    errNotImage: 'nepodporovaný typ obrázku',
  },
  nl: {
    dlgSaveTitle: 'HTML-document opslaan',
    filterHtml: 'HTML-documenten',
    dlgPickImage: 'Kies een afbeelding',
    filterImages: 'Afbeeldingen',
    untitledFile: 'Naamloos',
    closeUnsavedMsg: 'Dit document bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    btnSave: 'Opslaan',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
    dlgAddAttachment: 'Bijlagen toevoegen',
    filterSupported: 'Ondersteunde bestanden',
    filterAll: 'Alle bestanden',
    errUnsupportedExt: '.{ext}-bestanden worden niet ondersteund',
    errNotFile: 'geen bestand',
    errTooLarge: 'overschrijdt de limiet van {mb} MB',
    errImageTooLarge: 'afbeelding overschrijdt de limiet van 5 MB',
    errUnreadable: 'kan niet worden gelezen',
    errFileTooLarge: 'Bestand overschrijdt de maximale grootte',
    errParseFailed: 'Kan bestand niet parseren',
    errImageNoText:
      'Afbeeldingsbijlagen bevatten geen tekst; de afbeelding wordt samen met het gebruikersbericht verzonden en kan direct worden bekeken',
    errNotImage: 'geen ondersteund afbeeldingstype',
  },
  ms: {
    dlgSaveTitle: 'Simpan dokumen HTML',
    filterHtml: 'Dokumen HTML',
    dlgPickImage: 'Pilih imej',
    filterImages: 'Imej',
    untitledFile: 'Tanpa tajuk',
    closeUnsavedMsg: 'Dokumen ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    dlgAddAttachment: 'Tambah Lampiran',
    filterSupported: 'Fail yang Disokong',
    filterAll: 'Semua Fail',
    errUnsupportedExt: 'fail .{ext} tidak disokong',
    errNotFile: 'bukan fail',
    errTooLarge: 'melebihi had {mb}MB',
    errImageTooLarge: 'imej melebihi had 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'Fail melebihi had saiz',
    errParseFailed: 'Gagal menghurai fail',
    errImageNoText:
      'Lampiran imej tidak menyediakan teks; imej dihantar bersama mesej pengguna dan boleh dilihat terus',
    errNotImage: 'bukan jenis imej yang disokong',
  },
  he: {
    dlgSaveTitle: 'שמירת מסמך HTML',
    filterHtml: 'מסמכי HTML',
    dlgPickImage: 'בחרו תמונה',
    filterImages: 'תמונות',
    untitledFile: 'ללא שם',
    closeUnsavedMsg: 'במסמך הזה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    btnSave: 'שמירה',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
    dlgAddAttachment: 'הוספת קבצים מצורפים',
    filterSupported: 'קבצים נתמכים',
    filterAll: 'כל הקבצים',
    errUnsupportedExt: 'קובצי .{ext} אינם נתמכים',
    errNotFile: 'אינו קובץ',
    errTooLarge: 'חורג מהמגבלה של {mb}MB',
    errImageTooLarge: 'התמונה חורגת מהמגבלה של 5MB',
    errUnreadable: 'לא ניתן לקרוא',
    errFileTooLarge: 'הקובץ חורג ממגבלת הגודל',
    errParseFailed: 'ניתוח הקובץ נכשל',
    errImageNoText:
      'קבצים מצורפים מסוג תמונה אינם מספקים טקסט; התמונה נשלחת יחד עם הודעת המשתמש וניתן לצפות בה ישירות',
    errNotImage: 'סוג תמונה שאינו נתמך',
  },
  hi: {
    dlgSaveTitle: 'HTML दस्तावेज़ सहेजें',
    filterHtml: 'HTML दस्तावेज़',
    dlgPickImage: 'छवि चुनें',
    filterImages: 'छवियाँ',
    untitledFile: 'शीर्षकहीन',
    closeUnsavedMsg: 'इस दस्तावेज़ में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    btnSave: 'सहेजें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
    dlgAddAttachment: 'अनुलग्नक जोड़ें',
    filterSupported: 'समर्थित फ़ाइलें',
    filterAll: 'सभी फ़ाइलें',
    errUnsupportedExt: '.{ext} फ़ाइलें समर्थित नहीं हैं',
    errNotFile: 'फ़ाइल नहीं है',
    errTooLarge: '{mb}MB की सीमा से अधिक है',
    errImageTooLarge: 'छवि 5MB की सीमा से अधिक है',
    errUnreadable: 'पढ़ा नहीं जा सकता',
    errFileTooLarge: 'फ़ाइल आकार सीमा से अधिक है',
    errParseFailed: 'फ़ाइल पार्स करने में विफल',
    errImageNoText:
      'छवि अनुलग्नक टेक्स्ट प्रदान नहीं करते; छवि उपयोगकर्ता संदेश के साथ भेजी जाती है, उसे सीधे देखें',
    errNotImage: 'समर्थित छवि प्रकार नहीं है',
  },
  'zh-TW': {
    dlgSaveTitle: '儲存 HTML 文件',
    filterHtml: 'HTML 文件',
    dlgPickImage: '選擇圖片',
    filterImages: '圖片',
    untitledFile: '未命名文件',
    closeUnsavedMsg: '此文件有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否儲存？',
    btnSave: '儲存',
    btnDontSave: '不儲存',
    btnCancel: '取消',
    dlgAddAttachment: '新增附件',
    filterSupported: '支援的檔案',
    filterAll: '所有檔案',
    errUnsupportedExt: '暫不支援 .{ext} 類型',
    errNotFile: '不是檔案',
    errTooLarge: '超過 {mb}MB 上限',
    errImageTooLarge: '圖片超過 5MB 上限',
    errUnreadable: '無法讀取',
    errFileTooLarge: '檔案超過大小上限',
    errParseFailed: '檔案解析失敗',
    errImageNoText: '圖片附件不提供文字,已作為影像隨使用者訊息傳送,直接看圖即可',
    errNotImage: '不是支援的圖片類型',
  },
})
type DlgKey =
  | 'dlgSaveTitle'
  | 'filterHtml'
  | 'dlgPickImage'
  | 'filterImages'
  | 'untitledFile'
  | 'closeUnsavedMsg'
  | 'closeUnsavedDetail'
  | 'btnSave'
  | 'btnDontSave'
  | 'btnCancel'
  | 'dlgAddAttachment'
  | 'filterSupported'
  | 'filterAll'
  | 'errUnsupportedExt'
  | 'errNotFile'
  | 'errTooLarge'
  | 'errImageTooLarge'
  | 'errUnreadable'
  | 'errFileTooLarge'
  | 'errParseFailed'
  | 'errImageNoText'
  | 'errNotImage'
const tm = (key: DlgKey, vars?: Record<string, string | number>) => tDlg(getUiLang(), key, vars)

// ---- chat attachments: local files parsed for the agent (same contract as the docs panel) ----

const ATTACHMENT_MAX_BYTES = 50 * 1024 * 1024
const TEXT_EXTS = new Set([
  'txt',
  'md',
  'markdown',
  'csv',
  'tsv',
  'json',
  'yaml',
  'yml',
  'xml',
  'html',
  'htm',
  'log',
  'js',
  'ts',
  'tsx',
  'jsx',
  'py',
  'java',
  'c',
  'h',
  'cpp',
  'go',
  'rs',
  'rb',
  'sh',
  'sql',
  'css',
])
/** office/pdf formats get text extracted via @genoffice/file-parse; images skip extraction and go multimodal */
const ATTACHMENT_EXTS = new Set([
  ...TEXT_EXTS,
  'doc',
  'docx',
  'pdf',
  'pptx',
  'ppt',
  'xlsx',
  'xlsm',
  'xls',
  ...ATTACHMENT_IMAGE_EXTS,
])
const ATTACHMENT_IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
}
const ATTACHMENT_IMAGE_MAX_BYTES = 5 * 1024 * 1024

/** extracted text cache keyed by path; invalidated by mtime+size */
const attachmentTextCache = new Map<string, { stamp: string; text: string }>()

function statAttachment(filePath: string): { meta?: AttachmentMeta; error?: string } {
  const name = basename(filePath)
  const ext = name.split('.').pop()?.toLowerCase() ?? ''
  if (!ATTACHMENT_EXTS.has(ext)) return { error: `${name}: ${tm('errUnsupportedExt', { ext })}` }
  try {
    const stat = statSync(filePath)
    if (!stat.isFile()) return { error: `${name}: ${tm('errNotFile')}` }
    if (stat.size > ATTACHMENT_MAX_BYTES) {
      return {
        error: `${name}: ${tm('errTooLarge', { mb: Math.round(ATTACHMENT_MAX_BYTES / 1024 / 1024) })}`,
      }
    }
    if (ATTACHMENT_IMAGE_EXTS.has(ext) && stat.size > ATTACHMENT_IMAGE_MAX_BYTES) {
      return { error: `${name}: ${tm('errImageTooLarge')}` }
    }
    return { meta: { path: filePath, name, ext, sizeBytes: stat.size } }
  } catch {
    return { error: `${name}: ${tm('errUnreadable')}` }
  }
}

function collectAttachments(paths: string[]): AttachmentAddResult {
  const accepted: AttachmentMeta[] = []
  const rejected: string[] = []
  for (const p of paths) {
    const { meta, error } = statAttachment(p)
    if (meta) accepted.push(meta)
    else if (error) rejected.push(error)
  }
  return { accepted, rejected }
}

let pastedImageSeq = 0
let pastedDirPruned = false

/** drop pasted-image temp files older than 7 days (once per app run) */
function prunePastedImages(dir: string): void {
  if (pastedDirPruned) return
  pastedDirPruned = true
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000
  try {
    for (const name of readdirSync(dir)) {
      const p = join(dir, name)
      try {
        if (statSync(p).mtimeMs < cutoff) unlinkSync(p)
      } catch {
        // another tab may have removed it already
      }
    }
  } catch {
    // directory may not exist yet
  }
}

/** clipboard-pasted image bytes → temp file (shared with the docs panel), null for non-images or empty data */
function savePastedImage(data: unknown, ext: unknown): string | null {
  const cleanExt = typeof ext === 'string' ? ext.toLowerCase() : ''
  if (!ATTACHMENT_IMAGE_EXTS.has(cleanExt)) return null
  const bytes =
    data instanceof ArrayBuffer
      ? Buffer.from(data)
      : ArrayBuffer.isView(data)
        ? Buffer.from(data.buffer, data.byteOffset, data.byteLength)
        : null
  if (!bytes || bytes.byteLength === 0) return null
  const dir = join(app.getPath('temp'), 'genoffice-pasted')
  mkdirSync(dir, { recursive: true })
  prunePastedImages(dir)
  const stamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '-')
  const filePath = join(dir, `pasted-${stamp}-${++pastedImageSeq}.${cleanExt}`)
  writeFileSync(filePath, bytes)
  return filePath
}

async function extractAttachmentText(filePath: string): Promise<string> {
  const stat = statSync(filePath)
  const stamp = `${stat.mtimeMs}:${stat.size}`
  const cached = attachmentTextCache.get(filePath)
  if (cached && cached.stamp === stamp) return cached.text
  if (stat.size > ATTACHMENT_MAX_BYTES) throw new Error(tm('errFileTooLarge'))
  const parsed = await parseFileToText(filePath)
  if (!parsed.ok || parsed.kind !== 'text' || parsed.text == null) {
    throw new Error(parsed.error ?? tm('errParseFailed'))
  }
  attachmentTextCache.set(filePath, { stamp, text: parsed.text })
  if (attachmentTextCache.size > 8) {
    const oldest = attachmentTextCache.keys().next().value
    if (oldest) attachmentTextCache.delete(oldest)
  }
  return parsed.text
}

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
  /** Shell router used to open exported PDFs in a new GenOffice tab. */
  openGeneratedPath?: (path: string) => boolean
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configureHtmlRuntime(paths: RuntimePaths): void {
  runtime = paths
}

export { registerHtmlSchemes }

/** After a successful Html → PDF export: open the file in a PDF tab (shell)
 * or reveal it in the folder (standalone). Tab-opening failure must not
 * report the export itself as failed — the file is already persisted. */
function openExportedPdf(path: string): void {
  try {
    if (runtime.openGeneratedPath?.(path)) return
  } catch (err) {
    console.warn('[html] Failed to open exported PDF:', err)
  }
  shell.showItemInFolder(path)
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount.
 * Kept until the view is destroyed so a reload (View > Reload) consumes it again. */
const openPathByWc = new Map<number, string>()
/** File paths granted to each view — readFile/save only allow these */
const allowedByWc = new Map<number, Set<string>>()
/** Current save target per view; absent = untitled document */
const savePathByWc = new Map<number, string>()
/** Unsaved-changes flags mirrored from the renderer; drives the save prompt before closing a tab/window */
const dirtyByWc = new Set<number>()
/** Latest buffer text pushed by each renderer; served by html-preview:// to the preview iframe */
const previewTextByWc = new Map<number, string>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
/** Resolvers for menu-triggered saves, resolved when the renderer's save invoke completes */
const saveWaiters = new Map<number, (ok: boolean) => void>()

/** Fired after a save lands on a NEW path (untitled first save / Save As) — the shell syncs tab title, recents, projects */
let fileSavedHook: ((wc: WebContents, path: string) => void) | null = null

export function setHtmlFileSavedHook(hook: (wc: WebContents, path: string) => void): void {
  fileSavedHook = hook
}

/** An untitled document got a provisional name from the user's first AI request — the shell titles its tab */
let provisionalTitleHook: ((wc: WebContents, title: string) => void) | null = null

export function setHtmlProvisionalTitleHook(hook: (wc: WebContents, title: string) => void): void {
  provisionalTitleHook = hook
}

/** After a Word export the shell opens the new .docx in a docs tab; standalone reveals it */
let docxExportedHook: ((path: string) => void) | null = null
/** Before the .docx is written: the shell closes a docs tab already showing that path
 * (its unsaved-changes prompt applies); false = the user kept it, so the export is dropped */
let docxExportPrepareHook: ((path: string) => Promise<boolean>) | null = null

export function setHtmlDocxExportedHook(hook: (path: string) => void): void {
  docxExportedHook = hook
}

export function setHtmlDocxExportPrepareHook(hook: (path: string) => Promise<boolean>): void {
  docxExportPrepareHook = hook
}

function openExportedDocx(path: string): void {
  try {
    if (docxExportedHook) {
      docxExportedHook(path)
      return
    }
  } catch (err) {
    console.warn('[html] Failed to open exported Word file:', err)
  }
  shell.showItemInFolder(path)
}

function exportFileName(suggested: unknown): string {
  return (
    String(suggested || tm('untitledFile'))
      .replace(/[/\\:*?"<>|]/g, '_')
      .slice(0, 80)
      .trim() || tm('untitledFile')
  )
}

export interface HtmlPresentHooks {
  /** cover the tab strip with the presenting tab's view (shell tab mode) */
  setBleed?: (wc: WebContents, on: boolean) => void
  /** window hosting a view when BrowserWindow.fromWebContents cannot tell (shell WebContentsView) */
  hostWindow?: (wc: WebContents) => BrowserWindow | null
  /** open a chrome-free tab presenting the owner view's preview; false → a window is opened instead */
  openTab?: (owner: WebContents, title: string) => boolean
  /** close the tab hosting a present view; false → not a tab of this shell */
  closeTab?: (wc: WebContents) => boolean
}
let presentHooks: HtmlPresentHooks = {}
export function setHtmlPresentHooks(hooks: HtmlPresentHooks): void {
  presentHooks = hooks
}
/** present views → the editing view whose preview they show */
const presentOwnerByWc = new Map<number, number>()
/** standalone Present windows; unreferenced BrowserWindows may be garbage-collected */
const presentWindows = new Set<BrowserWindow>()

/** The owner view is gone, so is its preview: close every present view showing it */
function closePresentViewsOf(ownerWcId: number): void {
  for (const [id, owner] of presentOwnerByWc) {
    if (owner !== ownerWcId) continue
    presentOwnerByWc.delete(id)
    const wc = webContents.fromId(id)
    if (!wc || wc.isDestroyed()) continue
    if (presentHooks.closeTab?.(wc)) continue
    const win = BrowserWindow.fromWebContents(wc)
    if (win && !win.isDestroyed()) win.close()
    else wc.close()
  }
}

/** A4 at 96dpi; html2docx re-measures at the authored width itself when the page asks for more. */
const HTML2DOCX_VIEWPORT = { width: 794, height: 1123, deviceScaleFactor: 2 }

/** Print the document in a hidden script-free window (sheets-style). Relative assets
 * resolve through html-asset:// against the document's folder, exactly as in the preview. */
async function renderPrintPdf(
  html: string,
  docPath: string | undefined,
  workDir: string,
): Promise<Buffer> {
  const base = docPath ? assetBaseHref(dirname(docPath)) : null
  const htmlPath = join(workDir, 'print.html')
  await writeFile(htmlPath, buildPreviewDocument(html, base), 'utf8')
  const printWin = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await printWin.loadFile(htmlPath)
    return await printWin.webContents.printToPDF({
      pageSize: 'A4',
      printBackground: true,
      margins: { top: 0.6, bottom: 0.6, left: 0.6, right: 0.6 },
    })
  } finally {
    printWin.destroy()
  }
}

/** Shell menu export entry: ask the renderer to serialize and run the export flow */
export function sendHtmlExportRequest(contents: WebContents, format: ExportFormat): void {
  if (contents.isDestroyed() || presentOwnerByWc.has(contents.id)) return
  contents.send(HTML_CHANNELS.exportRequest, format)
}

/** Shell menu Print: ask the renderer to build the print HTML and open the system dialog */
export function sendHtmlPrintRequest(contents: WebContents): void {
  if (contents.isDestroyed() || presentOwnerByWc.has(contents.id)) return
  contents.send(HTML_CHANNELS.printRequest)
}

export function htmlIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

export function htmlFilePath(webContentsId: number): string | undefined {
  return savePathByWc.get(webContentsId)
}

/** The file was renamed on disk — re-grant the new path and tell the renderer */
export function htmlFileRenamed(contents: WebContents, oldPath: string, newPath: string): void {
  const wcId = contents.id
  if (savePathByWc.get(wcId) === oldPath) savePathByWc.set(wcId, newPath)
  if (openPathByWc.get(wcId) === oldPath) openPathByWc.set(wcId, newPath)
  const allowed = allowedByWc.get(wcId)
  if (allowed?.has(oldPath)) allowed.add(newPath)
  void renameOwnedAssetDocument(oldPath, newPath).catch((error) => {
    console.warn('[html] asset manifest rename sync failed:', error)
  })
  if (!contents.isDestroyed()) contents.send(HTML_CHANNELS.fileRenamed, newPath)
}

/**
 * Close guard: true means proceed with closing. Clean → true; dirty →
 * Save / Don't Save / Cancel. On Save, ask the renderer to serialize + write
 * and await the result; a canceled untitled-save dialog keeps the tab open.
 */
export async function requestHtmlClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!dirtyByWc.has(contents.id) || contents.isDestroyed()) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('btnSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) {
    const documentPath = savePathByWc.get(contents.id)
    if (documentPath) {
      const discarded = await discardPendingOwnedAssets(documentPath)
      if (discarded.errors.length > 0) {
        console.warn('[html] pending asset discard incomplete:', discarded.errors)
      }
    }
    return true
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(HTML_CHANNELS.closeSaveRequest)
  })
}

/** Menu Save / Save As: ask the renderer to serialize and save; clean views resolve true immediately on plain save */
export function requestHtmlSave(contents: WebContents, mode: SaveMode): Promise<boolean> {
  if (contents.isDestroyed()) return Promise.resolve(false)
  if (presentOwnerByWc.has(contents.id)) return Promise.resolve(true)
  if (mode === 'save' && !dirtyByWc.has(contents.id) && savePathByWc.has(contents.id)) {
    return Promise.resolve(true)
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      saveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    saveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(HTML_CHANNELS.saveRequest, mode)
  })
}

async function writeTextAtomic(path: string, text: string): Promise<void> {
  await atomicWriteFile(path, Buffer.from(text, 'utf8'))
}

function fileNameBase(name: string | undefined): string {
  return (name ?? '')
    .replace(/[/\\:*?"<>|]/g, '_')
    .slice(0, 80)
    .trim()
}

async function resolveSaveTarget(
  e: Electron.IpcMainInvokeEvent,
  mode: SaveMode,
  suggestedName?: string,
  defaultName?: string,
): Promise<string | null | 'canceled'> {
  const current = savePathByWc.get(e.sender.id)
  if (mode === 'save' && current) return current
  // AI auto-naming: silent first save of an untitled document
  if (mode === 'save' && !current && suggestedName) {
    const base = fileNameBase(suggestedName)
    if (base) {
      const dir = configuredDefaultSaveDir(app)
      let target = join(dir, `${base}.html`)
      for (let n = 1; existsSync(target); n++) target = join(dir, `${base}-${n}.html`)
      return target
    }
  }
  const win =
    BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
  const defaultPath = current
    ? join(dirname(current), basename(current))
    : join(configuredDefaultSaveDir(app), `${fileNameBase(defaultName) || tm('untitledFile')}.html`)
  const picked = await showSaveDialogWithMemory(dialog, win, {
    title: tm('dlgSaveTitle'),
    defaultPath,
    filters: [{ name: tm('filterHtml'), extensions: ['html', 'htm'] }],
  })
  if (picked.canceled || !picked.filePath) return 'canceled'
  return picked.filePath
}

/**
 * Serves authored image paths to the editor DOM. A plain file:// <img> URL is
 * blocked whenever the renderer page is served over http (dev server), so the
 * renderer resolves images to html-asset:// instead. Only files inside an open
 * document's directory are served: by extension, or for extensionless
 * "Save page as, complete" assets by signature / requesting slot.
 */
function registerImageProtocol(): void {
  protocol.handle('html-asset', async (request) => {
    let target: string
    try {
      target = decodeURIComponent(new URL(request.url).pathname)
    } catch {
      return new Response(null, { status: 400 })
    }
    if (/^\/[a-zA-Z]:\//.test(target)) target = target.slice(1)
    target = resolve(target)
    const knownExt = PREVIEW_ASSET_EXTS.has(extname(target).toLowerCase())
    if (!existsSync(target)) return new Response(null, { status: 404 })
    let inDocDir = false
    for (const doc of new Set([...openPathByWc.values(), ...savePathByWc.values()])) {
      const dir = resolve(dirname(doc))
      if (target === dir || !target.startsWith(dir + sep)) continue
      if (await resolveSafeRelativeImagePath(doc, relative(dir, target))) {
        inDocDir = true
        break
      }
    }
    if (!inDocDir) return new Response(null, { status: 403 })
    let mime: string | null = null
    if (!knownExt) {
      mime = extensionlessAssetMime(target, await readHead(target), request.headers.get('accept'))
      if (!mime) return new Response(null, { status: 404 })
    }
    const res = await net.fetch(pathToFileURL(target).toString())
    if (!mime || !res.ok) return res
    return new Response(res.body, { status: 200, headers: { 'Content-Type': mime } })
  })
}

async function readHead(target: string): Promise<Uint8Array> {
  const handle = await open(target, 'r')
  try {
    const buffer = Buffer.alloc(ASSET_SNIFF_BYTES)
    const { bytesRead } = await handle.read(buffer, 0, ASSET_SNIFF_BYTES, 0)
    return buffer.subarray(0, bytesRead)
  } finally {
    await handle.close()
  }
}

let ipcRegistered = false

function registerHtmlIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  registerImageProtocol()
  registerPreviewProtocol((wcId) => {
    const text = previewTextByWc.get(wcId)
    if (text === undefined) return null
    const doc = savePathByWc.get(wcId)
    return { text, baseHref: doc ? assetBaseHref(dirname(doc)) : null }
  })

  ipcMain.handle(HTML_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.on(HTML_CHANNELS.previewUpdate, (e, text: unknown) => {
    if (typeof text === 'string') previewTextByWc.set(e.sender.id, text)
  })

  ipcMain.handle(HTML_CHANNELS.previewInfo, (e) => ({
    url: previewUrlFor(presentOwnerByWc.get(e.sender.id) ?? e.sender.id),
  }))

  // Same shape as the slides show: the renderer asks for the screen in one call so the
  // macOS snap skips the Space animation; HTML fullscreen is left to the renderer elsewhere.
  let presentFsRelease: ReturnType<typeof setTimeout> | null = null
  ipcMain.handle(HTML_CHANNELS.presentFullScreen, (e, on: unknown) => {
    const wc = e.sender
    const win = BrowserWindow.fromWebContents(wc) ?? presentHooks.hostWindow?.(wc) ?? null
    if (!win || win.isDestroyed()) return
    if (presentFsRelease) {
      clearTimeout(presentFsRelease)
      presentFsRelease = null
    }
    if (on === true) {
      presentHooks.setBleed?.(wc, true)
      if (process.platform === 'darwin' && !win.isFullScreen()) {
        win.setFullScreenable(false)
        if (!win.isSimpleFullScreen()) win.setSimpleFullScreen(true)
      }
      // the snap can hand the first responder to the shell chrome; Esc must keep landing in the tab
      wc.focus()
      setTimeout(() => {
        if (!wc.isDestroyed()) wc.focus()
      }, 50)
      return
    }
    presentFsRelease = setTimeout(() => {
      presentFsRelease = null
      if (!wc.isDestroyed()) presentHooks.setBleed?.(wc, false)
      if (win.isDestroyed()) return
      if (process.platform === 'darwin') {
        if (win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
        win.setFullScreenable(true)
      }
    }, 150)
  })

  ipcMain.handle(HTML_CHANNELS.presentNewTab, (e, title: unknown) => {
    const label = typeof title === 'string' ? title : ''
    if (presentHooks.openTab?.(e.sender, label)) return true
    const parent = BrowserWindow.fromWebContents(e.sender) ?? presentHooks.hostWindow?.(e.sender)
    const bounds =
      parent && !parent.isDestroyed() ? parent.getContentBounds() : { width: 1200, height: 850 }
    const win = new BrowserWindow({
      width: bounds.width,
      height: bounds.height,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    presentWindows.add(win)
    win.once('closed', () => presentWindows.delete(win))
    bindPresentView(win.webContents, e.sender.id, label)
    return true
  })

  ipcMain.handle(HTML_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('html: path not granted to this view')
    }
    return await readFile(path, 'utf8')
  })

  ipcMain.handle(
    HTML_CHANNELS.save,
    async (e, request: SaveHtmlRequest): Promise<SaveHtmlResult> => {
      const waiter = saveWaiters.get(e.sender.id)
      saveWaiters.delete(e.sender.id)
      const done = (result: SaveHtmlResult): SaveHtmlResult => {
        waiter?.(result.ok && !('canceled' in result))
        return result
      }
      if (typeof request?.text !== 'string') {
        return done({ ok: false, error: 'html: bad save request' })
      }
      if (
        request.imageSources !== undefined &&
        (!Array.isArray(request.imageSources) ||
          request.imageSources.some((source) => typeof source !== 'string'))
      ) {
        return done({ ok: false, error: 'html: bad image references' })
      }
      const mode: SaveMode = request.mode === 'saveAs' ? 'saveAs' : 'save'
      const pathAtRequest = savePathByWc.get(e.sender.id)
      const pendingAtRequest = pathAtRequest
        ? await pendingOwnedAssetsForDocument(pathAtRequest)
        : []
      try {
        const suggestedName =
          typeof request.suggestedName === 'string' ? request.suggestedName : undefined
        const defaultName =
          typeof request.defaultName === 'string' ? request.defaultName : undefined
        const target = await resolveSaveTarget(e, mode, suggestedName, defaultName)
        if (target === 'canceled') return done({ ok: true, canceled: true })
        if (!target) return done({ ok: false, error: 'html: no save target' })
        const currentPath = pathAtRequest
        const isNewPath = currentPath !== target
        const imageSources = [...(request.imageSources ?? [])]
        const knownImageSources = new Set(imageSources)
        for (const source of extractHtmlImageSources(request.text)) {
          if (knownImageSources.has(source)) continue
          knownImageSources.add(source)
          imageSources.push(source)
        }
        const prepared =
          currentPath && resolve(dirname(currentPath)) !== resolve(dirname(target))
            ? await prepareAssetsForSaveAs(currentPath, target, request.text, imageSources)
            : null
        const textToWrite = prepared?.text ?? request.text
        const savedImageSources = prepared?.imageSources ?? imageSources
        try {
          await writeTextAtomic(target, textToWrite)
        } catch (error) {
          if (prepared) await rollbackPreparedSaveAsAssets(prepared).catch(() => {})
          throw error
        }
        savePathByWc.set(e.sender.id, target)
        // keep the reload path in sync — a stale openPathByWc would make a
        // reloaded renderer load the OLD file and then save it over the new one
        openPathByWc.set(e.sender.id, target)
        const allowed = allowedByWc.get(e.sender.id) ?? new Set<string>()
        allowed.add(target)
        allowedByWc.set(e.sender.id, allowed)
        dirtyByWc.delete(e.sender.id)
        const pendingNames = prepared
          ? prepared.created.map((record) => record.name)
          : currentPath && resolve(currentPath) === resolve(target)
            ? pendingAtRequest
            : []
        const reconciled = await reconcileOwnedAssets(target, savedImageSources, { pendingNames })
        if (reconciled.errors.length > 0) {
          console.warn('[html] asset reconciliation incomplete:', reconciled.errors)
        }
        if (mode === 'saveAs' && currentPath && resolve(currentPath) !== resolve(target)) {
          const sourceResolved = await resolveSourcePendingAfterSaveAs(
            currentPath,
            pendingAtRequest,
          )
          if (sourceResolved.errors.length > 0) {
            console.warn('[html] source asset reconciliation incomplete:', sourceResolved.errors)
          }
        }
        if (isNewPath) fileSavedHook?.(e.sender, target)
        return done({
          ok: true,
          path: target,
          ...(prepared?.rewrites.length ? { imageRewrites: prepared.rewrites } : {}),
        })
      } catch (err) {
        return done({ ok: false, error: err instanceof Error ? err.message : String(err) })
      }
    },
  )

  ipcMain.handle(HTML_CHANNELS.filesPick, async (e): Promise<AttachmentAddResult | null> => {
    const win =
      BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
    const picked = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgAddAttachment'),
      filters: [
        { name: tm('filterSupported'), extensions: [...ATTACHMENT_EXTS] },
        { name: tm('filterAll'), extensions: ['*'] },
      ],
      properties: ['openFile', 'multiSelections'],
    })
    if (picked.canceled || picked.filePaths.length === 0) return null
    return collectAttachments(picked.filePaths)
  })

  ipcMain.handle(HTML_CHANNELS.filesAdd, (_e, paths: unknown) =>
    collectAttachments(Array.isArray(paths) ? paths.filter((p) => typeof p === 'string') : []),
  )

  ipcMain.handle(
    HTML_CHANNELS.filesAddPastedImage,
    (_e, data: unknown, ext: unknown): AttachmentAddResult => {
      const filePath = savePastedImage(data, ext)
      return filePath
        ? collectAttachments([filePath])
        : { accepted: [], rejected: [tm('errNotImage')] }
    },
  )

  ipcMain.handle(
    HTML_CHANNELS.filesRead,
    async (
      _e,
      filePath: string,
      offset: number,
      maxChars: number,
    ): Promise<AttachmentReadResult> => {
      const name = basename(filePath)
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      if (!ATTACHMENT_EXTS.has(ext)) return { ok: false, error: tm('errUnsupportedExt', { ext }) }
      if (ATTACHMENT_IMAGE_EXTS.has(ext)) return { ok: false, error: tm('errImageNoText') }
      try {
        const text = await extractAttachmentText(filePath)
        const start = Math.max(0, Math.floor(offset) || 0)
        const size = Math.min(Math.max(1, Math.floor(maxChars) || 1), 48_000)
        return {
          ok: true,
          name,
          totalChars: text.length,
          offset: start,
          text: text.slice(start, start + size),
        }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(HTML_CHANNELS.filesReadImage, (_e, filePath: string): AttachmentImageResult => {
    const name = basename(filePath)
    const ext = name.split('.').pop()?.toLowerCase() ?? ''
    const mime = ATTACHMENT_IMAGE_MIME[ext]
    if (!mime) return { ok: false, error: `${name}: ${tm('errNotImage')}` }
    try {
      const stat = statSync(filePath)
      if (stat.size > ATTACHMENT_IMAGE_MAX_BYTES) {
        return { ok: false, error: `${name}: ${tm('errImageTooLarge')}` }
      }
      return { ok: true, base64: readFileSync(filePath).toString('base64'), mime }
    } catch {
      return { ok: false, error: `${name}: ${tm('errUnreadable')}` }
    }
  })

  ipcMain.handle(HTML_CHANNELS.pickImage, async (e): Promise<string | null> => {
    const docPath = savePathByWc.get(e.sender.id)
    if (!docPath) return null
    const win =
      BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
    const picked = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgPickImage'),
      // only formats readImage/DOCX export can round-trip (docx-engine NewImage mimes)
      filters: [{ name: tm('filterImages'), extensions: ['png', 'jpg', 'jpeg', 'gif'] }],
      properties: ['openFile'],
    })
    const source = picked.filePaths[0]
    if (picked.canceled || !source) return null
    return copyImageIntoOwnedAssets(docPath, source)
  })

  ipcMain.handle(
    HTML_CHANNELS.saveImage,
    async (e, data: { base64?: unknown; ext?: unknown }): Promise<string | null> => {
      const docPath = savePathByWc.get(e.sender.id)
      const ext = String(data?.ext ?? '').toLowerCase()
      if (!docPath || typeof data?.base64 !== 'string' || !data.base64) return null
      // keep in sync with readImage's MIME map — every authored asset must stay DOCX-exportable
      if (!['png', 'jpg', 'jpeg', 'gif'].includes(ext)) return null
      return writeImageIntoOwnedAssets(docPath, `image.${ext}`, Buffer.from(data.base64, 'base64'))
    },
  )

  // html-owned (like docs:ai-generate-image): the shared ai:* handlers are
  // shell-registered, but image generation is gated per app
  ipcMain.handle(
    HTML_CHANNELS.aiGenerateImage,
    (_e, op: { prompt?: unknown; aspectRatio?: unknown }) =>
      generateImageTool(join(app.getPath('userData'), 'ai-settings.json'), {
        prompt: String(op?.prompt ?? ''),
        aspectRatio: op?.aspectRatio ? String(op.aspectRatio) : undefined,
      }),
  )

  const MIME_BY_EXT: Record<string, ImageData['mime']> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
  }

  ipcMain.handle(HTML_CHANNELS.readImage, async (e, src: unknown): Promise<ImageData | null> => {
    const docPath = savePathByWc.get(e.sender.id)
    if (!docPath || typeof src !== 'string' || /^[a-z][a-z0-9+.-]*:/i.test(src)) return null
    const target = await resolveSafeRelativeImagePath(docPath, src)
    if (!target || !existsSync(target)) return null
    try {
      // "save page as, complete" assets have no extension: type them from the signature
      const mime =
        MIME_BY_EXT[extname(target).toLowerCase()] ?? editableImageMime(await readHead(target))
      if (!mime) return null
      return { base64: (await readFile(target)).toString('base64'), mime }
    } catch {
      return null
    }
  })

  // remote pictures (AI-generated or hot-linked) are downloaded here: the frame's fetch is
  // CORS-bound, and fetchRemoteImage refuses private/link-local targets
  ipcMain.handle(HTML_CHANNELS.fetchImage, async (_e, url: unknown): Promise<ImageData | null> => {
    if (typeof url !== 'string' || !/^https?:/i.test(url)) return null
    try {
      const resp = await fetchRemoteImage(url)
      if (!resp?.ok) return null
      const ct = resp.headers.get('content-type') ?? ''
      const mime = ct.includes('png')
        ? 'image/png'
        : ct.includes('gif')
          ? 'image/gif'
          : 'image/jpeg'
      return { base64: Buffer.from(await resp.arrayBuffer()).toString('base64'), mime }
    } catch {
      return null
    }
  })

  ipcMain.handle(
    HTML_CHANNELS.exportDocx,
    async (e, request: ExportDocxRequest): Promise<ExportResult> => {
      if (typeof request?.html !== 'string') {
        return { ok: false, error: 'html: bad export request' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showSaveDialogWithMemory(
        dialog,
        win,
        {
          defaultPath: `${exportFileName(request.suggestedName)}.docx`,
          filters: [{ name: 'Word', extensions: ['docx'] }],
        },
        configuredDefaultSaveDir(app),
      )
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      if (docxExportPrepareHook && !(await docxExportPrepareHook(picked.filePath))) {
        return { ok: true, canceled: true }
      }
      const workDir = await mkdtemp(join(tmpdir(), 'genoffice-html-docx-'))
      let driver: ElectronBrowserDriver | null = null
      try {
        // Same document the preview shows (scripts on, relative assets via html-asset://):
        // html2docx extracts from the rendered DOM, not from a print.
        const docPath = savePathByWc.get(e.sender.id)
        const base = docPath ? assetBaseHref(dirname(docPath)) : null
        const htmlPath = join(workDir, 'export.html')
        await writeFile(htmlPath, buildPreviewDocument(request.html, base), 'utf8')
        driver = await ElectronBrowserDriver.create(HTML2DOCX_VIEWPORT)
        const { docx } = await convertHtmlToDocx({ url: pathToFileURL(htmlPath).href }, driver)
        await writeFile(picked.filePath, docx)
        openExportedDocx(picked.filePath)
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      } finally {
        await driver?.close()
        await rm(workDir, { recursive: true, force: true }).catch(() => {})
      }
    },
  )

  ipcMain.handle(
    HTML_CHANNELS.exportPdf,
    async (e, request: ExportPdfRequest): Promise<ExportResult> => {
      if (typeof request?.html !== 'string') {
        return { ok: false, error: 'html: bad export request' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showSaveDialogWithMemory(
        dialog,
        win,
        {
          defaultPath: `${exportFileName(request.suggestedName)}.pdf`,
          filters: [{ name: 'PDF', extensions: ['pdf'] }],
        },
        configuredDefaultSaveDir(app),
      )
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      const workDir = await mkdtemp(join(tmpdir(), 'genoffice-html-pdf-'))
      try {
        const docPath = savePathByWc.get(e.sender.id)
        await writeFile(picked.filePath, await renderPrintPdf(request.html, docPath, workDir))
        openExportedPdf(picked.filePath)
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      } finally {
        await rm(workDir, { recursive: true, force: true }).catch(() => {})
      }
    },
  )

  ipcMain.on(HTML_CHANNELS.provisionalTitle, (e, title: unknown) => {
    if (typeof title !== 'string' || savePathByWc.has(e.sender.id)) return
    const clean = title.replace(/\s+/g, ' ').trim().slice(0, 80)
    if (clean) provisionalTitleHook?.(e.sender, clean)
  })

  ipcMain.on(HTML_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(HTML_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // safety net for menu saves the renderer declined without invoking save()
  // (busy / still loading) — the save handler itself resolves the normal path
  ipcMain.on(HTML_CHANNELS.saveRequestAck, (e, ok: unknown) => {
    const waiter = saveWaiters.get(e.sender.id)
    saveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // Language channel shared with other modules; removeHandler tolerates duplicate registration
  ipcMain.removeHandler(HTML_CHANNELS.getLanguage)
  ipcMain.handle(HTML_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath)) {
    openPathByWc.set(wcId, openPath)
    savePathByWc.set(wcId, openPath)
    allowedByWc.set(wcId, new Set([openPath]))
  }
  installExternalLinkOpener(wc)
  wc.once('destroyed', () => {
    closePresentViewsOf(wcId)
    openPathByWc.delete(wcId)
    allowedByWc.delete(wcId)
    savePathByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    previewTextByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
    saveWaiters.get(wcId)?.(false)
    saveWaiters.delete(wcId)
  })
}

function installExternalLinkOpener(wc: WebContents): void {
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
}

/** A present view renders only the owner's preview (renderer route `?present=<owner wc id>`) */
function bindPresentView(wc: WebContents, ownerWcId: number, title: string): void {
  const wcId = wc.id
  presentOwnerByWc.set(wcId, ownerWcId)
  installExternalLinkOpener(wc)
  wc.once('destroyed', () => presentOwnerByWc.delete(wcId))
  const query = { present: String(ownerWcId), title }
  if (runtime.rendererUrl) {
    const url = new URL(runtime.rendererUrl)
    for (const [k, v] of Object.entries(query)) url.searchParams.set(k, v)
    void wc.loadURL(url.toString())
  } else if (runtime.rendererFile) void wc.loadFile(runtime.rendererFile, { query })
}

export function createHtmlPresentView(owner: WebContents, title: string): WebContentsView {
  registerHtmlIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  bindPresentView(view.webContents, owner.id, title)
  return view
}

export function createHtmlView(openPath?: string | null): WebContentsView {
  registerHtmlIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  grantAndTrack(view.webContents, openPath)
  if (runtime.rendererUrl) void view.webContents.loadURL(runtime.rendererUrl)
  else if (runtime.rendererFile) void view.webContents.loadFile(runtime.rendererFile)
  return view
}

/** Standalone window mode: `npm run dev -w @genoffice/html`, md path passed via argv */
export function startHtmlStandalone(): void {
  registerHtmlSchemes()
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configureHtmlRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerHtmlIpc()
    const win = new BrowserWindow({
      width: 1200,
      height: 850,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    const argPath = process.argv.slice(1).find((a) => /\.html?$/i.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) void win.loadURL(runtime.rendererUrl)
    else if (runtime.rendererFile) void win.loadFile(runtime.rendererFile)
  })
  app.on('window-all-closed', () => app.quit())
}
