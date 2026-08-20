#!/usr/bin/env node
/**
 * Inserts the About-dialog keys (aboutTitle, aboutSubtitle, aboutAttribution,
 * aboutLicense, aboutClose) into every locale block of the shell strings.ts,
 * anchored after each block's `closeTab:` line.
 */
import { readFileSync, writeFileSync } from 'node:fs'

const file = 'apps/shell/src/renderer/src/strings.ts'

const KEYS = {
  zh: {
    aboutTitle: '关于 KĀRYA',
    aboutSubtitle: 'KĀRYA — 智能工作台',
    aboutAttribution:
      'KĀRYA 派生自并修改自 GenOffice——一个由 Mainfunc, Inc.（Genspark）开发的开放源代码 AI 原生办公套件。本产品与 Genspark 无隶属或背书关系。',
    aboutLicense: '依据 Apache License 2.0 许可。',
    aboutClose: '关闭',
  },
  en: {
    aboutTitle: 'About KĀRYA',
    aboutSubtitle: 'KĀRYA — Intelligent Workspace',
    aboutAttribution:
      'KĀRYA is derived from and modified from GenOffice, an open-source AI-native office suite originally developed by Mainfunc, Inc. (Genspark). This product is not affiliated with or endorsed by Genspark.',
    aboutLicense: 'Licensed under the Apache License, Version 2.0.',
    aboutClose: 'Close',
  },
  ja: {
    aboutTitle: 'KĀRYA について',
    aboutSubtitle: 'KĀRYA — インテリジェントワークスペース',
    aboutAttribution:
      'KĀRYA は、Mainfunc, Inc.（Genspark）が開発したオープンソースの AI ネイティブオフィススイート GenOffice を元にしたフォークです。本製品は Genspark とは提携・承認関係にありません。',
    aboutLicense: 'Apache License 2.0 に基づいてライセンスされています。',
    aboutClose: '閉じる',
  },
  ko: {
    aboutTitle: 'KĀRYA 정보',
    aboutSubtitle: 'KĀRYA — 지능형 워크스페이스',
    aboutAttribution:
      'KĀRYA는 Mainfunc, Inc.(Genspark)가 개발한 오픈소스 AI 네이티브 오피스 제품군 GenOffice에서 파생·수정되었습니다. 본 제품은 Genspark와 제휴 또는 보증 관계가 없습니다.',
    aboutLicense: 'Apache License 2.0에 따라 사용이 허가됩니다.',
    aboutClose: '닫기',
  },
  fr: {
    aboutTitle: 'À propos de KĀRYA',
    aboutSubtitle: 'KĀRYA — Espace de travail intelligent',
    aboutAttribution:
      "KĀRYA est dérivé et modifié de GenOffice, une suite bureautique open source native IA développée à l'origine par Mainfunc, Inc. (Genspark). Ce produit n'est ni affilié à Genspark ni approuvé par Genspark.",
    aboutLicense: 'Sous licence Apache License 2.0.',
    aboutClose: 'Fermer',
  },
  de: {
    aboutTitle: 'Über KĀRYA',
    aboutSubtitle: 'KĀRYA — Intelligenter Arbeitsbereich',
    aboutAttribution:
      'KĀRYA ist eine Abwandlung von GenOffice, einer Open-Source-Office-Suite mit KI-Fokus, die ursprünglich von Mainfunc, Inc. (Genspark) entwickelt wurde. Dieses Produkt ist weder mit Genspark verbunden noch von Genspark unterstützt.',
    aboutLicense: 'Lizenziert unter der Apache License 2.0.',
    aboutClose: 'Schließen',
  },
  es: {
    aboutTitle: 'Acerca de KĀRYA',
    aboutSubtitle: 'KĀRYA — Espacio de trabajo inteligente',
    aboutAttribution:
      'KĀRYA se deriva de GenOffice y se basa en él, una suite ofimática de código abierto con IA desarrollada originalmente por Mainfunc, Inc. (Genspark). Este producto no está afiliado ni respaldado por Genspark.',
    aboutLicense: 'Con licencia Apache License 2.0.',
    aboutClose: 'Cerrar',
  },
  th: {
    aboutTitle: 'เกี่ยวกับ KĀRYA',
    aboutSubtitle: 'KĀRYA — พื้นที่ทำงานอัจฉริยะ',
    aboutAttribution:
      'KĀRYA พัฒนาต่อยอดและปรับปรุงจาก GenOffice ชุดโปรแกรมสำนักงานโอเพนซอร์สที่เน้น AI ซึ่งพัฒนาโดย Mainfunc, Inc. (Genspark) ผลิตภัณฑ์นี้ไม่มีความเกี่ยวข้องหรือได้รับการรับรองจาก Genspark',
    aboutLicense: 'อยู่ภายใต้สัญญาอนุญาต Apache License 2.0',
    aboutClose: 'ปิด',
  },
  id: {
    aboutTitle: 'Tentang KĀRYA',
    aboutSubtitle: 'KĀRYA — Ruang Kerja Cerdas',
    aboutAttribution:
      'KĀRYA diturunkan dan dimodifikasi dari GenOffice, paket aplikasi perkantoran sumber terbuka berbasis AI yang awalnya dikembangkan oleh Mainfunc, Inc. (Genspark). Produk ini tidak berafiliasi dengan atau didukung oleh Genspark.',
    aboutLicense: 'Dilisensikan di bawah Apache License 2.0.',
    aboutClose: 'Tutup',
  },
  ru: {
    aboutTitle: 'О KĀRYA',
    aboutSubtitle: 'KĀRYA — интеллектуальное рабочее пространство',
    aboutAttribution:
      'KĀRYA является производным от GenOffice — офисного пакета с открытым исходным кодом и встроенным ИИ, разработанного Mainfunc, Inc. (Genspark). Этот продукт не связан с Genspark и не одобрен им.',
    aboutLicense: 'Распространяется по лицензии Apache License 2.0.',
    aboutClose: 'Закрыть',
  },
  ar: {
    aboutTitle: 'حول KĀRYA',
    aboutSubtitle: 'KĀRYA — مساحة عمل ذكية',
    aboutAttribution:
      'KĀRYA مشتق من GenOffice، حزمة المكاتب مفتوحة المصدر القائمة على الذكاء الاصطناعي التي طورتها Mainfunc, Inc. (Genspark). هذا المنتج غير تابع لـ Genspark ولا يحظى بتأييدها.',
    aboutLicense: 'مرخّص بموجب رخصة Apache License 2.0.',
    aboutClose: 'إغلاق',
  },
  pt: {
    aboutTitle: 'Sobre o KĀRYA',
    aboutSubtitle: 'KĀRYA — Espaço de trabalho inteligente',
    aboutAttribution:
      'O KĀRYA é derivado e modificado do GenOffice, um pacote de escritório de código aberto com IA desenvolvido originalmente pela Mainfunc, Inc. (Genspark). Este produto não é afiliado nem endossado pela Genspark.',
    aboutLicense: 'Licenciado sob a Apache License 2.0.',
    aboutClose: 'Fechar',
  },
  it: {
    aboutTitle: 'Informazioni su KĀRYA',
    aboutSubtitle: 'KĀRYA — Area di lavoro intelligente',
    aboutAttribution:
      'KĀRYA è derivato e modificato da GenOffice, una suite per ufficio open source con IA originariamente sviluppata da Mainfunc, Inc. (Genspark). Questo prodotto non è affiliato né approvato da Genspark.',
    aboutLicense: 'Concesso in licenza con Apache License 2.0.',
    aboutClose: 'Chiudi',
  },
  pl: {
    aboutTitle: 'O KĀRYA',
    aboutSubtitle: 'KĀRYA — Inteligentne miejsce pracy',
    aboutAttribution:
      'KĀRYA pochodzi z GenOffice i jest jego modyfikacją — pakietu biurowego o otwartym kodzie źródłowym z wbudowaną AI, opracowanego pierwotnie przez Mainfunc, Inc. (Genspark). Ten produkt nie jest powiązany z Genspark ani przez niego zatwierdzony.',
    aboutLicense: 'Licencjonowane na Apache License 2.0.',
    aboutClose: 'Zamknij',
  },
  nl: {
    aboutTitle: 'Over KĀRYA',
    aboutSubtitle: 'KĀRYA — Intelligente werkruimte',
    aboutAttribution:
      'KĀRYA is afgeleid van en aangepast op basis van GenOffice, een opensource-kantoorpakket met AI dat oorspronkelijk is ontwikkeld door Mainfunc, Inc. (Genspark). Dit product is niet verbonden aan of goedgekeurd door Genspark.',
    aboutLicense: 'Gelicenseerd onder de Apache License 2.0.',
    aboutClose: 'Sluiten',
  },
  ms: {
    aboutTitle: 'Perihal KĀRYA',
    aboutSubtitle: 'KĀRYA — Ruang Kerja Pintar',
    aboutAttribution:
      'KĀRYA diterbitkan dan diubah suai daripada GenOffice, suite pejabat sumber terbuka berasaskan AI yang dibangunkan oleh Mainfunc, Inc. (Genspark). Produk ini tidak bergabung dengan atau disokong oleh Genspark.',
    aboutLicense: 'Dilesenkan di bawah Apache License 2.0.',
    aboutClose: 'Tutup',
  },
  he: {
    aboutTitle: 'אודות KĀRYA',
    aboutSubtitle: 'KĀRYA — סביבת עבודה חכמה',
    aboutAttribution:
      'KĀRYA נגזר ומשונה מ-GenOffice, חבילת משרדים בקוד פתוח עם בינה מלאכותית שפותחה במקור על ידי Mainfunc, Inc. (Genspark). מוצר זה אינו מסונף ל-Genspark ואינו מאושר על ידה.',
    aboutLicense: 'ברישיון Apache License 2.0.',
    aboutClose: 'סגירה',
  },
  hi: {
    aboutTitle: 'KĀRYA के बारे में',
    aboutSubtitle: 'KĀRYA — बुद्धिमान कार्यक्षेत्र',
    aboutAttribution:
      'KĀRYA GenOffice से व्युत्पन्न और संशोधित है, जो मूल रूप से Mainfunc, Inc. (Genspark) द्वारा विकसित एक ओपन-सोर्स AI-नेटिव ऑफिस सुइट है। यह उत्पाद Genspark से संबद्ध या समर्थित नहीं है।',
    aboutLicense: 'Apache License 2.0 के अंतर्गत लाइसेंस प्राप्त।',
    aboutClose: 'बंद करें',
  },
  'zh-TW': {
    aboutTitle: '關於 KĀRYA',
    aboutSubtitle: 'KĀRYA — 智慧工作台',
    aboutAttribution:
      'KĀRYA 衍生自並修改自 GenOffice——一個由 Mainfunc, Inc.（Genspark）開發的開放原始碼 AI 原生辦公套件。本產品與 Genspark 無隸屬或背書關係。',
    aboutLicense: '依據 Apache License 2.0 許可。',
    aboutClose: '關閉',
  },
}

const src = readFileSync(file, 'utf8')
const lines = src.split('\n')
const out = []
let locale = null
let inserted = 0

for (const line of lines) {
  const block = /^ {2}'?([A-Za-z-]+)'?: \{/.exec(line)
  if (block) {
    locale = block[1]
  }
  out.push(line)
  if (locale && /^ {4}closeTab: /.test(line)) {
    const kv = KEYS[locale]
    if (!kv) throw new Error(`no translations for locale ${locale}`)
    for (const [k, v] of Object.entries(kv)) {
      out.push(`    ${k}: ${JSON.stringify(v)},`)
    }
    inserted += 5
    locale = null // one insertion per block
  }
}

if (inserted !== 19 * 5) throw new Error(`expected 95 insertions, got ${inserted}`)
writeFileSync(file, out.join('\n'))
console.log(`inserted ${inserted} keys across 19 locale blocks`)
