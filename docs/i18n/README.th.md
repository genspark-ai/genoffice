<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — สวีท Office ที่ขับเคลื่อนด้วย AI แบบโอเพนซอร์ส: Docs, Sheets, Slides, PDF, Markdown และ HTML พร้อมแผง AI ในตัว" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>ชุดโปรแกรม Office AI แบบโอเพนซอร์สที่ครบฟีเจอร์รายแรกของโลก</b><br>
ไฟล์ Word, Excel, PowerPoint และ PDF แก้ไขโดยคุณและ AI ของคุณ บันทึกกลับเป็นฟอร์แมตจริง</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <b>ไทย</b> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>ดาวน์โหลด</b></a> ·
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
  <a href="https://genoffice.ai/"><b>เว็บไซต์</b></a> ·
  <a href="https://genoffice.ai/join"><b>คอมมูนิตี้</b></a> ·
  <a href="../../PRIVACY.md"><b>ความเป็นส่วนตัว</b></a>
</p>

GenOffice คือทางเลือกฟรีและโอเพนซอร์สของ Microsoft Office สำหรับ macOS,
Windows และ Linux เปิดและบันทึกไฟล์ `.docx`, `.xlsx` และ `.pptx` แบบดั้งเดิม
แก้ไข PDF, Markdown และ HTML ได้ และวางเอเจนต์ AI ไว้ข้างเอกสารทุกฉบับ —
ไม่ใช่กล่องแชทที่แปะไว้ข้างๆ แต่เป็นเอดิเตอร์ที่อ่านไฟล์ ทำการแก้ไข
และแสดงให้เห็นชัดเจนว่าแก้ไขส่วนใดไปบ้าง

- **ฟอร์แมตจริง คงบิตต่อบิต** เฉพาะส่วนที่คุณแก้ไขเท่านั้นที่จะถูกเขียนใหม่
  ส่วนอื่นๆ ของไฟล์จะยังคงเดิมทุกไบต์ เอกสารจึงใช้งานได้ต่อใน Word,
  Excel และ PowerPoint
- **AI ที่ตรวจสอบได้** การแก้ไขจะมาในรูปแบบ tracked changes และ diff
  พร้อมปุ่มย้อนกลับคลิกเดียว สเปรดชีตได้สูตรที่คำนวณสด ไม่ใช่ตัวเลขที่วางทับ
  สไลด์และหน้าเอกสารถูกสร้างขึ้นบนแคนวาสและยังแก้ไขได้เต็มรูปแบบ
- **ทำงานในเครื่องโดยออกแบบมาแบบนั้น** ไฟล์เปิด แก้ไข บันทึก
  และแปลงบนเครื่องของคุณเอง PDF → Word / Excel / PowerPoint, Markdown → Word
  และ HTML → Word ทั้งหมดทำงานในเครื่อง มีเพียงการเรียก AI เท่านั้นที่ออกจาก
  เครื่องไปยังผู้ให้บริการที่คุณเลือก
- **ใช้คีย์ของคุณเองหรือไม่ต้องใช้เลยก็ได้** ลงชื่อเข้าใช้ด้วย Genspark
  แล้วข้ามการตั้งค่าคีย์ไปได้เลย หรือใช้คีย์ของคุณเองกับ Claude, OpenAI,
  Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral,
  OpenRouter, Requesty หรือ endpoint ที่รองรับ OpenAI-compatible ใดๆ
  รวมถึงเซิร์ฟเวอร์โมเดลภายในเครื่องด้วย
- **เขียนสคริปต์ได้และพร้อมสำหรับเอเจนต์** แอปมาพร้อมบรรทัดคำสั่ง `genoffice`
  และสกิลสำหรับ Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot,
  OpenCode และ Windsurf ทำให้เอเจนต์เขียนโค้ดสามารถสร้าง แปลง อ่าน
  และแก้ไขไฟล์ Office จริงบนเครื่องของคุณได้โดยไม่ต้องเปิดหน้าต่างใดๆ

**ดาวน์โหลด:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon และ Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 และ Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
รายละเอียดและความต้องการของระบบดูได้ในหัวข้อ [ดาวน์โหลด](#download)

## เดโม

หกแอป แผง AI เดียว และบรรทัดคำสั่งสำหรับเอเจนต์เขียนโค้ดของคุณ
ทุกภาพหน้าจอคือแอปจริงที่รันบน macOS โดย AI ถูกสั่งงานจากพรอมป์ที่คุณอ่านได้ในแผง

### 1 · Docs — เปิดและแก้ไข `.docx` ด้วย AI ที่ตรวจสอบได้

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs แสดงหน้ารายงานประจำปีแบบสองคอลัมน์ พร้อมภาพหน้าปกเต็มความกว้าง ตาราง KPI แบบมีเงา หัวกระดาษและท้ายกระดาษ ที่ซูม 80% โดยแผง AI ถูกย่อไว้"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: ภาพรวมบริษัทพร้อมภาพแบนเนอร์ AI ได้ปรับส่วน Overview ให้กระชับขึ้นและแทรกส่วนบุลเลตใหม่ โดยแผงมีปุ่มย้อนกลับคลิกเดียว"></td>
</tr>
<tr>
<td><b>เปิดไฟล์ตามที่ Word จัดวาง</b> — ส่วนสองคอลัมน์ ภาพเต็มหน้า ตารางแบบมีเงา หัวกระดาษและท้ายกระดาษ การแบ่งหน้าตามมาตรวัดบรรทัดของ Word สไตล์ ความคิดเห็น tracked changes สมการ และงานเขียนด้วยลายมือคงเดิมทุกรอบการเปิด-บันทึก</td>
<td><b>ขอให้แก้ไข</b> — AI อ่านบล็อกที่จำเป็น เขียนส่วน Overview ใหม่ และแทรกส่วนบุลเลตใหม่ ทุกรอบของ AI คือสแนปช็อตที่ย้อนกลับได้ และเมื่อเปิด <b>Track changes</b> การแก้ไขจะมาในรูปแบบการแก้ไขสไตล์ Word</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` ที่มีสูตรและกราฟที่คำนวณสด ไม่ใช่ตัวเลขที่วางทับ

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: AI เพิ่มชีต Summary พร้อมรายรับตามภูมิภาคและหมวดหมู่โดยใช้สูตร SUMIF บวกกราฟแท่ง และรายงานว่ามีการเปลี่ยนแปลง 43 รายการพร้อมปุ่ม Undo"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: เมื่อถูกถามว่าภูมิภาคใดมีรายรับ Q2 นำหน้า AI ตอบว่ายุโรป พร้อมรายละเอียดแยกตามหมวดหมู่และอ้างอิงเซลล์ที่ใช้เป็นลิงก์ ข้างชีต Orders"></td>
</tr>
<tr>
<td><b>สร้างมันขึ้นมา</b> — จากประโยคเดียว เอเจนต์เพิ่มชีต Summary พร้อมสูตร <code>SUMIF</code> จริงแยกตามภูมิภาคและหมวดหมู่ แทรกกราฟแท่ง และรวมการเปลี่ยนแปลงทั้ง 43 รายการเป็นชุดเดียวที่ย้อนกลับได้</td>
<td><b>ถามมัน</b> — คำถามเกี่ยวกับเวิร์กบุ๊กจะได้คำตอบพร้อมเหตุผลและเซลล์ที่ใช้จริงในรูปแบบลิงก์อ้างอิงที่คลิกได้ เบื้องหลังคือเอนจิน `.xlsx` ของ Rust ที่พัฒนาขึ้นเอง ตาราง pivot, slicer, การจัดรูปแบบตามเงื่อนไข และการตรวจสอบที่มาของสูตร</td>
</tr>
</table>

### 3 · Slides — จากพรอมป์สู่สไลด์ `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="ภาพไทม์แลปส์ของ GenOffice Slides ขณะสร้างสไลด์นำเสนอนักลงทุน Aurora Home: AI วางโครงเรื่องในแผงควบคุม สไลด์ปรากฏบนแคนวาสทีละแผ่น และเดคที่เสร็จสมบูรณ์ปิดท้ายด้วย closing ask" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: สไลด์หน้าปกของดาต้าเดคนักลงทุน Aurora Home ที่สร้างโดย AI บนแคนวาส พร้อมพรอมป์บรรทัดเดียวต้นฉบับและสรุปของ AI ว่าสร้างอะไรไปในแผง"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: สไลด์ปิดที่ออกแบบแล้วของเดคเดียวกันที่มี 11 สไลด์ พร้อมแถบภาพย่อทางซ้ายและแผง AI สรุปเนื้อเรื่อง"></td>
</tr>
<tr>
<td><b>ใส่แค่บรรทัดเดียว</b> — "สร้างสไลด์นำเสนอนักลงทุน 10 สไลด์สำหรับ Aurora Home…" GenOffice วางเนื้อเรื่อง ค้นคว้าตัวเลข และร่างทุกสไลด์ลงบนแคนวาสเป็น `.pptx` จริง</td>
<td><b>ได้เดคที่เสร็จสมบูรณ์</b> — สไลด์ที่ออกแบบแล้ว 11 สไลด์ด้วยไทโปกราฟีและภาพที่สอดคล้องกัน และคำกระตุ้นการตัดสินใจปิดท้าย แก้ไขต่อได้ด้วย master, layout, smart guide และการครอปแบบไม่ทำลายต้นฉบับ หรือให้แผงช่วยปรับสไตล์ เขียนใหม่ และจัดลำดับใหม่</td>
</tr>
</table>

### 4 · PDF — แก้ไขข้อความ PDF ในตำแหน่งเดิม แปลง PDF เป็น Word ในเครื่อง

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: โหมด Edit text ขีดกรอบทุกบล็อกข้อความบนหน้าเพื่อแก้ไขในตำแหน่งเดิม ขณะที่แผง AI ตอบคำถามเกี่ยวกับรายงานพร้อมอ้างอิงหมายเลขหน้า"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs แสดงเอกสาร Word ที่แปลงในเครื่องจาก PDF รายงานรายไตรมาส Helios เปิดในแท็บที่สองข้างไฟล์ PDF ต้นฉบับ"></td>
</tr>
<tr>
<td><b>แก้ไขในหน้าเอกสาร</b> — โหมด Edit text ขีดกรอบทุกบล็อกข้อความเพื่อพิมพ์ทับในตำแหน่งเดิม โดย content stream ถูกเขียนใหม่ผ่าน PDFium ด้วยฟอนต์เดิม ไม่ใช่การแปะข้อความทับ ถาม AI เกี่ยวกับรายงานยาวๆ แล้วได้คำตอบพร้อมอ้างอิงหมายเลขหน้า</td>
<td><b>แปลงในเครื่อง</b> — <b>PDF Converter → PDF to Word</b> สร้าง `.docx` ที่แก้ไขได้ ซึ่งเปิดใน Docs ข้างไฟล์ต้นฉบับ โดยหัวเรื่อง แถวข้อมูลสถิติ และพารากราฟยังคงอยู่ครบ ปลายทาง Excel และ PowerPoint ก็ทำงานแบบเดียวกัน ส่วนหน้าที่สแกนมาจะผ่านระบบ OCR ของเครื่อง</td>
</tr>
</table>

### 5 · HTML — ตัวช่วยสร้างหน้าเว็บและ UI ด้วย AI โดยเริ่มจากบรีฟการออกแบบ

บอกว่าหน้านี้มีไว้เพื่ออะไรและสำหรับใคร AI จะเสนอ **บรีฟการออกแบบ**
ก่อน — hook, จานสี, ไทโปกราฟี และแนวทางสไตล์ — จากนั้นจึงสร้างไฟล์ `.html`
แบบครบในตัวไฟล์เดียวตามโทเคนเหล่านั้น

<img src="../assets/readme/html-restyle-motion.webp" alt="ภาพไทม์แลปส์ของ GenOffice HTML ขณะปรับสไตล์หน้า Landing Page ของ Lumen: คำขอ Restyle เพียงครั้งเดียวในแผงควบคุมเปลี่ยนหน้า Midnight Studio ที่มืดให้เป็นเวอร์ชัน Solar Daybreak ที่อบอุ่น โดยทุกส่วนและเนื้อหาทั้งหมดยังคงอยู่ในตำแหน่งเดิม" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: หน้า landing page ที่สร้างขึ้นสำหรับโคมไฟตั้งโต๊ะพลังงานแสงอาทิตย์ในแนว Midnight Studio โทนมืด แสดงในพรีวิวสดพร้อมแผง AI สรุปหน้าที่สร้างไปแล้ว"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="หน้า landing page Lumen เดิมที่ถูก AI ปรับสไตล์ใหม่เป็นแนว Solar Daybreak โทนอบอุ่น: พื้นหลังสีกระดาษ หัวข้อฟอนต์ serif และสีเน้นส้ม โดยทุกส่วนและเนื้อหาทั้งหมดยังคงอยู่"></td>
</tr>
<tr>
<td><b>สร้างจากพรอมป์เดียว</b> — hero ที่โดดเด่น การ์ดฟีเจอร์ ราคา และฟอร์ม waitlist สำหรับ Lumen สร้างในแนว Midnight Studio คลิกที่อีลิเมนต์ใดก็ได้เพื่อปรับสไตล์ ดับเบิลคลิกเพื่อแก้ไขข้อความ หรือสลับไปดูซอร์สโค้ดผ่าน CodeMirror</td>
<td><b>ดีไซน์เดิม แนวทางใหม่</b> — คำขอ <b>Restyle</b> เดียวสลับโทเคนของบรีฟ และหน้าเว็บก็ปรับตาม: พื้นกระดาษโทนอุ่น ฟอนต์ serif แบบบทบรรณาธิการ สีเน้นส้มแดด โดยไม่มีอะไรถูกเขียนใหม่ทั้งหมด นำเสนอแบบเต็มจอ หรือส่งออกเป็น PDF หรือเอกสาร Word ที่แก้ไขได้แบบเนทิฟ</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: หน้า dashboard ส่วนตัวที่สร้างขึ้นสำหรับนักออกแบบฟรีแลนซ์ในสไตล์ผ้าลินินโทนอุ่น มีแถบด้านซ้าย คำทักทายฟอนต์ serif และการ์ดเมตริกสี่ใบ"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: รายงานข้อมูลตลาดรถยนต์ไฟฟ้าที่สร้างขึ้นในสไตล์หนังสือพิมพ์ พร้อมหัวเรื่องฟอนต์ serif ตัวเลขเด่น 17.3 ล้าน และแถวข้อมูลสถิติ"></td>
</tr>
<tr>
<td><b>โมเดล UI จำลอง</b> — เทมเพลตเริ่มต้น "personal dashboard" เปลี่ยนเพอร์โซนาให้เป็นเลย์เอาต์ที่ใช้งานได้จริง: แถบด้านซ้าย คำทักทาย สปาร์กไลน์ชั่วโมงที่เรียกเก็บเงินได้ การ์ดใบแจ้งหนี้และการใช้งาน ทั้งหมดเป็น HTML จริงที่ส่งต่อให้นักพัฒนาได้</td>
<td><b>เรื่องราวจากข้อมูล</b> — เทมเพลตเริ่มต้น "data report" สร้างหนังสือพิมพ์บทบรรณาธิการ: หัวเรื่องฟอนต์ serif ตัวเลขเด่นตัวเดียว แถวข้อมูลสถิติที่มีเส้นแบ่ง กราฟ SVG แบบฝังในตัว และหมายเหตุวิธีการ</td>
</tr>
</table>

### 6 · Markdown — บล็อกเอดิเตอร์เหนือไฟล์ `.md` แบบข้อความล้วน พร้อม Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: พารากราฟที่เลือกไว้แสดงป็อปโอเวอร์ Ask AI พร้อมคำสั่งที่พิมพ์และชิปคำแนะนำ เช่น Polish, Make more concise, Expand และ Fix grammar พร้อมปุ่ม Send now และ Add to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown แสดงเอกสารบันทึกการเปิดตัวที่มีตาราง แผนภาพโฟลว์ Mermaid และรายการงาน พร้อมพรอมป์เริ่มต้นของแผง AI ทางซ้าย"></td>
</tr>
<tr>
<td><b>ถาม AI เกี่ยวกับข้อความที่เลือก</b> — เลือกข้อความส่วนใดก็ได้ แล้วชิป <b>Ask AI</b> จะปรากฏขึ้น: พิมพ์คำสั่งหรือเลือกคำแนะนำ ส่งทันที หรือจัดคิวการแก้ไขแบบยึดตำแหน่งไว้หลายรายการแล้วรันทีเดียว จุดเข้าใช้งานเดียวกันนี้มีอยู่ในทุกแอป</td>
<td><b>เรนเดอร์แล้วบันทึกกลับเป็น Markdown ล้วน</b> — หัวเรื่อง รายการ ตาราง ภาพ บล็อกโค้ด และไดอะแกรม Mermaid ในบล็อกเอดิเตอร์ Tiptap เขียนกลับเป็น <code>.md</code> ล้วน พร้อมการส่งออก <b>Markdown → Word</b> ที่ทำงานในเครื่องทั้งหมด</td>
</tr>
</table>

### 7 · CLI — เอเจนต์เขียนโค้ดของคุณสั่งงาน GenOffice บนเครื่องของคุณ

GenOffice มาพร้อมบรรทัดคำสั่ง `genoffice` และสกิลสำหรับเอเจนต์ ติดตั้งสกิลแล้ว
Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode หรือ Windsurf
ก็สร้าง แปลง อ่าน และแก้ไขไฟล์ Office จริงผ่านเอนจินเดียวกับตัวแอปได้
โดยไม่ต้องเปิดหน้าต่างใดๆ

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides แสดงเดคเรื่องระบบสุริยะ 8 สไลด์ที่เอเจนต์เขียนโค้ดสร้างผ่านบรรทัดคำสั่ง genoffice: สไลด์หน้าปกบนแคนวาส ภาพย่อ 8 ภาพทางซ้าย และแผง AI ที่เปิดอยู่" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="สไลด์ที่เรนเดอร์แล้วทั้ง 8 แผ่นของเดคระบบสุริยะเรียงเคียงกัน: หน้าปก ไทม์ไลน์การสำรวจ ตัวเลขสำคัญ 4 ตัว กราฟแท่งเส้นผ่านศูนย์กลางดาวเคราะห์ ดาวเคราะห์หินเทียบกับดาวยักษ์ ตัวเลขเด่น 99.8% ของดวงอาทิตย์ กริดดาวยักษ์ทั้ง 4 และบทสรุป"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="GenOffice Settings หน้า Integrations: สกิล genoffice ติดตั้งใน Claude Code แล้ว พร้อมปุ่ม Install ข้าง Codex และ Cursor"></td>
</tr>
<tr>
<td><b>พรอมป์เดียวถึงเอเจนต์ของคุณ</b> — "สร้างเดค 8 สไลด์เกี่ยวกับระบบสุริยะ" เอเจนต์อ่านสกิล เขียนสไตล์ชีต โครงร่าง และสเปกหน้าหนึ่งชุดต่อสไลด์ สร้างภาพถ่ายสองภาพด้วย <code>genoffice image</code> และให้ <code>genoffice slides check</code> ปฏิเสธทุกอย่างที่ล้นหรือซ้อนทับกันก่อนที่ <code>genoffice create</code> จะประกอบ <code>.pptx</code> และ <code>slides render</code> ส่ง PNG กลับมาให้ดูหนึ่งภาพต่อสไลด์</td>
<td><b>ติดตั้งครั้งเดียวจาก การตั้งค่า → การเชื่อมต่อ</b> — GenOffice แสดงรายการเอเจนต์เขียนโค้ดที่พบในคอมพิวเตอร์เครื่องนี้และเขียนสกิลลงในแต่ละตัวที่คุณเลือก หรือดาวน์โหลดสกิลเป็น zip หรือรัน <code>npx skills add genspark-ai/genoffice</code> คำสั่งและเวิร์กโฟลว์ฉบับเต็มอยู่ใน <a href="#command-line-and-agent-skill">บรรทัดคำสั่งและสกิลสำหรับเอเจนต์</a></td>
</tr>
</table>

### 8 · MCP — เครื่องมือชุดเดียวกันผ่าน Model Context Protocol

ทุกคำสั่ง `genoffice` ก็เป็นเครื่องมือ MCP ด้วยเช่นกัน Claude Code, Claude
Desktop, Cursor และไคลเอนต์ MCP อื่นใดสามารถสั่ง `genoffice mcp` ได้เอง
โดยไม่ต้องติดตั้งสกิลและไม่ต้องเปิดหน้าต่างใดๆ แล้วได้เครื่องมือ 29 ตัว
พร้อมเอกสารอ้างอิงการทำงานในรูปแบบ resource เซิร์ฟเวอร์ HTTP ตัวที่สอง
ภายในแอปยังช่วยให้เอเจนต์สร้างเอกสาร Word ในแท็บเอดิเตอร์ที่มองเห็นได้
ขณะที่คุณเฝ้าดูอยู่ได้ด้วย

<img src="../assets/readme/mcp-deck-motion.webp" alt="ภาพไทม์แลปส์ของ Claude Code กำลังสร้างเดคบรีฟให้นักลงทุนเรื่องพลังงานหมุนเวียน 8 สไลด์ผ่านเซิร์ฟเวอร์ genoffice MCP: ค้นหาข้อมูลและภาพถ่ายด้วย search ตรวจสอบภาพถ่ายที่เป็นตัวเลือกแต่ละภาพด้วย media, deck_start เขียนสไตล์ชีตและโครงร่าง, deck_page เพิ่มหน้าที่ผ่านการตรวจสอบทีละหน้า, deck_build ประกอบไฟล์ .pptx และ slides_render ส่งภาพของทุกสไลด์กลับมา; เดคที่เสร็จแล้วจึงเปิดใน GenOffice Slides" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides แสดงเดคเรื่องพลังงานหมุนเวียน 2026 จำนวน 8 สไลด์ที่ Claude Code สร้างผ่านเซิร์ฟเวอร์ genoffice MCP: สไลด์หน้าปกที่มีภาพถ่ายฟาร์มกังหันลมบนแคนวาสและภาพย่อ 8 ภาพทางซ้าย"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="GenOffice Settings หน้า Integrations ส่วน MCP: คำสั่ง claude mcp add บรรทัดเดียวสำหรับ Claude Code, บล็อก JSON สำหรับ Cursor, Claude Desktop และไคลเอนต์ MCP อื่นๆ และตัวเลือกเซิร์ฟเวอร์ HTTP ในเครื่องด้านล่าง"></td>
</tr>
<tr>
<td><b>พรอมป์เดียว เรียกเครื่องมือสามสิบแปดครั้ง ไม่ต้องใช้เทอร์มินัล</b> — "สร้างบรีฟให้นักลงทุน 8 สไลด์เกี่ยวกับพลังงานหมุนเวียนในปี 2026 ใส่ภาพถ่ายจริงที่หน้าปกและทุกจุดที่ภาพช่วยได้" เอเจนต์ดึงข้อมูลและภาพถ่ายด้วย <code>search</code> ถาม <code>media</code> ว่าภาพที่เป็นตัวเลือกแต่ละภาพเป็นภาพถ่ายจริงหรือไม่ จากนั้นเรียก <code>deck_start</code> พร้อมสไตล์ชีตและโครงร่าง แล้วเรียก <code>deck_page</code> ทีละสไลด์ ทุกหน้าจะถูกตรวจสอบกับโครงร่างและจานสีก่อนจะถูกเก็บไว้ <code>deck_build</code> ประกอบ <code>.pptx</code>, <code>slides_audit</code> ตรวจหาความล้น, <code>slides_render</code> ส่ง PNG กลับมาทีละสไลด์เป็นเนื้อหาภาพให้โมเดลดู และสุดท้าย <code>deck_replace</code> แก้ไขสามหน้าที่ไม่ถูกใจ</td>
<td><b>เชื่อมต่อครั้งเดียวจาก การตั้งค่า → การเชื่อมต่อ</b> — คัดลอกคำสั่ง <code>claude mcp add</code> สำหรับ Claude Code หรือบล็อก JSON ไปวางใน Cursor, Claude Desktop หรือไคลเอนต์ MCP อื่นใด ตัวเลือก B เปิดใช้เซิร์ฟเวอร์ HTTP ในเครื่องสำหรับเอดิเตอร์ Word ที่มองเห็นได้ ทั้งสองแบบอธิบายไว้ใน <a href="#mcp-server">เซิร์ฟเวอร์ MCP</a></td>
</tr>
</table>

## ทำไมต้อง GenOffice

- **โอเพนซอร์ส** Apache-2.0 พัฒนาแบบเปิดบน GitHub
- **เป็นของคุณให้รันเอง** แอปเนทิฟสำหรับ macOS, Windows และ Linux ไฟล์อยู่บนดิสก์
  ของคุณ และการแก้ไข บันทึก และแปลงไฟล์ทุกครั้งเกิดขึ้นบนเครื่องของคุณเอง
- **ไฟล์ Office ตัวจริง** `.docx`, `.xlsx` และ `.pptx` แบบเนทิฟ คงบิตต่อบิต:
  ส่วนของไฟล์ที่คุณไม่ได้แก้ไขจะถูกคัดลอกไว้เหมือนเดิม
- **AI ที่แก้ไขเอกสารได้เอง** tracked changes ใน Docs สูตรและกราฟที่คำนวณสดใน
  Sheets สไลด์ที่วาดตรงบนแคนวาส ทุกรอบของ AI คือสแนปช็อตที่คุณย้อนกลับได้
- **โมเดลของคุณ คีย์ของคุณ** ลงชื่อเข้าใช้ด้วย Genspark หรือใส่คีย์ของคุณเองสำหรับ
  Claude, OpenAI, Gemini, DeepSeek และอื่นๆ รวมถึงเซิร์ฟเวอร์ในเครื่องและ
  endpoint ที่รองรับ OpenAI-compatible ใดๆ
- **PDF ที่ทำได้อย่างถูกต้อง** แก้ไขข้อความในตำแหน่งเดิมบนหน้า และแปลง PDF เป็น
  Word, Excel หรือ PowerPoint ในเครื่อง พร้อม OCR ของระบบสำหรับไฟล์สแกน
- **Markdown และ HTML ด้วย** ใช้แผง AI ตัวเดียวกันและส่งออกไปยัง Word ในเครื่อง
- **เขียนสคริปต์ได้** บรรทัดคำสั่ง `genoffice`, สกิลสำหรับเอเจนต์ และเซิร์ฟเวอร์
  MCP นำทุกเอนจินมารับใช้ Claude Code, Claude Desktop, Codex, Cursor และ
  เอเจนต์อื่นๆ โดยยังคงทำงานในเครื่อง
- **ฟรี** สำหรับบุคคลทั่วไปและทีมเช่นเดียวกัน

## แบ็กเอนด์ AI

**ลงชื่อเข้าใช้ด้วย Genspark** แล้วไม่ต้องตั้งค่าอะไรเลย: การเรียกโมเดล
จะส่งผ่านพร็อกซีของ Genspark (ตระกูล Claude, GPT และ Gemini) และเอเจนต์
จะได้รับความสามารถค้นหาเว็บและภาพ สร้างภาพ และวิเคราะห์ภาพ/เสียง/วิดีโอ

**หรือใช้คีย์ของคุณเอง** เมนู Settings → AI มีให้เลือก Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral,
OpenRouter, Requesty และ OpenCode Zen/Go พร้อมช่องสำหรับ endpoint ที่รองรับ
OpenAI-compatible แบบกำหนดเอง (base URL + คีย์) รวมถึงเซิร์ฟเวอร์โมเดล
ในเครื่องด้วย การค้นหาและมีเดียมีผู้ให้บริการแยกตามความสามารถของตัวเอง
ภายใต้ **AI Media & Search**: Serper หรือ Tavily สำหรับการค้นหาเว็บ
และ OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax หรือ
endpoint ที่รองรับ OpenAI-compatible ใดๆ สำหรับการสร้างภาพและการวิเคราะห์
ภาพ/วิดีโอ

สวีททั้งชุดมาพร้อมธีมสว่าง มืด และตามระบบ ธีมจะเปลี่ยนแค่สิ่งที่แสดง
บนหน้าจอเท่านั้น: การส่งออก การพิมพ์ และไฟล์ที่บันทึกจะคงสีของเอกสาร
เดิมไว้เสมอ

<a id="command-line-and-agent-skill"></a>

## บรรทัดคำสั่งและสกิลสำหรับเอเจนต์

ทุกอย่างที่แอปทำกับไฟล์ได้ บรรทัดคำสั่ง `genoffice` ก็ทำได้จากเทอร์มินัล:
ตรวจสอบ แปลง สร้าง อ่าน และแก้ไข Word, Excel, PowerPoint, PDF, Markdown
และ HTML บนเอนจินเดียวกันแบบ headless มันติดตั้งมาพร้อม GenOffice
ไม่ต้องการรันไทม์ของตัวเอง และไม่เคยส่งเอกสารไปที่ใดเลย เมื่อจับคู่กับ
**สกิลสำหรับเอเจนต์** ที่รวมมาให้ มันเปลี่ยนเอเจนต์เขียนโค้ดให้เป็นผู้ทำงาน
เอกสารที่ผลิตไฟล์ Office จริง ไม่ใช่ของเทียบเคียงในรูป Markdown

**ใช้งานได้กับ:** Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot,
OpenCode และ Windsurf ได้ทันทีโดยไม่ต้องตั้งค่าเพิ่ม รวมถึงเอเจนต์อื่นใดที่อ่านสกิลได้
และผ่าน [เซิร์ฟเวอร์ MCP](#mcp-server) ยังรองรับ Claude Desktop และไคลเอนต์ MCP ทุกตัว

### ติดตั้งสกิล

| วิธี                                   | สิ่งที่เกิดขึ้น                                                                                                                                                    |
| -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **การตั้งค่า → การเชื่อมต่อ** ในแอป    | แสดงรายการเอเจนต์ที่พบในคอมพิวเตอร์เครื่องนี้ คลิกเดียวเขียนสกิลลงในแต่ละตัวที่คุณเลือก ปุ่ม **Update** จะปรากฏเมื่อรีลีสของ GenOffice มาพร้อมสกิลเวอร์ชันใหม่กว่า |
| **ดาวน์โหลดเป็น zip** ในหน้าเดียวกัน   | โครงสร้างที่ claude.ai แอปเดสก์ท็อป Claude และผู้ช่วยอื่นๆ ยอมรับเป็นสกิลที่อัปโหลด                                                                                |
| `npx skills add genspark-ai/genoffice` | ติดตั้งจาก repository นี้ลงในเอเจนต์ใดก็ได้ที่รองรับสกิล                                                                                                           |

จากนั้นเริ่มแชทใหม่แล้วขอเอกสารสักฉบับ สกิลจะสอนเอเจนต์ว่าเมื่อใดควรใช้
`genoffice` วิธีอ่านไฟล์ก่อนแก้ไข และวิธีตรวจสอบงานของตัวเอง

### เริ่มต้นอย่างรวดเร็วจากเทอร์มินัล

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

ทุกคำสั่งพิมพ์สรุปหนึ่งบรรทัด หรือออบเจ็กต์ JSON เดียวเมื่อใช้ `--json`
การแก้ไขเป็นแบบอะตอมิก: op ที่ถูกปฏิเสธจะไม่แตะไฟล์เลยและส่งกลับพร้อม
ข้อความแนะนำวิธีแก้ `genoffice help` แสดงรายการคำสั่งทั้งหมดในปัจจุบัน
เอกสารอ้างอิงฉบับเต็มอยู่ที่ [packages/cli/README.md](../../packages/cli/README.md)

### สิ่งที่เอเจนต์รันจริงๆ

เดคระบบสุริยะในเดโมด้านบนใช้พรอมป์เพียงครั้งเดียวใน Claude Code เบื้องหลังนั้น
เอเจนต์ทำตามเวิร์กโฟลว์เป็นขั้นตอนของสกิล และ CLI ตรวจสอบทุกขั้นตอน
ก่อนที่ขั้นตอนถัดไปจะเริ่ม:

```bash
genoffice capabilities --json                        # which cloud tools GenOffice has configured
genoffice guide slides design                        # the deck workflow and layout library
genoffice image "the eight planets in a row …" --aspect 16:9 --out deck/assets/cover.jpg
genoffice slides check deck/outline.json --json      # 8 pages, no findings
genoffice slides check deck/pages/01.json --json     # builds one slide, audits overflow and overlap
…                                                    # one page file per slide, fixed until each check is clean
genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/solar-system.pptx --json
genoffice slides render deck/solar-system.pptx --out deck/shots --json
genoffice slides audit deck/solar-system.pptx --json    # 8 slides, no layout issues
genoffice slides replace deck/solar-system.pptx --slide 4 --spec deck/pages/05.json --json
genoffice open deck/solar-system.pptx
```

ไม่มีการเรียกโมเดลเกิดขึ้นภายใน `genoffice`: เอเจนต์เป็นฝ่ายคิด CLI
เป็นฝ่ายสร้างและตรวจสอบ และผลลัพธ์เปิดใน GenOffice หรือ PowerPoint
เป็นไฟล์ `.pptx` ธรรมดา

<a id="mcp-server"></a>

### เซิร์ฟเวอร์ MCP

คำสั่งชุดเดียวกันนี้ใช้งานได้เป็นเครื่องมือ
[Model Context Protocol](https://modelcontextprotocol.io) เช่นกัน สำหรับ
ผู้ช่วยที่รันเทอร์มินัลไม่ได้ หรือที่คุณไม่อยากให้สิทธิ์เทอร์มินัล มีสองวิธี
ในการเชื่อมต่อ ทั้งสองแบบมาพร้อมโค้ดตัวอย่างที่คัดลอกใช้ได้ทันทีใน
**การตั้งค่า → การเชื่อมต่อ → MCP**:

| วิธี                               | รายละเอียด                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| ---------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **A · `genoffice mcp`** (แนะนำ)    | เซิร์ฟเวอร์ stdio ที่ผู้ช่วยสั่งรันเอง โดยไม่ต้องเปิด GenOffice ค้างไว้ มีเครื่องมือหนึ่งตัวต่อหนึ่งคำสั่ง (`info`, `convert`, `create_docx`, `create_xlsx`, `create_pptx`, `create_pdf`, `docs_read` / `docs_apply` / `docs_check`, `sheet_*`, `slides_*`, `render`, `guide`, `search`, `image`, `media`, `open`) รวมถึงขั้นตอนสร้างเดคแบบเป็นขั้น `deck_start` → `deck_page` → `deck_build` → `deck_replace` ops, spec และ Markdown ส่งแบบ inline ทั้งหมด ไคลเอนต์ที่ไม่มีระบบไฟล์ก็ยังใช้งานได้ |
| **B · เซิร์ฟเวอร์ HTTP ในเครื่อง** | ทำงานอยู่ภายในแอป GenOffice ที่ `http://127.0.0.1:3093/mcp` (Streamable HTTP รองรับ SSE แบบเก่าด้วย) เครื่องมือของมันขับเคลื่อนแท็บเอดิเตอร์ Word ที่มองเห็นได้: `create_session`, `insert_content`, `replace_blocks`, `apply_ops`, `read_document`, `save_session` คุณเฝ้าดูเอกสารค่อยๆ เป็นรูปเป็นร่างได้ ปิดอยู่โดยค่าเริ่มต้น เปิดใช้ได้ในแผงตั้งค่าเดียวกัน                                                                                                                                   |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor, Claude Desktop หรือไคลเอนต์ MCP อื่นใด
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

`genoffice` ที่นี่คือ CLI ตัวเดียวกับที่มาพร้อมแอป (บน macOS อยู่ที่
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`; แผงตั้งค่า
จะแสดง path ที่แน่นอนของเครื่องคุณ) เซิร์ฟเวอร์นี้มีคำแนะนำเวิร์กโฟลว์ของ
ตัวเองและเปิดให้ใช้เอกสารอ้างอิงการทำงานเป็น resource `genoffice://guide/*`
จึงไม่จำเป็นต้องใช้สกิลเลย สกิลกับเซิร์ฟเวอร์ MCP อยู่ร่วมกันได้ โดยผู้ช่วย
เป็นฝ่ายเลือกใช้ตัวใดตัวหนึ่ง ความสามารถบนคลาวด์ (`search`, `image`,
`media`) ยังคงผ่านผู้ให้บริการที่ตั้งค่าไว้ใน GenOffice ส่วนที่เหลือทั้งหมด
ทำงานในเครื่อง และ `GENOFFICE_ALLOWED_ROOTS` จำกัดทุกเครื่องมือให้อยู่ใน
โฟลเดอร์ที่คุณระบุเท่านั้น

เดคพลังงานหมุนเวียนในเดโมด้านบนคือสิ่งที่พรอมป์เดียวใน Claude Code ซึ่ง
เชื่อมต่อกับเซิร์ฟเวอร์ genoffice MCP เพียงตัวเดียว มีหน้าตาเป็นอย่างไรใน
ระดับโปรโตคอล:

```text
capabilities · guide(slides, spec) · guide(slides, design)
search(query) ×4                         → IEA, BNEF and IRENA figures for the slides
search(query, images) ×7 · media(url, ask) ×7
                                         → candidate photos, each one checked to be a real photograph
deck_start(dir, style, outline)          → outline checked: 8 pages to write
deck_page(dir, 0, page) … deck_page(dir, 7, page)
                                         → each page checked against the outline and the palette; one page sent again
deck_build(dir, out)                     → renewables-2026.pptx, no image failures
slides_audit(file) · slides_render(file, out)
                                         → no layout findings; 8 PNGs come back as image content
deck_replace(dir, n, page) ×3 · slides_render(file, out)
                                         → three pages fixed after looking at the renders
```

สามสิบแปดครั้งของการเรียก ประมาณสิบสามนาที และผู้ช่วยไม่แตะเทอร์มินัลเลยสักครั้ง: ข้อมูล ภาพถ่าย คู่มือ การตรวจสอบ และผลเรนเดอร์ทั้งหมดเดินทางมาในรูปผลลัพธ์ของเครื่องมือ MCP มีเพียง `search` และ `media` เท่านั้นที่ออกจากเครื่อง ไปยังผู้ให้บริการที่ตั้งค่าไว้ใน GenOffice

<a id="download"></a>

## ดาวน์โหลด

| แพลตฟอร์ม                            | ความต้องการของระบบ                                    | ดาวน์โหลด                                                                                  |
| ------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [`.dmg` ล่าสุด (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)          |
| **macOS** — Intel (x64)              | macOS 11+                                             | [`.dmg` ล่าสุด (x64)](https://github.com/genspark-ai/genoffice/releases/latest)            |
| **Windows** (x64, พีซีทั่วไป)        | Windows 10+, Intel/AMD                                | [ตัวติดตั้ง `-x64.exe` ล่าสุด](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** บน Arm (ARM64)           | Windows 11 on Arm (Snapdragon X และรุ่นใกล้เคียง)     | [ตัวติดตั้ง `-arm64.exe` ล่าสุด](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 ขึ้นไป)             | [`.deb` ล่าสุด](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [`.rpm` ล่าสุด](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — ดิสโทรอื่นๆ              | x86_64, glibc 2.34+, FUSE 2                           | [`.AppImage` ล่าสุด](https://github.com/genspark-ai/genoffice/releases/latest)             |

บิลด์ทั้งหมดมาจาก `main` ตัวติดตั้ง macOS และ Windows ผ่านการลงนามแล้ว
เวอร์ชันเก่ากว่านี้ดูได้ที่หน้า [Releases](https://github.com/genspark-ai/genoffice/releases)

<details>
<summary><b>การติดตั้งบน Linux</b></summary>

ไฟล์ deb ติดตั้งด้วย apt — จะดึงดีเพนเดนซีที่ต้องใช้และเพิ่ม GenOffice
ลงในเมนูแอปพลิเคชันให้อัตโนมัติ:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

บน Fedora / ตระกูล RHEL / openSUSE ให้ติดตั้งไฟล์ rpm แทน:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

ไฟล์ AppImage รันได้ทันทีในตำแหน่งที่อยู่: ติดตั้งรันไทม์ FUSE 2
(`sudo apt install libfuse2`; บน Ubuntu 24.04 แพ็กเกจคือ `libfuse2t64`)
กำหนดสิทธิ์ให้ไฟล์รันได้ แล้วรัน:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## หลักการทำงาน

แอป Electron เจ็ดตัว — Docs, Sheets, Slides, PDF, Markdown, HTML และเชลล์
แบบแท็บ — ใช้เอนจินร่วมกันเป็นชุดแพ็กเกจ TypeScript แท้ๆ บวก sidecar Rust
สำหรับ `.xlsx` ไฟล์ต้นฉบับคือแหล่งข้อมูลจริงเสมอ: การแก้ไขจะถูกใช้เป็นแพตช์
เฉพาะจุด และทุกอย่างที่เอดิเตอร์ไม่ได้แตะจะคงเดิมทุกรอบการเปิด-บันทึก

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

ทัวร์แพ็กเกจแต่ละตัว (เอนจิน docx/pptx, `pdf2docx`, `html2docx`, แกนหลัก
ของเอเจนต์ และผู้ให้บริการ) อยู่ใน [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages)

## การพัฒนา

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

แอป sheets ยังต้องการ Rust toolchain สำหรับ xlsx sidecar เพิ่มเติม
(`cargo` ต้องอยู่ใน PATH); `npm run build -w @genoffice/sheets` จะคอมไพล์
ให้อัตโนมัติ ดู [CONTRIBUTING.md](../../CONTRIBUTING.md) สำหรับการตรวจสอบ
ที่การเปลี่ยนแปลงทุกครั้งต้องผ่าน และวิธีที่ pull request จะถูกรวมเข้าโปรเจกต์

## คอมมูนิตี้

GenOffice อยู่ในระหว่างการพัฒนาอย่างต่อเนื่อง และความคิดเห็นของคุณ
มีส่วนกำหนดทิศทางของมัน

- **รายงานบั๊กหรือขอฟีเจอร์ใหม่** ได้ที่
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues)
- **เข้าร่วมกลุ่มแชทของ GenOffice** ได้ที่
  [GenTeam](https://genoffice.ai/join) เพื่อพูดคุยกับทีมงานและผู้ใช้คนอื่นๆ
- **ให้ดาว repo นี้** หาก GenOffice มีประโยชน์กับคุณ — เป็นวิธีที่ดีที่สุด
  ในการสนับสนุนโปรเจกต์นี้

## คำถามที่พบบ่อย

<details>
<summary><b>GenOffice ฟรีหรือไม่</b></summary>

ใช่ GenOffice ฟรีและโอเพนซอร์สภายใต้สัญญาอนุญาต Apache-2.0 — ไม่มีช่วง
ทดลองใช้ ไม่มีระดับที่ต้องเสียเงินสำหรับตัวแอปเอง

</details>

<details>
<summary><b>GenOffice เปิดไฟล์ Microsoft Word, Excel และ PowerPoint ได้หรือไม่</b></summary>

ได้ GenOffice เปิดและบันทึกไฟล์ `.docx`, `.xlsx` และ `.pptx` แบบดั้งเดิม
การบันทึกคงบิตต่อบิต: ส่วนของไฟล์ที่คุณไม่ได้แตะจะถูกเขียนกลับเหมือนเดิม
ทุกไบต์ เอกสารจึงยังใช้งานได้ต่อใน Microsoft Office

</details>

<details>
<summary><b>GenOffice ใช้งานแบบออฟไลน์ได้หรือไม่</b></summary>

การแก้ไขเอกสารทำงานในเครื่องทั้งหมด — ไฟล์ไม่ออกจากเครื่องของคุณเลย
เมื่อเปิด แก้ไข บันทึก หรือแปลง ฟีเจอร์ AI (เอเจนต์ การค้นหา เครื่องมือภาพ)
ต้องมีการเชื่อมต่อเครือข่าย โดยลงชื่อเข้าใช้ด้วย Genspark หรือใช้คีย์ API
โมเดลของคุณเอง

</details>

<details>
<summary><b>GenOffice แก้ไขไฟล์ PDF ได้หรือไม่</b></summary>

ได้ — การแก้ไขข้อความและภาพใน PDF จริง ที่เขียน content stream ของหน้า
ใหม่โดยคงฟอนต์เดิมไว้ ไม่ใช่การแปะ annotation ทับ

</details>

<details>
<summary><b>GenOffice แปลง PDF เป็น Word, Excel หรือ PowerPoint ได้หรือไม่</b></summary>

ได้ — ทำงานในเครื่องทั้งหมด: การดึงข้อมูลระดับตัวอักษรด้วย PDFium
ร่วมกับการวิเคราะห์เลย์เอาต์แบบตามเรขาคณิต ไม่ต้องใช้บริการคลาวด์
ไม่ต้องอัปโหลด หน้าที่สแกนมาก็รองรับด้วย: บน macOS และ Windows ระบบ OCR
จะอ่านให้ จึงแปลงเป็นข้อความที่แก้ไขได้ ไม่ใช่ภาพของหน้ากระดาษ

</details>

<details>
<summary><b>ใช้โมเดล AI หรือคีย์ API ของตัวเองได้หรือไม่</b></summary>

ได้ นอกจากการลงชื่อเข้าใช้ด้วย Genspark ที่ไม่ต้องใช้คีย์แล้ว GenOffice
ยังรองรับการใช้คีย์ของคุณเองกับ Claude, OpenAI, Gemini, DeepSeek, Kimi,
GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty และ OpenCode Zen/Go
พร้อม endpoint ที่รองรับ OpenAI-compatible ใดๆ — รวมถึงเซิร์ฟเวอร์โมเดล
ในเครื่องด้วย การค้นหา การสร้างภาพ และการวิเคราะห์ภาพ/วิดีโอ ใช้คีย์
แยกของตัวเองภายใต้ Settings → AI Media & Search

</details>

<details>
<summary><b>GenOffice แปลง HTML เป็น Word ได้หรือไม่</b></summary>

ได้ — Export as Word ในแอป HTML สร้าง `.docx` ที่แก้ไขได้แบบเนทิฟทั้งหมด
ในเครื่อง หน้าเว็บถูกเรนเดอร์ใน Chromium ที่มีในตัวและลดรูปลงเป็นโครงสร้าง
Word จริง: หัวเรื่อง พารากราฟ รายการ ตาราง การ์ด แถวข้อมูล KPI ฟิลด์ฟอร์ม
และพื้นหลังของหน้า มีเพียงองค์ประกอบภาพที่ไม่มีสิ่งเทียบเท่าใน Word
(กราฟ ไอคอน กล่องที่ตกแต่ง) เท่านั้นที่จะถูกฝังเป็นภาพ

</details>

<details>
<summary><b>สั่งงาน GenOffice จาก Claude Code, Codex, Cursor หรือสคริปต์ได้หรือไม่</b></summary>

ได้ GenOffice ติดตั้งบรรทัดคำสั่ง `genoffice` ที่รันเอนจินเดียวกันแบบ headless:
ตรวจสอบ แปลง สร้าง อ่าน และแก้ไขเอกสารจากเทอร์มินัลหรือสคริปต์ พร้อมเอาต์พุต
`--json` สำหรับโปรแกรม สกิลสำหรับเอเจนต์ที่รวมมาให้สอน Claude Code, Codex,
Cursor, Gemini CLI, GitHub Copilot, OpenCode และ Windsurf ให้ใช้งานมัน
ติดตั้งได้จาก **การตั้งค่า → การเชื่อมต่อ** ดู
[บรรทัดคำสั่งและสกิลสำหรับเอเจนต์](#command-line-and-agent-skill)

</details>

<details>
<summary><b>GenOffice เก็บข้อมูลอะไรบ้างหรือไม่</b></summary>

บิลด์ที่แพ็กเกจอย่างเป็นทางการส่งข้อมูลวิเคราะห์การใช้งานแบบจำกัดเป็นค่า
เริ่มต้น และคุณสามารถปิดการรายงานได้ทุกเมื่อที่ Settings → General ข้อมูล
วิเคราะห์จะไม่ส่งเนื้อหาเอกสาร ชื่อไฟล์ พาธไฟล์ ข้อมูลบัญชี หรือที่อยู่อีเมล
ดู [GenOffice Privacy](../../PRIVACY.md) สำหรับรายละเอียดเหตุการณ์และข้อมูล
ทั้งหมดที่เปิดเผย

</details>

## ความปลอดภัย

ดู [SECURITY.md](../../SECURITY.md) สำหรับสถานะความปลอดภัยของโปรเซส
(การแซนด์บ็อกซ์ renderer, การตรวจสอบ IPC, การกั้นลิงก์ภายนอก) และโมเดล
ภัยคุกคามสำหรับเนื้อหาที่สร้างโดย AI

## คำขอบคุณ

GenOffice จะไม่สามารถเกิดขึ้นได้หากไม่มีโครงการโอเพนซอร์สเหล่านี้:

- [Electron](https://www.electronjs.org/) — รันไทม์เดสก์ท็อปสำหรับทุกแอป
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — คอร์ของ UI
  สเปรดชีตที่ Sheets ต่อยอดมา
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, ฝังผ่าน
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) — เอนจิน
  content-stream ที่อยู่เบื้องหลังการแก้ไขข้อความและภาพ PDF ของจริง
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) และ
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — การเรนเดอร์ PDF
  และการประกอบเอกสาร
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  บล็อกเอดิเตอร์ใน Docs และ Markdown
- [CodeMirror](https://codemirror.net/) (MIT) — ซอร์สเอดิเตอร์ใน HTML
- [Konva](https://konvajs.org/) — การเรนเดอร์บนแคนวาสสำหรับ Slides และกราฟ
  ใน Sheets
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — มาตรวัดการ
  shaping ข้อความสำหรับสคริปต์ที่ซับซ้อน
- [calamine](https://github.com/tafia/calamine) และ
  [IronCalc](https://github.com/ironcalc/IronCalc) — เลเยอร์การอ่านและ
  การคำนวณของ xlsx sidecar ที่เขียนด้วย Rust
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — ตัวถอดรหัส
  MicroType Express สำหรับฟอนต์ PowerPoint ที่ฝังมา ซึ่งพอร์ตมาเป็น
  TypeScript
- [React](https://react.dev/) (MIT) — เลเยอร์ UI ของทุกแอป
- [Mermaid](https://mermaid.js.org/) (MIT) และ [KaTeX](https://katex.org/)
  (MIT) — แผนภาพและสมการคณิตศาสตร์ใน Markdown และ Docs
- [opentype.js](https://opentype.js.org/) (MIT) — การแยกวิเคราะห์ฟอนต์สำหรับ
  มาตรวัดและการค้นหา glyph
- [JSZip](https://stuk.github.io/jszip/) (MIT) และ
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — เลเยอร์คอนเทนเนอร์ OOXML และ XML
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — ชุดไอคอนทั่วทั้งริบบอน
- [electron-updater](https://www.electron.build/) (MIT) — การอัปเดตในแอป
- ฟอนต์ Liberation, Carlito, Caladea และ Noto CJK (OFL/Apache-2.0) —
  ฟอนต์เอกสารที่รวมมาให้

`npm run notices` จะสร้างสรุปสัญญาอนุญาตของบุคคลที่สามที่รวมมาให้ใหม่
(`tools/gen-third-party-notices.mjs`); ดีเพนเดนซีรันไทม์ทั้งหมดอยู่ภายใต้
MIT/Apache-2.0/BSD-3-Clause/OFL

## สัญญาอนุญาต

GenOffice อยู่ภายใต้สัญญาอนุญาต [Apache License 2.0](../../LICENSE)
โดยมีข้อยกเว้นหนึ่งข้อ: ไดเรกทอรี `ee/` สงวนไว้สำหรับโมดูลระดับองค์กร
ในอนาคต และอยู่ภายใต้ [GenOffice Enterprise License](../../ee/LICENSE)

ชื่อและโลโก้ของ GenOffice และ Genspark เป็นเครื่องหมายการค้าของ
Mainfunc, Inc. สัญญาอนุญาต Apache-2.0 ไม่ได้ให้สิทธิ์ในการใช้เครื่องหมาย
เหล่านี้ (ดูข้อ 6) โปรเจกต์ที่ fork ไปควรใช้แบรนด์ของตัวเอง
