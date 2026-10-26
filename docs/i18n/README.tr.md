<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — açık kaynaklı yapay zekâ Office paketi: yerleşik yapay zekâ paneliyle Docs, Sheets, Slides, PDF, Markdown ve HTML" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>Dünyanın ilk tam özellikli açık kaynaklı yapay zekâ Office paketi.</b><br>
Word, Excel, PowerPoint ve PDF dosyaları; siz ve yapay zekânız tarafından düzenlenir, gerçek biçimlerinde geri kaydedilir.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Lisans: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Son sürüm"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="İndirmeler"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub yıldızları"></a>
  <a href="https://x.com/merrickbuilds"><img src="https://img.shields.io/badge/follow-%40merrickbuilds-000000?logo=x&logoColor=white" alt="X'te @merrickbuilds'i takip edin"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a> · <a href="README.vi.md">Tiếng Việt</a> · <b>Türkçe</b></p><!-- lang-switcher · public-hygiene: allow -->

<p align="center">
  <a href="#download"><b>İndir</b></a> ·
  <a href="#skills"><b>Skills</b></a> ·
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
  <a href="https://genoffice.ai/"><b>Web sitesi</b></a> ·
  <a href="https://genoffice.ai/join"><b>Topluluk</b></a> ·
  <a href="https://x.com/merrickbuilds"><b>X</b></a> ·
  <a href="../../PRIVACY.md"><b>Gizlilik</b></a>
</p>

GenOffice; macOS, Windows ve Linux için ücretsiz, açık kaynaklı bir Microsoft
Office alternatifidir. Yerel `.docx`, `.xlsx` ve `.pptx` dosyalarını açar ve
kaydeder, PDF, Markdown ve HTML düzenler ve her belgenin yanına bir yapay zekâ
ajanı koyar — kenara iliştirilmiş bir sohbet kutusu değil, dosyayı okuyan,
değişikliği yapan ve tam olarak neye dokunduğunu size gösteren bir düzenleyici.
ODF dosyaları (`.odt`, `.ods`, `.odp` ve `.ott`/`.ots`/`.otp` şablonları)
düzenlenebilir OOXML kopyaları olarak açılır; içe aktarma yalnızca açma
yönündedir — kaydetme her zaman OOXML yazar, asla ODF yazmaz.

- **Gerçek biçimler, bayt bayt korunur.** Yalnızca düzenlediğiniz kısım yeniden
  yazılır. Dosyadaki diğer her şey bayt bayt korunur; böylece belgeler Word,
  Excel ve PowerPoint'te çalışmaya devam eder.
- **Gözden geçirebileceğiniz bir yapay zekâ.** Düzenlemeler, tek tıkla geri
  alınabilen değişiklik izleme ve fark görünümleri olarak gelir. Elektronik
  tablolar yapıştırılmış sayılar değil, canlı formüller alır. Sunumlar ve
  sayfalar doğrudan tuvale üretilir ve tamamen düzenlenebilir kalır. Bir Docs
  seçimini veya bir Sheets hücre aralığını modelden gizleyin; yapay zekâ
  yalnızca adlandırılmış bir yer tutucu görür, özgün sözcükler dosyada kalır ve
  geri gelir. Gizleme bu uygulamanın yapay zekâsını kapsar — CLI, MCP araçları
  ve başsız dışa aktarma dosyayı doğrudan okur ve bu sözcükleri görmeye devam
  eder.
- **Tasarımı gereği yerel.** Dosyalar makinenizde açılır, düzenlenir,
  kaydedilir ve dönüştürülür. PDF → Word / Excel / PowerPoint, Markdown → Word
  ve HTML → Word dönüşümlerinin tamamı cihaz üzerinde çalışır. Makineden
  yalnızca yapay zekâ çağrıları, seçtiğiniz sağlayıcıya gider.
- **Dosyaları söyledikleriyle bulun.** Ana ekran; `.docx`, `.xlsx`, `.pptx`,
  PDF, Markdown ve HTML dosyalarınızın adlarını, klasörlerini ve tam metnini,
  CJK dahil, yerel bir SQLite dizininden arar. İsteğe bağlı olarak en iyi
  sonuçlar, System One yargı modeli **[TypeSafe Jev](https://typesafe.ai/)**
  tarafından yeniden sıralanır; böylece sorunuzu yanıtlayan dosya en üste
  gelir.
- **Dosyalarınıza sorun, kaynaklarıyla.** Bir belge kümesine yıldız ekleyin;
  ana ekrandaki **Ask AI** soruları bu belgeler üzerinden yanıtlar ve her
  iddianın arkasındaki dosyayı kaynak gösterir, böylece açıp kontrol
  edebilirsiniz.
- **Kendi anahtarlarınız ya da hiç anahtar yok.** Genspark ile oturum açıp
  anahtar işini atlayın ya da Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM,
  Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty, Opper, Cheaper
  Inference, Atlas Cloud veya yerel sunucular dahil herhangi bir OpenAI uyumlu
  uç nokta için kendi anahtarınızı getirin.
- **Betiklenebilir ve ajanlara hazır.** Uygulama bir `genoffice` komut satırı
  ve Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode ve
  Windsurf için bir skill ile gelir; böylece bir kodlama ajanı, hiçbir pencere
  açmadan makinenizde gerçek Office dosyaları oluşturabilir, dönüştürebilir,
  okuyabilir ve düzenleyebilir.

**Edinin:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon ve Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 ve Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
ayrıntılar ve gereksinimler [İndir](#download) bölümünde.

## Demo

Altı uygulama, tek bir yapay zekâ paneli, TypeSafe Jev ile yeniden sıralanan
bir dosya araması, yıldızlı dosyalarınızdan kaynak göstererek yanıt veren bir
Ask AI ve kodlama ajanınız için bir komut satırı. Her ekran görüntüsü macOS
üzerinde gerçek uygulamadır; yapay zekâ, panelde okuyabileceğiniz istemle
yönlendirilmiştir.

### 1 · Docs — gözden geçirebileceğiniz bir yapay zekâyla `.docx` açın ve düzenleyin

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs; tam genişlikte kapak görseli, gölgeli KPI tablosu, üst bilgi ve alt bilgi içeren iki sütunlu bir yıllık rapor sayfasını, yapay zekâ paneli daraltılmış hâlde %80 yakınlaştırmada gösteriyor"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: banner görselli bir şirket tanıtımı; yapay zekâ Genel Bakış bölümünü sıkılaştırdı ve madde işaretli yeni bir bölüm ekledi, panel tek tıkla geri alma sunuyor"></td>
</tr>
<tr>
<td><b>Dosyayı Word'ün yerleştirdiği gibi açar</b> — iki sütunlu bölümler, taşmalı görseller, gölgeli tablolar, üst ve alt bilgiler, Word'ün satır ölçülerine göre sayfalama. Stiller, yorumlar, izlenen değişiklikler, denklemler, mürekkep ve içerik denetimleri dokunulmadan gidip gelir. <b>Mailings</b> sekmesi belgeyi bir birleştirme şablonuna dönüştürür: bir <code>.xlsx</code> veya <code>.csv</code> alıcı listesi seçin, gerçek <code>MERGEFIELD</code> alanları ekleyin, kayıt kayıt önizleyin ve tek bir belgeye ya da kayıt başına bir <code>.docx</code> dosyasına birleştirin.</td>
<td><b>Düzenlemeyi isteyin</b> — yapay zekâ ihtiyaç duyduğu blokları okur, Genel Bakış'ı yeniden yazar ve madde işaretli yeni bir bölüm ekler. Her yapay zekâ turu geri alabileceğiniz bir anlık görüntüdür; <b>Track changes</b> açıkken düzenlemeler Word tarzı revizyonlar olarak gelir. Bunun yerine bir zihin haritası isteyin; yanıt, tüm belgenin etkileşimli bir diyagramı olur: daraltılabilir dallar, yakınlaştırma ve kaydırma, bir düğüme tıklamak takip sorusu sorar ve alıntılar geldikleri bloğa geri atlar.</td>
</tr>
</table>

### 2 · Sheets — yapıştırılmış sayılar değil, canlı formüller ve grafiklerle `.xlsx`

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: yapay zekâ, SUMIF formülleriyle bölge ve kategoriye göre geliri içeren bir Summary sayfası ve bir sütun grafiği ekledi ve bir Geri Al düğmesiyle 43 uygulanan değişiklik bildiriyor"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: 2. çeyrek gelirinde hangi bölgenin önde olduğu sorulduğunda yapay zekâ, kategori dökümüyle birlikte Avrupa yanıtını veriyor ve kullandığı hücreleri Orders sayfasının yanında bağlantı olarak gösteriyor"></td>
</tr>
<tr>
<td><b>Oluşturun</b> — tek bir cümleden ajan, bölge ve kategoriye göre gerçek <code>SUMIF</code> formülleri içeren bir Summary sayfası ekler, bir sütun grafiği yerleştirir ve 43 değişikliği tek bir geri alınabilir toplu işlem olarak uygular.</td>
<td><b>Sorun</b> — çalışma kitabıyla ilgili sorular, gerekçesiyle ve kullanılan hücreler tıklanabilir alıntılar olarak geri döner. Arka planda: kendi geliştirdiğimiz Rust tabanlı <code>.xlsx</code> motoru, özet tablolar, dilimleyiciler, koşullu biçimlendirme, formül izleme ve hücrelere Excel'in yerel zengin değerleri olarak gömülen resimler — Excel onları yeniden açtığında hâlâ hücrenin içindedir.</td>
</tr>
</table>

### 3 · Slides — bir istemden `.pptx` sunumuna

<img src="../assets/readme/slides-generate.webp" alt="GenOffice Slides'ın Aurora Home yatırımcı sunumunu üretmesinin hızlandırılmış kaydı: yapay zekâ panelde hikâye akışını planlıyor, slaytlar tuvalde art arda beliriyor ve bitmiş sunum kapanış çağrısıyla sona eriyor" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: yapay zekâ tarafından üretilen Aurora Home yatırımcı sunumunun kapak slaytı tuvalde, özgün tek satırlık istem ve yapay zekânın ne oluşturduğuna dair özeti panelde"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: aynı 11 slaytlık sunumun tasarlanmış kapanış slaytı, solda küçük resim şeridi ve hikâye akışını özetleyen yapay zekâ paneliyle"></td>
</tr>
<tr>
<td><b>Tek satır girdi</b> — "Aurora Home için 10 slaytlık bir yatırımcı sunumu oluştur…". GenOffice hikâye akışını planlar, rakamları araştırır ve her slaytı gerçek bir <code>.pptx</code> olarak tuvale taslaklar.</td>
<td><b>Bitmiş bir sunum çıktı</b> — tutarlı tipografi, görseller ve kapanış eylem çağrısıyla on bir tasarlanmış slayt; asıl slaytlar, düzenler, akıllı kılavuzlar ve yıkıcı olmayan kırpma ile düzenlemeye devam edin ya da panelden yeniden biçimlendirmesini, yeniden yazmasını ve yeniden sıralamasını isteyin. Bir <code>.glb</code> ekleyin; gerçek bir 3B modeldir: düzenleyicide sürükleyerek döndürün, PowerPoint'in 3B sakladığı biçimde saklanır, PowerPoint'te de 3B kalır.</td>
</tr>
</table>

**Slide Show → Record**, sunumu slayt slayt, duraklatma ve devam etme
desteğiyle seslendirir. Anlatım ve ölçtüğü otomatik ilerleme süreleri
`.pptx` içine kaydedilir ve PowerPoint bunları aynı şekilde oynatır.

### 4 · PDF — PDF metnini yerinde düzenleyin, PDF'yi cihaz üzerinde Word'e dönüştürün

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: Metin düzenleme modu, yerinde düzenleme için sayfadaki her metin bloğunu çerçeveliyor; yapay zekâ paneli raporla ilgili bir soruyu sayfa alıntılarıyla yanıtlıyor"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs, Helios üç aylık değerlendirme PDF'sinden yerel olarak dönüştürülmüş bir Word belgesini, özgün PDF'nin yanındaki ikinci bir sekmede gösteriyor"></td>
</tr>
<tr>
<td><b>Sayfanın içinde düzenleyin</b> — Metin düzenleme modu, yerinde yeniden yazmak için her metin bloğunu çerçeveler; içerik akışı, üstünü örten bir açıklama notuyla değil, özgün yazı tipleriyle PDFium üzerinden yeniden yazılır. Yapay zekâya uzun bir rapor hakkında soru sorun ve sayfa alıntılı yanıtlar alın.</td>
<td><b>Cihaz üzerinde dönüştürün</b> — <b>PDF Converter → PDF to Word</b>, kaynağın yanında Docs'ta açılan, başlıkları, istatistik satırları ve paragrafları bozulmamış düzenlenebilir bir <code>.docx</code> üretir. Excel ve PowerPoint hedefleri aynı şekilde çalışır; taranmış sayfalar sistem OCR'ından geçer.</td>
</tr>
</table>

### 5 · HTML — önce tasarım özeti hazırlayan bir yapay zekâ sayfa ve arayüz oluşturucu

Sayfanın ne için ve kimin için olduğunu söyleyin. Yapay zekâ önce bir **tasarım
özeti** önerir — kanca, palet, tipografi ve stil yönleri — ardından bu
belirteçlere göre tek, kendi kendine yeten bir `.html` dosyası oluşturur.

<img src="../assets/readme/html-restyle-motion.webp" alt="GenOffice HTML'nin Lumen açılış sayfasını yeniden biçimlendirmesinin hızlandırılmış kaydı: paneldeki tek bir Restyle isteği, karanlık Midnight Studio sayfasını sıcak Solar Daybreak sürümüne dönüştürürken her bölüm ve tüm metin yerinde kalıyor" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: güneş enerjili bir masa lambası için karanlık Midnight Studio yönünde üretilmiş bir açılış sayfası, canlı önizlemede ve az önce oluşturduğu sayfayı özetleyen yapay zekâ paneliyle gösteriliyor"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Aynı Lumen açılış sayfası yapay zekâ tarafından sıcak Solar Daybreak yönüne yeniden biçimlendirilmiş: kâğıt arka plan, serifli başlıklar ve turuncu bir vurgu; her bölüm ve tüm metin korunmuş"></td>
</tr>
<tr>
<td><b>Tek bir istemden üretildi</b> — Lumen için Midnight Studio yönünde oluşturulmuş cesur bir hero alanı, özellik kartları, fiyatlandırma ve bekleme listesi formu. Yeniden biçimlendirmek için herhangi bir öğeye tıklayın, metni düzenlemek için çift tıklayın ya da CodeMirror kaynak görünümüne geçin.</td>
<td><b>Aynı tasarım, yeni yön</b> — tek bir <b>Restyle</b> isteği özetin belirteçlerini değiştirir ve sayfa buna uyar: sıcak kâğıt, editoryal serif, güneş turuncusu vurgu, hiçbir şey yeniden yazılmaz. Tam ekran sunun ya da PDF veya yerel, düzenlenebilir bir Word belgesi olarak dışa aktarın.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: serbest çalışan bir tasarımcı için sıcak keten tarzında üretilmiş kişisel bir gösterge paneli arayüzü; sol şerit, serifli karşılama ve dört metrik kartıyla"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: geniş sayfa gazete tarzında üretilmiş bir elektrikli araç pazarı veri raporu; serifli bir başlık bandı, 17,3 milyonluk bir manşet rakamı ve bir istatistik satırıyla"></td>
</tr>
<tr>
<td><b>Arayüz taslakları</b> — "kişisel gösterge paneli" başlangıç şablonu bir personayı çalışan bir düzene dönüştürür: sol şerit, karşılama, faturalandırılabilir saat çizgi grafiği, fatura ve kullanım kartları; hepsi bir geliştiriciye teslim edebileceğiniz gerçek HTML.</td>
<td><b>Veri hikâyeleri</b> — "veri raporu" başlangıç şablonu editoryal bir geniş sayfa oluşturur: serifli başlık bandı, tek bir manşet rakamı, çizgilerle ayrılmış bir istatistik satırı, satır içi SVG grafikler ve bir yöntem notu.</td>
</tr>
</table>

### 6 · Markdown — düz `.md` üzerinde bir blok düzenleyici, Ask AI ile

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: seçili bir paragrafta, yazılmış bir talimat ve Polish, Make more concise, Expand ve Fix grammar gibi öneri çipleri ile Send now ve Add to queue düğmeleri içeren bir Ask AI açılır penceresi görünüyor"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown; bir tablo, bir Mermaid akış şeması ve bir görev listesi içeren lansman notları belgesini, solda yapay zekâ panelinin başlangıç istemleriyle gösteriyor"></td>
</tr>
<tr>
<td><b>Bir seçim hakkında yapay zekâya sorun</b> — herhangi bir pasajı seçin, bir <b>Ask AI</b> çipi belirir: bir talimat yazın ya da bir öneri seçin, hemen gönderin ya da birkaç bağlantılı düzenlemeyi kuyruğa alıp tek seferde çalıştırın. Aynı giriş her uygulamada vardır.</td>
<td><b>İşlenmiş görünür, düz Markdown olarak kaydedilir</b> — bir Tiptap blok düzenleyicisinde başlıklar, listeler, tablolar, görseller, kod blokları ve Mermaid diyagramları; düz <code>.md</code> olarak geri yazılır, tamamen yerel bir <b>Markdown → Word</b> dışa aktarımıyla.</td>
</tr>
</table>

**Obsidian'ın yaptığı gibi bağlantılı notlar.** `[[wiki links]]` adlandırdığı
nota tıklanarak gidilir, `[[` yazmak not adı tamamlama sunar ve `![[…]]` bir
notu ya da resmi satır içine gömer. **Backlinks** paneli, açık nota işaret eden
her notu listeler ve **Graph** görünümü tüm kasanın bağlantılarını haritalar.

### 7 · Arama — soruyu yanıtlayan dosyayı bulun, TypeSafe Jev ile

Çalışma klasörünüzdeki her dosya cihaz üzerinde dizinlenir: adlar, klasörler
ve Word, Excel, PowerPoint, PDF, Markdown ve HTML dosyalarından çıkarılan
metin, CJK'ya duyarlı belirteçleme yapan bir SQLite tam metin dizininde. **Jev
search reranking** seçeneğini açın; en iyi 20 yerel sonuç,
[TypeSafe Jev](https://typesafe.ai/blog/introducing-system-one-models-and-jev)
tarafından, yani metin üretmek yerine tek bir çağrıda belge başına tek bir
kalibre edilmiş alaka puanı döndüren System One modeliyle değerlendirilir.
Liste bu puana göre yeniden sıralanır; çağrı başarısız olur veya zaman aşımına
uğrarsa yerel sıra korunur.

<img src="../assets/readme/search-jev-motion.webp" alt="GenOffice ana ekranının ekran kaydı: laptop refresh policy yazıldığında önce bir tarayıcı önbelleği politikası, bir marka yenileme planı ve bir gösterge paneli yenileme takvimi listeleniyor, ekipman standartları belgesi ise en sonda; Ayarlar sayfası, AI Media & Search altında TypeSafe uç noktasıyla Jev search reranking seçeneğinin açıldığını gösteriyor; aynı arama ardından ekipman standartları belgesini, sonuç sayısının yanındaki Jev rozetiyle ilk sırada gösteriyor" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/search-jev-before-after.webp" alt="laptop refresh policy sorgusu için yan yana iki sonuç listesi: Jev olmadan gösterge paneli yenileme takvimi, marka yenileme planı ve tarayıcı önbelleği yenileme politikası önde, şirket ekipman standartları belgesi beşinci; Jev ile üç yıllık dizüstü bilgisayar değiştirme döngüsünü belirten ekipman standartları belgesi ilk sırada"></td>
<td width="50%"><img src="../assets/readme/search-jev-settings.webp" alt="GenOffice Ayarlar, AI Media & Search sayfası: Jev search reranking anahtarı, OpenRouter ile TypeSafe arasındaki uç nokta seçimi ve API anahtarı alanını içeren Local file search bloğu"></td>
</tr>
<tr>
<td><b>Aynı sözcükler, farklı yanıtlar</b> — "laptop refresh policy" bir tarayıcı önbelleği yenileme politikası, bir marka yenileme planı ve bir gösterge paneli yenileme takvimiyle sözcük sözcük eşleşir; bu yüzden tam metin sıralaması onları öne koyar. Jev alıntıları okur ve üç yıllık değiştirme döngüsünü belirten ekipman standartları belgesini en üste taşır. Sonuç sayısının yanındaki <b>Jev</b> rozeti, sıranın modelden geldiğini gösterir.</td>
<td><b>Varsayılan olarak kapalı, tek bir anahtarla açılır</b> — Ayarlar → AI Media & Search → Local file search. OpenRouter'ı ya da doğrudan TypeSafe'i seçin, bir anahtar yapıştırın, Test connection'a basın. Yalnızca anahtar açıkken en iyi sonuçların alıntıları cihazdan çıkar; dizinin kendisi asla çıkmaz.</td>
</tr>
</table>

### 8 · Ask AI — yıldızlı dosyalarınız, kaynaklarıyla yanıtlanır

Birlikte çalıştığınız dosyalara yıldız ekleyin; ana ekrandaki **Ask AI** onları
bir bilgi tabanı olarak ele alır. Sorunuzu doğal dille sorun; ajan yıldızlı
dosyaları arar, ihtiyaç duyduklarını okur ve her iddiadan sonra bir alıntıyla
yanıtlar. Bir alıntıya tıklayın; dosya doğru uygulamada açılır. Soruya
sürükleyerek, yapıştırarak ya da **+** menüsünden dosya ekleyin — görseller
modele görsel girdi olarak gider — ve her soru kendi konuşması olarak
saklanır: aralarında geçiş yapın, yenisini başlatın, birini silin; liste
yeniden başlatmadan sonra da korunur.

<img src="../assets/readme/home-ask-ai-motion.webp" alt="GenOffice Ask AI ekran kaydı: ana ekrana kısa özet, 3. çeyrek büyümemizi ne sürükledi ve Asya-Pasifik nasıl performans gösterdi sorusu yazılıyor; ajan iki yıldızlı dosyayı okuyup Aurora X1 lansmanı ve Asya-Pasifik genişlemesiyle yanıtlıyor, her iddiayı üç aylık raporu ya da bölgesel satış çalışma kitabını adlandıran bir alıntı çipi izliyor; çalışma kitabı çipine tıklamak onu Asya-Pasifik satırı vurgulanmış hâlde Sheets'te açıyor" width="100%">

Ask AI,
[Arama](#7--search--find-the-file-that-answers-the-question-with-typesafe-jev)
ile aynı cihaz üzeri tam metin dizininde arar ve dosyaları sayfa sayfa okur;
böylece model yalnızca isabetleri ve istediği pasajları, seçtiğiniz sağlayıcı
üzerinden görür. Ayrıca dosyaları doğrudan oluşturmak için yerleşik
`genoffice` CLI'yi çağırabilir; her biri yanıtın altında bir yapıt kartı
olarak görünür ve bir tıklamayla açılır.

### 9 · CLI — kodlama ajanınız GenOffice'i makinenizde yönetir

GenOffice bir `genoffice` komut satırı ve bir ajan skill'i ile gelir. Skill'i
kurun; Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode ya da
Windsurf, hiçbir pencere açmadan uygulamalarla aynı motorlar üzerinden gerçek
Office dosyaları oluşturabilir, dönüştürebilir, okuyabilir ve düzenleyebilir.

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides, bir kodlama ajanının genoffice komut satırı üzerinden oluşturduğu sekiz slaytlık bir Güneş Sistemi sunumunu gösteriyor: kapak slaytı tuvalde, solda sekiz küçük resim ve yapay zekâ paneli açık" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="Güneş Sistemi sunumunun işlenmiş sekiz slaytı yan yana: kapak, keşif zaman çizelgesi, dört temel rakam, gezegen çapı çubuk grafiği, kayaç gezegenler ile devler, Güneş'in %99,8'lik hero rakamı, dört dev gezegen ızgarası ve çıkarımlar"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="GenOffice Ayarlar, Integrations sayfası: genoffice skill'i Claude Code'a kurulmuş, Codex ve Cursor'ın yanında Install düğmeleri"></td>
</tr>
<tr>
<td><b>Ajanınıza tek bir istem</b> — "Güneş Sistemi hakkında sekiz slaytlık bir sunum oluştur." Ajan skill'i okur, bir stil sayfası, bir ana hat ve slayt başına bir sayfa spesifikasyonu yazar, iki fotoğrafı <code>genoffice image</code> ile üretir ve <code>genoffice create</code> dosyayı <code>.pptx</code> olarak birleştirip <code>slides render</code> bakmanız için slayt başına bir PNG vermeden önce <code>genoffice slides check</code> komutunun taşan ya da çakışan her şeyi reddetmesine izin verir.</td>
<td><b>Bir kez kurun, Ayarlar → Integrations üzerinden</b> — GenOffice bu bilgisayarda bulduğu kodlama ajanlarını listeler ve skill'i seçtiğiniz her birine yazar. Ya da skill'i zip olarak indirin ya da <code>npx skills add genspark-ai/genoffice</code> çalıştırın. Komutlar ve tam iş akışı <a href="#command-line-and-agent-skill">Komut satırı ve ajan skill'i</a> bölümünde.</td>
</tr>
</table>

### 10 · MCP — aynı araçlar, Model Context Protocol üzerinden

Her `genoffice` komutu aynı zamanda bir MCP aracıdır. Claude Code, Claude
Desktop, Cursor ve diğer tüm MCP istemcileri, kurulacak skill ve açık pencere
olmadan `genoffice mcp` komutunu kendileri başlatabilir ve 29 aracın yanı sıra
op başvurularını kaynak olarak alabilir. Uygulama içindeki ikinci bir HTTP
sunucusu, bir ajanın siz izlerken görünür bir düzenleyici sekmesinde Word
belgesi oluşturmasına izin verir.

<img src="../assets/readme/mcp-deck-motion.webp" alt="Claude Code'un genoffice MCP sunucusu üzerinden sekiz slaytlık bir yenilenebilir enerji yatırımcı brifingi oluşturmasının hızlandırılmış kaydı: rakamları ve fotoğrafları arıyor, her aday resmi media ile kontrol ediyor, deck_start stil sayfasını ve ana hattı yazıyor, deck_page her seferinde kontrol edilmiş bir sayfa ekliyor, deck_build .pptx dosyasını birleştiriyor ve slides_render her slaytın bir resmini döndürüyor; bitmiş sunum ardından GenOffice Slides'ta açılıyor" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides, Claude Code'un genoffice MCP sunucusu üzerinden oluşturduğu sekiz slaytlık Renewable Energy 2026 sunumunu gösteriyor: rüzgâr çiftliği fotoğraflı kapak slaytı tuvalde ve solda sekiz küçük resim"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="GenOffice Ayarlar, Integrations sayfası, MCP bölümü: Claude Code için tek satırlık claude mcp add komutu, Cursor, Claude Desktop ve diğer MCP istemcileri için JSON bloğu ve altında yerel HTTP sunucusu seçeneği"></td>
</tr>
<tr>
<td><b>Tek istem, 38 araç çağrısı, kabuk yok</b> — "2026'da yenilenebilir enerji hakkında, kapakta ve fotoğrafın yardımcı olduğu her yerde gerçek bir fotoğraf bulunan sekiz slaytlık bir yatırımcı brifingi oluştur." Ajan rakamları ve fotoğrafları <code>search</code> ile çeker, her aday resmin gerçek bir fotoğraf olup olmadığını <code>media</code> ile sorar, bir stil sayfası ve ana hat ile <code>deck_start</code>'ı, ardından slayt başına bir kez <code>deck_page</code>'i çağırır; her sayfa saklanmadan önce ana hatta ve palete göre kontrol edilir, <code>deck_build</code> <code>.pptx</code> dosyasını birleştirir, <code>slides_audit</code> taşma arar, <code>slides_render</code> modelin bakabileceği görsel içerik olarak slayt başına bir PNG verir ve <code>deck_replace</code> beğenmediği üç sayfayı düzeltir.</td>
<td><b>Bir kez bağlanın, Ayarlar → Integrations üzerinden</b> — Claude Code için <code>claude mcp add</code> satırını ya da JSON bloğunu Cursor, Claude Desktop veya başka bir MCP istemcisine kopyalayın. B seçeneği, görünür Word düzenleyicisi için yerel HTTP sunucusunu açar. İkisi de <a href="#mcp-server">MCP sunucusu</a> bölümünde açıklanmıştır.</td>
</tr>
</table>

## Neden GenOffice

- **Açık kaynak**, Apache-2.0, GitHub'da açıkta geliştiriliyor.
- **Sizin çalıştırdığınız.** macOS, Windows ve Linux için yerel uygulamalar;
  dosyalar diskinizde kalır ve her düzenleme, kaydetme ve dönüştürme
  makinenizde gerçekleşir.
- **Gerçek Office dosyaları.** Yerel `.docx`, `.xlsx` ve `.pptx`, bayt bayt
  korunur: dosyanın dokunmadığınız kısımları olduğu gibi kopyalanır.
- **Belgenin kendisini düzenleyen bir yapay zekâ.** Docs'ta izlenen
  değişiklikler, Sheets'te canlı formüller ve grafikler, tuvale çizilen
  slaytlar; her yapay zekâ turu geri alabileceğiniz bir anlık görüntü.
- **Sizin modeliniz, sizin anahtarınız.** Genspark ile oturum açın ya da
  Claude, OpenAI, Gemini, DeepSeek ve daha fazlası için bir anahtar getirin;
  yerel sunucular ve her OpenAI uyumlu uç nokta dahil.
- **Doğru dürüst PDF.** Sayfanın içinde metin düzenleyin ve PDF'yi cihaz
  üzerinde, taramalar için sistem OCR'ıyla Word, Excel ya da PowerPoint'e
  dönüştürün.
- **Markdown ve HTML de**, aynı yapay zekâ paneli ve Word'e yerel dışa aktarma
  ile.
- **Anahtar sözcüğü değil, yanıtı bulan arama.** Klasörlerinizdeki her belge
  üzerinde, cihazda dizinlenmiş tam metin araması; System One yargı modeli
  TypeSafe Jev ile isteğe bağlı yeniden sıralama.
- **Dosyalarınız arasında Ask AI.** Önemli dosyalara yıldız ekleyin ve doğal
  dille soru sorun; her yanıt geldiği dosyayı, tek tık uzağınızda kaynak
  gösterir.
- **Skills.** Hazır skill'ler ya da kendi `SKILL.md` dosyanız, herhangi bir
  yapay zekâ panelinin **+** menüsünden eklenir.
- **Betiklenebilir.** Bir `genoffice` komut satırı, bir ajan skill'i ve bir
  MCP sunucusu, her motoru Claude Code, Claude Desktop, Codex, Cursor ve diğer
  ajanların hizmetine, yine cihaz üzerinde sunar.
- **Ücretsiz**, bireyler ve ekipler için aynı şekilde.

## Yapay zekâ arka uçları

**Genspark ile oturum açın**; yapılandırılacak hiçbir şey yoktur: model
çağrıları Genspark proxy'si üzerinden yönlendirilir (Claude, GPT ve Gemini
aileleri) ve ajanlar web ve görsel arama, görsel üretimi ve görsel/ses/video
analizi kazanır.

**Ya da kendi anahtarınızı getirin.** Ayarlar → AI; Claude, OpenAI, Gemini,
DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter,
Requesty, Opper, Cheaper Inference, Atlas Cloud ve OpenCode Zen/Go'yu, ayrıca
yerel model sunucuları dahil herhangi bir OpenAI uyumlu uç nokta (temel URL +
anahtar) için özel bir yuvayı listeler. Arama ve medyanın **AI Media &
Search** altında yetenek başına kendi sağlayıcıları vardır: web araması için
Serper, Serply, Tavily, Parallel, Keenable, Exa ya da Firecrawl; görsel üretimi
ve görsel/video analizi için OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen,
MiniMax ya da herhangi bir OpenAI uyumlu görsel uç noktası; görsel analizi için
ayrıca DeepSeek V4.1 Flash.

**TypeSafe Jev**, ana ekranın dosya aramasını yeniden sıralar. **AI Media &
Search → Local file search** altında Jev search reranking seçeneğini açın ve
bir uç nokta seçin: [OpenRouter](https://openrouter.ai/typesafe) (model
`typesafe/jev-1.13`) ya da TypeSafe'in kendi API'si. Anahtar yalnızca bu
cihazda saklanır. Varsayılan olarak kapalıdır; açıkken en iyi 20 yerel sonucun
alıntıları değerlendirilmek üzere gönderilir ve makineden başka hiçbir şey
çıkmaz.

**Parallel** hesap olmadan çalışır: ücretsiz Search MCP'si (hız sınırlı),
Genspark girişi ya da arama anahtarı yapılandırılmamışken varsayılan web
aramasıdır ve DuckDuckGo kazımasından önce çalışır. Bunun yerine Search API'yi
kullanmak için Web search altında Parallel'i seçin ve bir
[Parallel](https://platform.parallel.ai/) anahtarı girin. Keenable de
anahtarsızdır; anahtar yalnızca hız sınırlarını yükseltir.

Paketin tamamı açık, koyu ve sistem temalarıyla gelir. Temalar yalnızca
ekrandakini değiştirir: dışa aktarmalar, yazdırmalar ve kaydedilen dosyalar
her zaman belgenin kendi renklerini korur.

## Skills

Bir skill, yapay zekâya tek bir işi iyi yapmayı anlatan bir `SKILL.md`
dosyasıdır: haftalık rapor, gizlilik sözleşmesi, ders planı. GenOffice bunların
bir setiyle gelir ve kodlama ajanlarının kullandığı biçimde kendi skill'lerinizi
de kabul eder.

- **Ayarlar → Skills.** Önerilen bir skill'i **Install** edin ya da bir
  klasörü, bir zip'i veya tek bir `SKILL.md` dosyasını **Import** edin. İçe
  aktarmada hiçbir şey çalışmaz.
- **Bir sohbette.** Herhangi bir yapay zekâ panelinde oluşturucunun yanındaki
  **+** menüsünü açın, **Skills**'i seçin, birini işaretleyin, gönderin. Skill
  yalnızca o mesajla birlikte gider.

---

## Komut satırı ve ajan skill'i

Uygulamaların bir dosyaya yapabildiği her şeyi `genoffice` komut satırı bir
terminalden yapabilir: Word, Excel, PowerPoint, PDF, Markdown ve HTML'yi aynı
motorlar üzerinde, başsız olarak inceleyin, dönüştürün, oluşturun, okuyun ve
düzenleyin. GenOffice ile birlikte kurulur, kendi çalışma zamanına ihtiyaç
duymaz ve hiçbir belgeyi hiçbir yere göndermez. Birlikte gelen **ajan
skill'i** ile eşleştirildiğinde, bir kodlama ajanını Markdown yaklaşımları
yerine gerçek Office dosyaları üreten bir belge çalışanına dönüştürür.

**Uyumlu olduğu araçlar:** kutudan çıktığı hâliyle Claude Code, Codex, Cursor,
Gemini CLI, GitHub Copilot, OpenCode ve Windsurf; skill okuyan diğer tüm
ajanlar ve [MCP sunucusu](#mcp-server) üzerinden Claude Desktop ile her MCP
istemcisi.

### Skill'i kurun

| Nasıl                                  | Ne olur                                                                                                                                                                       |
| -------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Uygulamada **Ayarlar → Integrations**  | Bu bilgisayarda bulunan ajanları listeler; tek tık skill'i seçtiğiniz her birine yazar. Bir GenOffice sürümü daha yeni bir skill getirdiğinde bir **Update** düğmesi belirir. |
| Aynı sayfadaki **Download as zip**     | claude.ai, Claude masaüstü uygulamaları ve diğer asistanların yüklenen skill olarak kabul ettiği düzen.                                                                       |
| `npx skills add genspark-ai/genoffice` | Bu depodan skill uyumlu herhangi bir ajana kurar.                                                                                                                             |

Ardından yeni bir sohbet başlatın ve bir belge isteyin. Skill, ajana ne zaman
`genoffice`'e başvuracağını, bir dosyayı düzenlemeden önce nasıl okuyacağını
ve kendi işini nasıl kontrol edeceğini öğretir.

### Terminalden hızlı başlangıç

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx/md/html, odt/ods/odp → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice docs compare original.docx revised.docx --out redline.docx   # the differences as tracked Word revisions
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

Her komut tek satırlık bir özet ya da `--json` ile tek bir JSON nesnesi
yazdırır. Düzenlemeler atomiktir: reddedilen bir op dosyayı olduğu gibi bırakır
ve yol gösterici bir hatayla geri döner. `genoffice help` güncel komut
yüzeyini listeler; tam başvuru
[packages/cli/README.md](../../packages/cli/README.md) dosyasındadır.

### Ajanın gerçekte çalıştırdıkları

[Demo](#demo) bölümündeki Güneş Sistemi sunumu Claude Code'da tek bir istem
aldı. Arkasında ajan, skill'in aşamalı iş akışını izledi ve CLI, bir sonraki
başlamadan önce her aşamayı kontrol etti:

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

`genoffice` içinde hiçbir model çağrısı yapılmaz: düşünmeyi ajan yapar,
oluşturmayı ve kontrolü CLI yapar ve sonuç GenOffice'te ya da PowerPoint'te
sıradan bir `.pptx` olarak açılır.

### MCP sunucusu

Aynı komutlar, terminal çalıştıramayan ya da terminal vermeyi tercih
etmediğiniz asistanlar için [Model Context Protocol](https://modelcontextprotocol.io)
araçları olarak da sunulur. İki giriş yolu vardır; ikisi de **Ayarlar →
Integrations → MCP** altında kopyalamaya hazır parçacıklarla gösterilir:

| Yol                                | Nedir                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ---------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **A · `genoffice mcp`** (önerilen) | Asistanın kendisinin başlattığı bir stdio sunucusu; GenOffice'in açık olması gerekmez. Komut başına bir araç (`info`, `convert`, `create_docx`, `create_xlsx`, `create_pptx`, `create_pdf`, `docs_read` / `docs_apply` / `docs_check`, `sheet_*`, `slides_*`, `render`, `guide`, `search`, `image`, `media`, `open`) ve aşamalı sunum akışı `deck_start` → `deck_page` → `deck_build` → `deck_replace`. Op'lar, spesifikasyonlar ve Markdown satır içi aktarılır; böylece dosya sistemi olmayan bir istemci de çalışır. |
| **B · Yerel HTTP sunucusu**        | GenOffice uygulamasının içinde `http://127.0.0.1:3093/mcp` adresinde çalışır (Streamable HTTP, eski SSE ile). Araçları görünür bir Word düzenleyici sekmesini yönetir: `create_session`, `insert_content`, `replace_blocks`, `apply_ops`, `read_document`, `save_session`; belgenin şekillenmesini izlersiniz. Varsayılan olarak kapalıdır; aynı ayarlar bölmesinden açın.                                                                                                                                              |
| **C · `genoffice mcp --http`**     | Stdio araç seti, başka makinelerdeki istemciler için bir Streamable HTTP sunucusu olarak: bir konteyner, bir sandbox, ağınızdaki paylaşılan bir makine. Dosyalar çağrılarla birlikte yolculuk eder: `PUT /files/<name>` bir dosya yükler ve bir URL döndürür, her `file` parametresi bir http(s) URL'si alır ve dosya yazan bir araç onu bir indirme URL'si olarak, küçükse ayrıca baytları bir MCP kaynağı olarak geri verir. `--host 0.0.0.0` onu ağa açar, `--token` korur.                                          |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor, Claude Desktop or any other MCP client
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

```bash
# On the machine that has GenOffice (private network; add --token for a shared box)
genoffice mcp --http 3093 --host 0.0.0.0 --token "$GENOFFICE_MCP_TOKEN"

# From the client: upload, then use the URL wherever a tool takes a file
curl -T report.docx -H "Authorization: Bearer $GENOFFICE_MCP_TOKEN" http://server:3093/files/
#   → { "url": "http://server:3093/files/<id>/report.docx", ... }
#   docs_read({ "file": "http://server:3093/files/<id>/report.docx" })
#   docs_apply(...) → output_url, downloadable with curl -o
```

HTTP üzerinden her oturum özel bir çalışma klasörü alır, göreli yollar ve
sunum klasörleri bunun içinde çözümlenir, `open` sunulmaz ve
`GENOFFICE_ALLOWED_ROOTS` ayarlanmamışken araçlar sunucunun kendi dosya
deposunun dışına çıkamaz. `render`, PDF'ye `convert` ve `create_pdf` yine de
gizli bir GenOffice süreci başlatır; bu yüzden başsız bir ana makinede
uygulamanın kurulu olması ve bir sanal ekran (`xvfb-run`) gerekir.

Buradaki `genoffice`, uygulamanın içinde gelen CLI'dir (macOS'te
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`; ayarlar
bölmesi kurulumunuz için tam yolu yazdırır). Sunucu kendi iş akışı
talimatlarını taşır ve op başvurularını `genoffice://guide/*` kaynakları
olarak sunar; bu yüzden skill gerekmez. Skill ve MCP sunucusu bir arada
bulunabilir, asistan birini seçer. Bulut özellikleri (`search`, `image`,
`media`) yine GenOffice'te yapılandırılan sağlayıcı üzerinden gider; geri
kalan her şey yerel çalışır ve `GENOFFICE_ALLOWED_ROOTS` her aracı
listelediğiniz klasörlerle sınırlar.

[Demo](#demo) bölümündeki yenilenebilir enerji sunumu, Claude Code'da yalnızca
`genoffice` MCP sunucusu bağlıyken tek bir istemin protokol tarafından nasıl
göründüğüdür:

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

Otuz sekiz çağrı, yaklaşık on üç dakika ve asistan hiç kabuğa dokunmadı:
rakamlar, fotoğraflar, kılavuzlar, kontroller ve işlenmiş görüntülerin tamamı
MCP araç sonuçları olarak yolculuk etti. Makineden yalnızca `search` ve
`media`, GenOffice'te yapılandırılan sağlayıcıya gitti.

## İndir

| Platform                             | Gereksinimler                                         | İndirme                                                                                |
| ------------------------------------ | ----------------------------------------------------- | -------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [Son `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **macOS** — Intel (x64)              | macOS 11+                                             | [Son `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)           |
| **Windows** (x64, çoğu PC)           | Windows 10+, Intel/AMD                                | [Son `-x64.exe` yükleyici](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** Arm üzerinde (ARM64)     | Arm üzerinde Windows 11 (Snapdragon X ve benzerleri)  | [Son `-arm64.exe` yükleyici](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 veya daha yeni)     | [Son `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Son `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — diğer dağıtımlar         | x86_64, glibc 2.34+, FUSE 2                           | [Son `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)            |

Tüm derlemeler `main` dalından gelir; macOS ve Windows yükleyicileri imzalıdır.
Eski sürümler [Releases](https://github.com/genspark-ai/genoffice/releases) sayfasındadır.

<details>
<summary><b>Linux'a kurulum</b></summary>

Deb paketi apt ile kurulur — bağımlılıkları çeker ve GenOffice'i uygulama
menüsüne ekler:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Fedora / RHEL ailesi / openSUSE üzerinde bunun yerine rpm'i kurun:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage olduğu yerde çalışır: FUSE 2 çalışma zamanını kurun
(`sudo apt install libfuse2`; Ubuntu 24.04'te paket `libfuse2t64`), dosyayı
çalıştırılabilir yapın, ardından çalıştırın:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Nasıl çalışır

Yedi Electron uygulaması — Docs, Sheets, Slides, PDF, Markdown, HTML ve sekmeli
kabuk — saf TypeScript paketlerinden oluşan tek bir motor katmanını ve `.xlsx`
için bir Rust yardımcı sürecini paylaşır. Özgün dosya her zaman tek doğru
kaynaktır: düzenlemeler dar yamalar olarak uygulanır ve düzenleyicinin
dokunmadığı her şey gidiş dönüşten dokunulmadan çıkar.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

Paket paket tur (docx/pptx motorları, `pdf2docx`, `html2docx`, ajan çekirdeği
ve sağlayıcılar) [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages)
dosyasındadır.

## Geliştirme

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

Sheets uygulaması xlsx yardımcı süreci için ayrıca bir Rust araç zincirine
(PATH üzerinde `cargo`) ihtiyaç duyar; `npm run build -w @genoffice/sheets`
bunu otomatik derler. Her değişikliğin geçmesi gereken kontroller ve çekme
isteklerinin nasıl birleştirildiği için [CONTRIBUTING.md](../../CONTRIBUTING.md)
dosyasına bakın.

## Topluluk

GenOffice etkin geliştirme aşamasındadır ve geri bildiriminiz onu şekillendirir.

- **Hata bildirin ya da özellik isteyin:**
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **GenOffice grup sohbetine katılın:** ekiple ve diğer kullanıcılarla konuşmak
  için [GenTeam](https://genoffice.ai/join).
- **X'te [@merrickbuilds](https://x.com/merrickbuilds) hesabını takip edin:**
  sürüm notları, demolar ve sırada ne olduğu için.
- **Depoya yıldız verin:** GenOffice size yararlıysa, projeyi desteklemenin en
  iyi yolu budur.

## SSS

<details>
<summary><b>GenOffice ücretsiz mi?</b></summary>

Evet. GenOffice, Apache-2.0 lisansı altında ücretsiz ve açık kaynaklıdır —
deneme süresi yok, uygulamaların kendisi için ücretli katman yok.

</details>

<details>
<summary><b>GenOffice, Microsoft Word, Excel ve PowerPoint dosyalarını açabilir mi?</b></summary>

Evet. GenOffice yerel `.docx`, `.xlsx` ve `.pptx` dosyalarını açar ve kaydeder.
Kaydetme bayt bayt korur: dosyanın dokunmadığınız kısımları bayt bayt geri
yazılır; böylece belgeler Microsoft Office'te çalışmaya devam eder.

</details>

<details>
<summary><b>GenOffice çevrimdışı çalışır mı?</b></summary>

Belge düzenleme tamamen yereldir — dosyalar açılmak, düzenlenmek, kaydedilmek
ya da dönüştürülmek için makinenizden asla çıkmaz. Yapay zekâ özellikleri
(ajanlar, arama, görsel araçları) bir ağ bağlantısı ve ya Genspark oturumu ya
da kendi model API anahtarınızı gerektirir.

</details>

<details>
<summary><b>GenOffice PDF dosyalarını düzenleyebilir mi?</b></summary>

Evet — üstünü örten açıklama notları değil, sayfa içerik akışını özgün yazı
tiplerini koruyarak yeniden yazan gerçek PDF metin ve görsel düzenleme.

</details>

<details>
<summary><b>GenOffice PDF'yi Word, Excel ya da PowerPoint'e dönüştürebilir mi?</b></summary>

Evet — tamamen cihaz üzerinde: PDFium karakter düzeyinde çıkarma ve geometri
tabanlı düzen analizi; bulut hizmeti yok, yükleme yok. Taranmış sayfalar da
kapsanır: macOS ve Windows'ta sistem OCR'ı onları okur; böylece bir sayfa
görüntüsü yerine düzenlenebilir metne dönüşürler.

</details>

<details>
<summary><b>Kendi yapay zekâ modelimi ya da API anahtarımı kullanabilir miyim?</b></summary>

Evet. Anahtarsız Genspark oturumunun yanı sıra GenOffice; Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter,
Requesty, Opper, Cheaper Inference, Atlas Cloud ve OpenCode Zen/Go için kendi
anahtarınızı getirmeyi, ayrıca yerel model sunucuları dahil herhangi bir
OpenAI uyumlu uç noktayı destekler. Arama, görsel üretimi ve görsel/video
analizi, Ayarlar → AI Media & Search altında kendi anahtarlarını alır.

</details>

<details>
<summary><b>Yapay zekâ panelini kapatabilir miyim?</b></summary>

Evet. Ayarlar → AI altında bir **Show the AI sidebar** anahtarı vardır.
Kapalıyken her düzenleyici kenar çubuğunu gizler; yeni belgeler onu kendi
başına açmayı bırakır. Açık bir yapay zekâ eylemi — bir şerit düğmesi, bir menü
öğesi — paneli o belge için yine açar ve kapatıldığında gizli duruma döner.

</details>

<details>
<summary><b>GenOffice HTML'yi Word'e dönüştürebilir mi?</b></summary>

Evet — HTML uygulamasındaki Export as Word, tamamen cihaz üzerinde yerel,
düzenlenebilir bir `.docx` üretir. Sayfa yerleşik Chromium'da işlenir ve
gerçek Word yapılarına indirgenir: başlıklar, paragraflar, listeler, tablolar,
kartlar, KPI satırları, form alanları ve sayfa arka planları; yalnızca Word
karşılığı olmayan görseller (grafikler, simgeler, süslü kutular) resim olarak
gömülür.

</details>

<details>
<summary><b>GenOffice'i Claude Code, Codex, Cursor ya da bir betikten yönetebilir miyim?</b></summary>

Evet. GenOffice, aynı motorları başsız çalıştıran bir `genoffice` komut satırı
kurar: belgeleri bir terminalden ya da betikten inceleyin, dönüştürün,
oluşturun, okuyun ve düzenleyin; programlar için `--json` çıktısıyla. Birlikte
gelen ajan skill'i Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot,
OpenCode ve Windsurf'e onu kullanmayı öğretir; **Ayarlar → Integrations**
üzerinden kurun. Bkz. [Komut satırı ve ajan skill'i](#command-line-and-agent-skill).

</details>

<details>
<summary><b>GenOffice herhangi bir veri topluyor mu?</b></summary>

Resmî paketlenmiş derlemeler varsayılan olarak sınırlı kullanım analitiği
gönderir ve raporlamayı Ayarlar → General altında istediğiniz zaman
kapatabilirsiniz. Analitik asla belge içeriği, dosya adları, dosya yolları,
hesap kimliği ya da e-posta adresi göndermez. Olay ve veri açıklamalarının
tamamı için [GenOffice Gizlilik](../../PRIVACY.md) sayfasına bakın.

</details>

## Güvenlik

Süreç güvenliği duruşu (işleyici sandbox'ı, IPC doğrulaması, dış bağlantı
denetimi) ve yapay zekâ tarafından üretilen içerik için tehdit modelleri için
[SECURITY.md](../../SECURITY.md) dosyasına bakın.

## Teşekkürler

GenOffice, şu açık kaynak projeler olmadan mümkün olmazdı:

- [Electron](https://www.electronjs.org/) — her uygulamanın masaüstü çalışma
  zamanı.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — Sheets'in
  genişlettiği elektronik tablo arayüz çekirdeği.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause,
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer) üzerinden
  paketlenmiş) — gerçek PDF metin ve görsel düzenlemenin arkasındaki içerik
  akışı motoru.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) ve
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — PDF işleme ve belge
  birleştirme.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  Docs ve Markdown'daki blok düzenleyiciler.
- [CodeMirror](https://codemirror.net/) (MIT) — HTML'deki kaynak düzenleyici.
- [Konva](https://konvajs.org/) — Slides ve Sheets grafikleri için tuval
  işleme.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — karmaşık yazı
  sistemleri için metin şekillendirme ölçüleri.
- [calamine](https://github.com/tafia/calamine) ve
  [IronCalc](https://github.com/ironcalc/IronCalc) — Rust xlsx yardımcı
  sürecinin okuma ve hesaplama katmanları.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — gömülü PowerPoint
  yazı tipleri için TypeScript'e taşınmış MicroType Express çözücüsü.
- [React](https://react.dev/) (MIT) — her uygulamanın arayüz katmanı.
- [Mermaid](https://mermaid.js.org/) (MIT) ve [KaTeX](https://katex.org/)
  (MIT) — Markdown ve Docs'ta diyagramlar ve matematik.
- [opentype.js](https://opentype.js.org/) (MIT) — ölçüler ve glif araması için
  yazı tipi ayrıştırma.
- [JSZip](https://stuk.github.io/jszip/) (MIT) ve
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — OOXML kapsayıcı ve XML katmanları.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — şeritlerdeki simge seti.
- [electron-updater](https://www.electron.build/) (MIT) — uygulama içi
  güncellemeler.
- Liberation, Carlito, Caladea ve Noto CJK yazı tipleri (OFL/Apache-2.0) —
  paketlenmiş belge yazı tipleri.

`npm run notices`, paketlenmiş üçüncü taraf lisans özetini yeniden üretir
(`tools/gen-third-party-notices.mjs`); tüm çalışma zamanı bağımlılıkları
MIT/Apache-2.0/BSD-3-Clause/OFL lisanslıdır.

## Lisans

GenOffice, tek bir istisnayla [Apache License 2.0](../../LICENSE) altında
lisanslanmıştır: `ee/` dizini gelecekteki kurumsal modüller için ayrılmıştır ve
[GenOffice Enterprise License](../../ee/LICENSE) kapsamındadır.

GenOffice ve Genspark adları ile logoları Mainfunc, Inc.'in ticari
markalarıdır. Apache-2.0 lisansı bunları kullanma izni vermez (bkz. 6. bölüm);
çatallar kendi markalarını kullanmalıdır.
