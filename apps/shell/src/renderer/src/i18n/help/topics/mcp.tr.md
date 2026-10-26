# Bir kodlama aracısını bağlama

GenOffice, Model Context Protocol'ü konuşur; böylece bir kodlama aracısı belgelerinizi uygulamanın kullandığı motorlarla okuyabilir, yazabilir ve işleyebilir. Aracı bir dosya biçimini tahmin etmez: yazılı işlem şemalarını, yürütücünün doğrulama yaptığı tanımların aynısından alır.

## Uygulamadan kaydetme

Bunu yapacağınız yer **Ayarlar ▸ Tümleştirmeler**'dir. Bölmenin iki yarısı vardır ve birini ya da ikisini birden kullanabilirsiniz.

**Beceri.** Bu makinede bulunan her kodlama aracısı için bir satır — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — her biri **Yükle**, **Güncelle** ve **Kaldır** ile, ayrıca **Başka bir klasöre yükle…**, **Beceriyi indir (zip)** ve **Yolu kopyala**. Yardımcınız listede yoksa GenOffice'i onun `SKILL.md` okuduğu bir klasöre yönlendirin ya da zip'i kaydedip yardımcının yüklemesine izin verin. Beceri ile MCP yan yana durabilir: yardımcı birini seçer ve ikisi de tamamen aynı şeyleri yapar.

**MCP.** İki yol sunulur: **Yardımcı tarafından başlatılan (önerilen)**, burada gösterilen yapılandırmayı istemcinize eklersiniz ve sunucuyu yardımcı kendisi başlatır; ve uygulamanın sizin için çalıştırdığı **Yerel HTTP sunucusu**. Her iki durumda da yardımcı sonunda GenOffice ile konuşur ve siz hiçbir zaman komut yazmazsınız.

## Komut satırından kaydetme

Aynı şey bir terminalden — bu gelişmiş yoldur ve aracı, bölmenin bulamayacağı bir yerde yaşıyorsa başvurulacak yoldur:

```sh
genoffice mcp install all
```

Bu makinedeki kodlama aracılarını bulur ve stdio sunucu girdisini her birinin kendi yapılandırmasına yazar; dosyanın geri kalanını bulduğu gibi bırakır.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Alışılmadık bir yere yüklediğiniz bir aracı `--dir <path>` alır; `--force` zaten var olan bir girdiyi yeniden yazar.

Sunucunun aldığı her şey tek ekrandadır — kaydetme, kaldırma ve listeleme biçimleri, HTTP üzerinden sunma ve iki şema bayrağı:

![genoffice mcp --help komutunun gerçek çıktısı: install, uninstall ve list biçimleri ile --http, --host, --token, --compact-schemas, --dir ve --force seçenekleri](img/mcp.png)

## Yardımcı olmadan çalıştırma

Başka bir makinedeki istemci için bunun yerine HTTP üzerinden sunun:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` dinlediği adresi değiştirir. Aynı belirteci istemciye de verin.

HTTP üzerinden dosyalar da taşınır: `PUT /files/<name>` bir dosya yükler, her araç bir yol yerine `http(s)` URL'si kabul eder ve çıktılar indirme URL'si olarak — yeterince küçük olduklarında gömülü kaynak olarak da — geri döner. İşlemler, spesifikasyonlar ve Markdown her iki durumda da satır içi aktarılır.

## Aracının eline geçenler

Her komut bir araçtır. İlginç olanlar:

- **`docs`, `sheet`, `slides`** — bir dosyayı uygulamanın kendi yazma yolu üzerinden, her seferinde bir **işlem** olarak okuyun ve düzenleyin. Yeni bir sunu `deck_start`, `deck_page`, `deck_build` şeklinde ilerler.
- **`render`** — uygulamanın işleyicisi tarafından yerleştirilmiş, sayfa başına bir PNG; böylece aracı bir slaydı tahmin etmek yerine ona bakabilir.
- **`pdf`** — bir PDF'nin metin katmanı, sayfa sayfa, uygulama süreci olmadan.
- **`info`** — meta veriler ve yapı özeti; tanımadığınız bir dosyada genellikle en ucuz ilk çağrıdır.
- **`search`, `image`, `media`** — uygulamada yapılandırılmış sağlayıcılar; böylece aracının kendi anahtarlarına ihtiyacı olmaz.
- **`merge`** — bir `{{key}}` şablonunu doldurun.

## Şemalar ve daha küçük bir bütçe

`apply` ve `create`, `ops`, `cells` ve `data` parametrelerini `genoffice guide <domain> --json` komutundan üretilen, işlem başına yazılı şemayla duyurur. Bu kesindir ama küçük değildir. Bağlam penceresi dar olan bir istemci bunun yerine düz diziler isteyebilir:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Beceri neden var

İşlem sözlüğünü bilmeyen bir aracı tahmin yürütür. Beceri, başvuru belgesini ve tasarım kılavuzlarını — `genoffice guide` komutunun yazdırdığı malzemenin aynısını — taşır; böylece yardımcı, spesifikasyonunu gerçekten okuduğu işlemler yazar. Üstteki bölmeden ya da bir terminalden `genoffice skill` ile yükleyin.

## Uygulamada ulaşabildikleri

Sunucu diskteki dosyalarla sınırlı değildir. GenOffice çalışırken aracı pencere üzerinden de çalışabilir:

- **`open_in_genoffice`** bir dosyayı sekmede açar ve odaklar.
- **`open_documents`** açık olan her belgeyi listeler — kimlik, tür, yol ve kaydedilmemiş değişiklik olup olmadığı — sonra birinin canlı içeriğini okur, yerinde kaydeder ya da siz atmasını istemedikçe önce kaydederek kapatır.
- **İçerik araçları** bu kimliği (veya yolu) `document` bağımsız değişkeni olarak alır; böylece bir düzenleme zaten açık olan sekmeye iner ve pencere onu göstermek üzere geçiş yapar.

İki şey erişim dışında kalır: AI paneli yoktur ve uygulama içi güncelleme iletişim kutusu uygulanmaz.

## Yerel HTTP sunucusu paneli

**Yerel HTTP sunucusu** altında uygulama, sunucuyu yardımcıya bırakmak yerine kendisi çalıştırabilir: bir etkinleştirme anahtarı, bir **Bağlantı noktası** alanı, bir **Çalışıyor / Çalışmıyor** göstergesi, **Arka planda oluşturma** (arayüzü açmadan belgeleri doğrudan bir yola yazar) ve istemcinize kopyalayacağınız **İstemci yapılandırma örneği**. **Gelişmiş**'i açmak iki bağlantı URL'sini — Streamable HTTP ve eski SSE — bir **Durum denetimi** URL'sini ve sunucu ile araç etkinliğini oradan **Aç**abileceğiniz, **Yenile**yebileceğiniz veya **Temizle**yebileceğiniz yerel bir dosyaya kaydeden bir **Günlük kaydı** anahtarını ekler. Yalnızca localhost üzerinde dinler.
