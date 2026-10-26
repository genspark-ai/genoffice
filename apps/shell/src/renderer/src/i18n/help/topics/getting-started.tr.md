# Hızlı başlangıç: arayüz ve temeller

GenOffice tamamen makinenizde çalışan bir ofis paketidir: tek pencere, tek sekme sırası, altı düzenleyici — Docs (kelime işlem), Sheets (elektronik tablo), Slides (sunu), PDF, Markdown ve HTML. Dosyalar gerçek .docx / .xlsx / .pptx / .pdf'dir; Word, Excel ve PowerPoint ile tamamen değiştirilebilir. Ağ gerekmez.

## Arayüze genel bakış

![Giriş ekranı](img/home-screen.png)

Pencere üç parçadan oluşur:

- **Sekme şeridi (üstte)**: açık her dosya bir sekmedir. En soldaki Giriş sekmesi her zaman vardır ve kapatılamaz; diğer sekmeler belgelerinizdir. Dosyasını satır içinde yeniden adlandırmak için bir sekmeye çift tıklayın.
- **İçerik alanı**: etkin sekmeye ait düzenleyici (veya Giriş).
- **Menü çubuğu**: macOS'ta sistem menü çubuğunda, Windows/Linux'ta pencerenin üstünde. Dosya/Düzenle/Görünüm menüleri etkin düzenleyiciye göre değişir.

## Belge oluşturma

Şunlardan herhangi biri:

- **Giriş** ekranının [Hızlı başlangıç](help://getting-started) bölümündeki bir hızlı oluşturma kartına tıklayın (AI Docs, AI Sheets, AI Slides, ...).
- Menü **Dosya ▸ Yeni**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML veya PDF.
- Bir dosyayı pencereye sürükleyin ya da dosya yöneticinizde çift tıklayın (GenOffice varsayılan uygulamaysa).

Yeni bir belge adsız açılır; diskteki dosya ancak ilk kayıtta oluşturulur.

## Dosya açma

- Menü **Dosya ▸ Aç** (⌘O/ctrl+O) sistem seçicisini açar: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Giriş ekranının **Son Kullanılanlar** listesinde herhangi bir öğeye tıklayın.
- Terminalden `genoffice <dosya>` da dosya açar.

## Kayıt modeli

- **Elle kaydetme**: ⌘S/ctrl+S veya Dosya ▸ Kaydet / Farklı Kaydet. İlk kayıt konum ve ad sorar.
- **Otomatik kaydetme** yalnızca dosyayı en az bir kez elle kaydettikten sonra açılır — yalnızca okuduğunuz bir PDF asla sessizce yeniden yazılmaz. Otomatik kaydetme içerik değişikliğinden kısa süre sonra tetiklenir.
- Kaydedilmemiş değişiklikleri olan bir sekmeyi kapatmak önce Kaydet / Vazgeç / İptal sorar.
- Her yazma atomik olarak gerçekleşir (geçici dosya + yeniden adlandırma), böylece elektrik kesintisi yarım dosya bırakamaz.

## Sık kullanılan kısayollar

| Eylem          | macOS | Windows / Linux |
| -------------- | ----- | --------------- |
| Yeni belge     | ⌘N    | ctrl+N          |
| Aç             | ⌘O    | ctrl+O          |
| Kaydet         | ⌘S    | ctrl+S          |
| Sekmeyi kapat  | ⌘W    | ctrl+W          |
| Bu kılavuzu aç | F1    | F1              |
| Şeridi katla   | ⌥⌘R   | Ctrl+F1         |

**Şeridi katlama** her düzenleyicide çalışır. Sekme sırası kalır, altındaki komut bandı gizlenir; seçili sekme aynı zamanda daraltma denetimidir, bu yüzden şerit katlıyken hiçbir sekme seçili değildir ve herhangi bir sekmeye basmak bandı geri getirir. Bir sekmeye çift tıklamak da aynısını yapar. Bıraktığınız durum düzenleyici başına hatırlanır. Docs bir ara yoğunluk ekler: **⌥⌘K** / Ctrl+Alt+K bandı yalnızca simgelere sıkıştırır ve aynı şekilde hatırlanır.

Her düzenleyicinin içindeki kısayollar (biçim boyacısı, bul ve değiştir, tablo işlemleri, ...) kendi bölümlerindedir; Docs ayrıca kendi tuş bileşimlerini eksiksiz kapsayan aranabilir bir klavye kısayolları iletişim kutusu (**⌘/**) sunar.

## Option+Command tuş bileşimleri

Option+Command, Word'ün yapısal atlamalar için ayırdığı katmandır ve GenOffice onu aynı şekilde doldurur. Çoğunu Docs alır; Sheets, Excel ile eşitlik için kendine iki tane alır; bir bileşim her yerde çalışır.

**Docs**

| Bileşim (macOS) | Ne yapar                 | Windows / Linux    |
| --------------- | ------------------------ | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3 | Başlık 1 / 2 / 3         | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0             | Normal (gövde metni)     | Ctrl+Alt+0         |
| ⌥⌘M             | Paragraf iletişim kutusu | Ctrl+Alt+M         |
| ⌥⌘A             | Yeni açıklama            | Ctrl+Alt+A         |
| ⌥⌘F             | Dipnot ekle              | Ctrl+Alt+F         |
| ⌥⌘E             | Sonnot ekle              | Ctrl+Alt+D         |
| ⌥⌘G             | Git                      | Ctrl+G             |

Bunlardan ikisi Windows'ta yer değiştirir; Word'ün onları ayırmasıyla aynı nedenle. **⌥⌘D macOS'a aittir** — Dock'u gösterip gizler — bu yüzden sonnot Mac'te ⌥⌘E, diğer her yerde Ctrl+Alt+D'dir. **Git** ise Alt'ı bırakır: Ctrl+G, oysa Mac bileşimi onu taşır.

**Sheets**, ızgara odaktayken

| Bileşim (macOS) | Ne yapar        | Windows / Linux |
| --------------- | --------------- | --------------- |
| ⌥⌘0             | Dış kenarlıklar | Ctrl+Shift+7    |
| ⌥⌘−             | Kenarlık yok    | Ctrl+Shift+−    |

Windows, Mac çiftinin yeniden yazımı değildir. Excel for Mac, Sheets'e **ikisini de** verir — ⌘⇧7 ve ⌥⌘0 aynı dış kenarlık için iki tuştur — bu yüzden Windows'ta komut zaten sahip olduğu Ctrl+Shift yuvasını korur ve Option katmanı yalnızca yoktur.

**⌥⌘0'ın Docs'ta Normal, Sheets'te Dış kenarlıklar anlamına geldiğini** unutmayın. Aynı düzenleyicide asla birlikte görünmezler, bu yüzden kullanımda çakışma olmaz; ancak ⌥⌘0 ayrılmıştır ve genel bir bileşim olarak kullanılamaz.

**Her düzenleyici**: **⌥⌘R / Ctrl+F1** şeridi katlar, yukarıda anlatıldığı gibi.

Böylece ⌥⌘D, gelecekteki bir komut isterse GenOffice'in macOS'ta kullanması için boş kalır.

## Sırada ne var

- Dosyalar nerede yaşar: [Giriş ekranı](help://home-screen).
- Çok sayıda açık dosyayı yönetme: [Sekmeler ve pencere yönetimi](help://tabs-and-windows).
- İşi AI'ya yaptırma: [AI asistan paneli](help://ai-panel).
- Dil, tema, varsayılan uygulamalar: [Ayarlar, dil, tema ve MCP entegrasyonları](help://settings-integrations).
