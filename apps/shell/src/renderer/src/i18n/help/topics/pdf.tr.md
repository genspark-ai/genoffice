# PDF: okuma, açıklama ekleme ve karartma

PDF düzenleyicisinin beş şerit sekmesi vardır: **Giriş / Açıklama / Düzenle / Sayfa / Görünüm**. Okur ve yazar: metin düzenlenebilir, içerik karartılabilir, imzalanabilir ve formlar doldurulabilir.

## Okuma ve gezinme

- Sol kenar çubuğu: **küçük resimler** (atlamak için tıklayın, görünen aralık vurgulanır) veya **ana hat** (varsa yer işaretleri).
- Yakınlaştırma: sağ alttaki oran denetimi; ctrl+tekerlek yakınlaştırmayı kademeli değiştirir.
- Döndürme: Sayfa menüsünden sayfa başına veya tüm sayfalar; döndürmeler kaydederken geri yazılır.
- Arama: ctrl+F ile tüm eşleşmeler vurgulanarak tam metin araması.
- Şifreli PDF'ler: bir parola istemi (kendi küçük penceresi) onları açar; parola yalnızca bu oturum için kullanılır.

## Metin seçimi ve işaretleme (Açıklama)

Herhangi bir paragrafta deneyin:

1. **Fareyi bir cümlenin üzerinden sürükleyin** — bıraktığınızda üstünde bir açıklama çubuğu belirir:

![Metin seçildikten sonra açıklama çubuğu](img/pdf-highlight.png)

2. **Vurgu** (sarı renk örneği bir renk paleti açar), **altı çizili** veya **üstü çizili** seçin; **AI'ya Sor** seçimi sorunuzla birlikte AI paneline gönderir.
3. Bir açıklamayı geri almak için aynı bölümü yeniden sürükleyerek seçin ve çubuktaki etkin düğmeye tıklayın (Word tarzı açma/kapama) ya da onu seçip Delete'e basın.

Başvuru:

- Metnin üzerinden sürükleyin, bir açılır çubuk belirir: **vurgu / altı çizili / üstü çizili / kopyala / AI'ya Sor**.
- Renkler paletten gelir; **zaten işaretli bir aralığa aynı işaretlemeyi uygulamak onu kaldırır** (Word tarzı açma/kapama).
- Dosyaya daha önce kaydedilmiş işaretlemeler seçilip silinebilir (⋯ menüsü veya Delete).
- **Not**: bir çizim aracı etkinken metin katmanı seçilemez — araç her yerleştirmeden sonra kendini devre dışı bırakır, böylece bir sonraki eylem için yeniden seçim modundasınız.

## Çizim araçları (Açıklama)

Altı araç: **mürekkep, dikdörtgen, elips, ok, not** ve Açıklama sekmesindeki **karartma kutusu**.

- Her araç bir açma/kapama düğmesidir: etkinleştirmek için tıklayın; **bir şekil yerleştiğinde kendini devre dışı bırakır** (devam etmek için araca yeniden tıklayın); etkin araca tıklamak da devre dışı bırakır.
- Mürekkep, çizgi kalınlığını izler; dikdörtgen/elips/ok sürüklenerek çizilir; renkler çizim paletinden gelir.
- Yerleştirilen şekiller seçilebilir, silinebilir, sürüklenebilir ve (dikdörtgen/elips) yeniden boyutlandırılabilir.
- **Karartma, tam iş akışı** (bir metin satırını gizleme):

  1. Açıklama sekmesi ▸ **Alanı karart**'a tıklayın (araç etkinleşir).
  2. **İçeriğin üzerine bir kutu sürükleyin** — taralı bir işaretle örtülür ve araç çubuğuna **işaretleri temizle / karartmayı uygula** düğmeleri eklenir:

  ![Bir karartma işaretlendikten sonra sayfa](img/pdf-redact.png)

  3. **Karartmayı uygula**'ya tıklayıp onaylayın — sonuç, örtülen metin ve görsellerin (örtülmek yerine) fiziksel olarak kaldırıldığı bir çalışma kopyasıdır ve geri alınamaz; özgün belge dokunulmadan kalır.

  Hata mı yaptınız? İşaretleri temizle düğmesi geçerli işaretleri siler, böylece yeniden çizebilirsiniz.

## Yapışkan notlar ve yorum dizileri

- **Not aracı** bir raptiye bırakır ve metin için kenar boşluğunda bir kart açar (yazar adı yapılandırılabilir); onaylamak onu standart bir PDF Metin açıklaması olarak kaydeder.
- Diziyi açmak için bir raptiyeye tıklayın: **yanıtla** (düz, WPS/Acrobat tarzı diziler), yorumunuzu **düzenle**, tek bir yorumu veya tüm diziyi **sil**.
- Devam eden düzenlemeler, kaydetme yeni metni dosyadaki aynı açıklamaya geri yazana kadar korunur ve yanıt zincirleri bozulmaz.

## PDF içeriğini düzenleme (Düzenle)

- **Metni düzenle**: metne tıklayarak blok blok düzenleyin (pdfium motoru; yazı tipleri elden geldiğince eşleştirilir).
- **Metin ekle**: seçtiğiniz yazı tipi/boyut/renkle aranabilir metin yerleştirin.
- **Görsel / damga ekle**.
- Formlar: AcroForm alanları doğrudan doldurulur; değerler kaydederken yazılır.

## İmzalar

- **Mürekkep imzası**: çizin; bir form imza alanına eklenebilir.
- **Görsel imza**: imza olarak bir resim yerleştirin.
- Kaydedilen imzalar yeniden kullanılabilir.
- **Sertifikayla imzala** (Açıklama sekmesi): sistem deposundaki bir sertifikayla veya bir .p12/.pfx dosyasıyla imzalayın; isteğe bağlı olarak imzayı sayfada gösterin ve isteğe bağlı olarak belgeyi sonraki değişikliklere karşı onaylayın.
- İmzalama **yeni bir dosya** yazar (artımlı güncelleme — önceki imzalar geçerli kalır) ve imzalı kopyayı onun yerine açar. Önce kaydedin: bir imza, dosyanın diske kaydedilmiş halini kapsar.
- Belgenin üstündeki bir şerit mevcut imzaları listeler; **İmzaları görüntüle** her birinin geçerliliğini, kapsamını ve sertifika zincirini gösterir. İptal (CRL/OCSP) denetlenmez.

## Sayfa işlemleri (Sayfa)

- **Döndür / sil / yeniden sırala**: yeniden sıralamak için küçük resimleri sürükleyin; silme onay ister.
- **Sayfaları içe aktar**: başka bir PDF'den sayfa çekin. **Boş sayfa ekle** boş bir sayfa ekler.
- **Sayfaları değiştir** bir aralığı başka yerden gelen sayfalarla değiştirir; **Sayfaları kırp** kırpar, tümüne uygulama seçeneğiyle.
- **Sayfa boyutu** her sayfayı tek bir kağıt boyutuna yeniden ölçekler; **Sırayı tersine çevir** belgeyi baştan sona ters çevirir.
- **Sayfaları ayıkla**: seçili sayfaları yeni bir PDF'ye dışa aktarın.
- **Böl**: iki biçimi var — aralıklara göre birden çok dosyaya ya da her sayfayı daha küçük sayfalardan oluşan bir ızgaraya keserek.
- **Birleştir**: iki biçimi var — başka PDF'leri sona ekleyin ya da birden çok sayfayı tek bir yaprakta birleştirin. Boyutlar herhangi bir şey okunmadan **önce** toplanır ve **toplam 1 GiB üzeri** okunabilir bir hatayla **reddedilir** (belleği sınırlı tutar).
- Sayfa düzeyindeki değişiklikler bir sonraki kaydetmede geri yazılır; Farklı Kaydet özgün dosyayı dokunulmadan bırakır.

## Dışa aktarma ve yazdırma

- Dosya menüsünde **Word… / PowerPoint… / Excel… olarak dışa aktar** veya şeritteki **PDF Dönüştürücü**'den aynı üçü — hepsi yerel, yükleme yok. .pptx sayfa başına bir slayt, .xlsx sayfa başına bir çalışma sayfası olarak çıkar. Her biri nereye kaydedileceğini sorar.
- **Yazdır**: geçerli sıra ve döndürmelerle sistem iletişim kutusu üzerinden; sayfa aralıkları desteklenir.

## Kaydetme

- Normal Kaydet/otomatik kaydet, açıklamaları ve düzenlemeleri dosyaya geri yazar (atomik).
- **Karartma, Uygula akışından geçer** ve bir kopya üretir; özgün dosya dokunulmadan kalır, böylece hassas içerik içinde kalmaz.
