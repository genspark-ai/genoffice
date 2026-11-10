# Sheets: elektronik tablolar

Sheets, Excel benzeri düzenleyicidir; hesaplama ayrı bir Rust motoru sürecinde çalışır (oradaki bir çökme uygulamayı asla kapatmaz). Gerçek .xlsx açar ve kaydeder; .csv ve .tsv tablo olarak açılır.

## Arayüz

- **Şerit**: sekiz sekme, aşağıda tek tek listelenmiştir.
- **Formül çubuğu**: etkin hücrenin formülünü gösterir ve düzenler; yaygın işlevler desteklenir.
- **Sayfa sekmeleri** (altta): sayfa ekleme / yeniden adlandırma / silme / taşıma.
- **Hücre düzenleme**: çift tıklayın veya doğrudan yazın; Enter onaylayıp aşağı iner, Tab sağa gider, Escape iptal eder (Excel alışkanlıkları).
- **Kısayollar**: Excel ailesiyle uyumlu (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Şerit sekmeleri

- **Giriş**: yazı tipi, dolgu, kenarlıklar, sayı biçimleri (para birimi/yüzde/binlik, ondalık artır/azalt), hizalama, birleştirme, satır/sütun ekleme ve boyutu, koşullu biçimlendirme, tablo olarak biçimlendir, hücre stilleri, pano ve biçim boyacısı, sırala ve filtrele.
- **Ekle**: resimler, şekiller, metin kutuları, bağlantılar, yorumlar, onay kutusu, üstbilgi/altbilgi, simgeler, denklemler; **dokuz grafik türü** (sütun, çubuk, çizgi, alan, pasta, dağılım, radar, halka ve sütun+çizgi birleşimi) ve seçim için türleri sıralayıp önizleyen **Önerilen Grafikler**; bulunduğunuz hücreden **PivotChart**; **Mini Grafikler** (çizgi, sütun, kazanç/kayıp); bir PivotTable'ı filtrelemek için **Dilimleyici** ve **Zaman Çizelgesi**.
  - Grafikler, PivotChart'lar ve mini grafikler kaydedilen çalışma kitabında gerçek nesnelerdir.
  - **Dilimleyici** ve **Zaman Çizelgesi** oturum denetimleridir: yaptıkları filtreleme kaydedilir ve Excel aynı filtrelenmiş PivotTable'ı gösterir, ancak dilimleyici düğmesinin kendisi dosyanın parçası değildir.
- **Sayfa Düzeni**: tema renkleri ve yazı tipleri, kılavuz çizgilerini/başlıkları yazdırma anahtarları, sayfa sonu önizlemesi.
- **Formüller**: Otomatik Toplam ve işlev ekleme, ad tanımlama (seçimden de), etkileyenleri/etkilenenleri izleme, Gözcü Penceresi, sayfayı/çalışma kitabını yeniden hesaplama.
- **Veri**: sırala ve filtrele (gelişmiş filtre, filtreyi temizle dahil), metni sütunlara dönüştür, çalışma kitaplarını birleştir, tümünü yenile.
- **Gözden Geçir**: yorumlara göz atma (göster, önceki/sonraki), çeviri ve bir **Koruma** grubu — **Sayfayı Koru**, **Çalışma Kitabını Koru** ve **Düzenlenebilir Aralıklar**.
  - Koruma .xlsx dosyasına yazılır ve bu uygulamanın uyguladığı korumada parola yoktur; bu yüzden aynı düğme **Korumayı Kaldır …** olur ve geri alır. Başka bir programdan gelen ve parola _taşıyan_ koruma buradan kaldırılamaz.
  - İki koruma düğmesi hemen uygulanır — iptal edilecek bir iletişim kutusu yoktur, yalnızca kaydederken yazılacağını belirten bir durum çubuğu notu vardır.
  - **Düzenlenebilir Aralıklar**, sayfanın geri kalanı kilitliyken düzenlenebilir kalacak hücreleri işaretler.
- **Görünüm**: kılavuz çizgileri, formül çubuğu, başlıklar ve çapraz vurgu anahtarları; yakınlaştırma; **Normal** ve **Sayfa Sonu Önizleme**. Kılavuz çizgileri ve başlıklar sayfayla kaydedilir; çapraz vurgu sizin kendi tercihinizdir.
- **Grafik Tasarımı**: bir grafik seçiliyken görünür — grafik türü, stiller ve renkler, veri aralığını düzenleme.

Veri sekmesi, düğme düğme (resimde soldan sağa):

![Veri sekmesi](img/sheets-data.png)

- **Pivot tablo**: geçerli aralıktan bir pivot oluşturun; toplamak için alanları sürükleyin.
- **Yenile**: geçerli pivotun verilerini yeniden hesaplayın.
- **Metin/CSV'den**: bir .csv/.txt dosyasını ayırıcıya göre bölerek yeni sayfa olarak içe aktarın.
- **Çalışma kitaplarını birleştir**: diğer .xlsx dosyalarındaki sayfaları bu dosyaya çekin.
- **Tümünü yenile**: her pivotu ve dış veri kümesini yeniden hesaplayın.
- **Sırala** (açılır liste): artan / azalan / özel sıralama (çok sütunlu kurallar).
- **Filtre**: başlık satırına ▼ açılır listeleri ekler; tutulacak değerleri işaretleyin.
- Yanındaki küçük üst üste düğmeler: **Temizle** (her satırı geri getirir), **Yeniden Uygula** (geçerli filtreyi yeniden çalıştırır), **Gelişmiş** (ölçüt aralığıyla filtreleme).
- **Metni sütunlara dönüştür** (açılır liste): bir sütunu ayırıcıya veya sabit genişliğe göre birden çok sütuna bölün.
- **Hızlı doldurma**: bir örnek verin, sütunun geri kalanı örneğe göre dolsun (ctrl+E).
- **Yinelenenleri kaldır**: seçili sütunlara göre yinelenen satırları atın.
- **Veri doğrulama** (açılır liste): seçim için giriş kuralları (açılır listeler, sayı aralıkları...).
- **Birleştir**: birden çok aralığı kategoriye göre tek yerde toplayın.
- **Durum çözümlemesi** (açılır liste): Hedef Ara — bir formül hücresi hedefe ulaşsın diye tek bir giriş hücresini çözün.
- **Gruplandır / Grubu Çöz** (açılır liste): daraltılıp genişletilebilen satır/sütun grupları.
- **Alt toplam**: kategori başına alt toplam satırları ekleyin.

Formüller sekmesi, düğme düğme:

![Formüller sekmesi](img/sheets-formulas.png)

- **İşlev ekle** (fx): bağımsız değişken sihirbazıyla işlev arayın.
- **Otomatik Toplam** (açılır liste): tek tıkla SUM, ayrıca ortalama/sayı/en büyük/en küçük.
- **Son kullanılan / Finansal / Mantıksal / Metin / Tarih ve saat / Arama ve başvuru / Matematik ve trigonometri / Diğer**: işlevlere kategoriye göre göz atın ve ekleyin.
- **Ad yöneticisi**: adlandırılmış aralıkları görüntüleyin, oluşturun ve silin.
- **Ad tanımla** (açılır liste): seçimi adlandırın; **Formülde kullan** var olan bir adı ekler; **Seçimden oluştur** aralıkları başlık satırı/sütunundan adlandırır.
- **Etkileyenleri izle / Etkilenenleri izle**: bir formülün verisinin nereden geldiğini ve nereyi beslediğini gösteren mavi oklar; **Okları kaldır** onları temizler.
- **Formülleri göster**: hücreler sonuç yerine formülün kendisini gösterir.
- **Hata denetimi**: formül hatalarını bulun ve açıklayın.
- **Gözcü penceresi**: önemsediğiniz hücreleri sabitleyin ve canlı değerlerini izleyin.
- **Hesaplama seçenekleri** (açılır liste): otomatik ya da elle yeniden hesaplama; elle modda **Şimdi hesapla / Sayfayı hesapla** bunu elle tetikler.

## Sayılar ve biçimlendirme

- Sayı biçimleri: genel, sayı, para birimi, yüzde, tarih/saat, kesir, bilimsel ve daha fazlası.
- Hizalama, kaydırma, birleştirilmiş hücreler, kenarlıklar ve dolgular.
- Satır yükseklikleri ve sütun genişlikleri sürükleyerek; otomatik sığdırmak için bir sınıra çift tıklayın.

## Veri

**Sırala ve filtrele** (örneğin bir sütuna göre azalan):

1. **O sütundaki herhangi bir hücreye** tıklayın (tüm sütunu seçmeye gerek yok).
2. Giriş sekmesi ▸ **Sırala ve Filtrele** ▸ **Azalan**; satırlar bütün olarak birlikte yeniden sıralanır (alan tek parça olarak sıralanır).
3. Özel kurallar için (birden çok sütun, renge göre): aynı yol, **Özel Sıralama**.
4. Filtre: başlık satırını seçin ve **Sırala ve Filtrele ▸ Filtre**'ye tıklayın — her başlıkta tutulacak değerleri işaretlediğiniz bir ▼ açılır listesi belirir; her şeyi geri getirmek için filtreyi temizleyin.

- Sıralama ve filtreleme.
- Dondurulmuş bölmeler.
- .csv / .tsv: doğrudan tablo olarak açılır (sekmeyle ayrılmış tsv tek parça olarak ayrıştırılır); kaydetme özgün biçimi geri yazar.

## Sağ tıklama menüleri

- **Kılavuzda**: düzenleyicinin (Univer'in) kendi menüsü — kes/kopyala/yapıştır, satır/sütun ekleme ve silme, gizleme, hücre birleştirme, bölmeleri dondurma ve diğer günlük öğeler.
- **Alttaki durum şeridinde**: durum çubuğunun hangi istatistikleri göstereceğini seçin (ortalama / sayı / toplam, ...); seçim kalıcıdır.
- **Alttaki bir sayfa sekmesinde**: sayfa ekleme / yeniden adlandırma / silme / renklendirme / gizleme (Univer'in sekme menüsü).
- Üst sekme şeridinin bağlam menüsü [Sekmeler ve pencere yönetimi](help://tabs-and-windows) bölümünde anlatılmıştır.

## AI

- Yan AI paneli: bir aralık seçin ve düz dille yönerge verin (yeniden biçimlendir, veri üret, formül yaz).
- 📎 düğmesiyle isteme dosya ekleyin ya da panele sürükleyin; soruyla birlikte giderler ve görseller küçük resim olarak geri gelir.
- Bir yanıt bir hücreye başvurabilir — başvuruya tıklayın, kılavuz oraya atlar.
- AI düzenlemeleri panelden geri alınabilir.

## Kaydetme ve dışa aktarma

- .xlsx kaydeder (formüller ve biçimler korunur); Farklı Kaydet; PDF'ye Dışa Aktar yazdırma sayfalandırmasını izler.
- Otomatik kaydet genel kurala uyar (ilk elle kaydetmeden sonra açık).

## Kararlılık

- Rust hesaplama yardımcı süreci arayüzden süreç olarak yalıtılmıştır: aşırı veri onu öldürürse bir ileti ve oturum kurtarma denemesi alırsınız — uygulama çökmesi değil.
