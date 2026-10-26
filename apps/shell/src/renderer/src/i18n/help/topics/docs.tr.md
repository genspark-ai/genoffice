# Docs: kelime işlem

Docs, Word benzeri kelime işlemcidir: gerçek .docx okur ve yazar, gerçek WYSIWYG sayfalandırma yapar.

## Şerit

Sekmeler: **Giriş / Ekle / Çiz / Düzen / Tasarım / Başvurular / Gözden Geçir / Görünüm**, artı seçili nesneye göre bağlamsal sekmeler — tablo için tablo tasarımı ve düzeni, resim için resim biçimi, şekil için şekil biçimi, üstbilgi/altbilgi düzenlerken üstbilgi ve altbilgi. Docs'ta şeridin kendine özgü bir yoğunluğu vardır: **sıkışık** (⌥⌘K / Ctrl+Alt+K) simgeleri bırakıp etiketleri kaldırır ve seçim hatırlanır.

- **Giriş**: pano; yazı tipi (CJK boyutları ve vurgu işaretleri dahil) ile **Tüm Biçimlendirmeyi Temizle** ve bir **Biçimlendirme işaretlerini göster/gizle** düğmesi; paragraf (hizalama/girinti/aralık/listeler) artı kendi liste stillerinizi belgeye kaydeden **Yeni Madde İşareti / Numara Biçimi / Çok Düzeyli Liste Tanımla**; stiller (Başlık 1-6/Normal/Alıntı, değiştirilebilir) ve tam liste için bir **Stiller Bölmesi**. Bir stil Doğu Asya ve Latin yazı tiplerini ayrı ayrı belirleyebilir (örneğin bir CJK yüzü artı Times New Roman); çift .docx gidiş dönüşünü atlatır.
- **Ekle**: sayfa ve bölüm sonları; tablolar (satır × sütun ızgarası veya tam boyut için **Tablo Ekle…**); resimler, şekiller, metin kutuları; hazır galeriden **Kapak Sayfası** ve **Boş Sayfa**; **Grafik**; **Başlangıcı Büyüt**; **WordArt**; alanlar (tarih, saat, sayfa, sayfa sayısı, dosya adı); köprüler, **Yer İşareti** ve çapraz başvurular; açıklamalar; üstbilgi/altbilgi ve sayfa numaraları; simgeler ve denklemler.
  - **Grafik**, kendi verisi olan gerçek bir grafik nesnesi ekler — çubuk, çizgi veya pasta — resim değil. Word'ün _Veriyi Düzenle_ komutu arkasındaki sayıları açar.
- **Çiz**: sayfa üzerine mürekkep, bir **Çizim Araçları** grubunda — metin düzenlemeye dönmek için **Seç**, ardından **Kalem**, **Vurgulayıcı** ve **Silgi** (bir tıklama ya da kaydırma tüm vuruşu kaldırır). Yanlarında **Kalem Stili** / **Vurgulayıcı Stili**, renk örneklerini ve bir sıra kalınlığı tutan tek bir denetimdir; etiketi etkin araca göre değişir. Mürekkep belgede metnin üstünde yüzen bir ek açıklama olarak saklanır, bu yüzden kaydedip yeniden açınca korunur; sonraki gruptaki **Tümünü Temizle** hepsini kaldırır.
- **Düzen**: kenar boşlukları, yön ve kâğıt boyutu, sütunlar, paragraf girintileri ve aralığı.
- **Tasarım**: temalar, renk kümeleri, filigran, sayfa kenarlıkları.
- **Başvurular**: içindekiler tablosu (güncellenebilir), dipnotlar/sonnotlar, resim yazıları, çapraz başvurular.

  ![Başvurular sekmesi](img/docs-references.png)

- **Gözden Geçir**: **Düzenleyici** tüm belgeyi yazım, dil bilgisi ve noktalama açısından denetler; **Çevir**; yazım denetimi; açıklamalar (**AI ile Açıklamaları Çöz** açık olanları sırayla işler); Tümü/Basit işaretleme görünümleriyle değişiklikleri izleme, kabul/ret ve **AI Düzeltme Özeti**; sözcük sayısı; başka bir dosyayla **Karşılaştır**; **Belgeyi Koru**.
- **Görünüm**: dosyaya bakmanın beş yolu — **Sayfa Düzeni**, **Web Düzeni**, **Anahat**, **Okuma Modu** ve **Sayfa Önizleme**; uzaklaştır/yakınlaştır/%100/sayfa genişliği/tek sayfa; **AI Paneli**; **Koyu Mod**; cetvel, kılavuz çizgileri ve gezinti bölmesi; **Yeni Sekme**, **Böl** ve **Sekme Değiştir**; aranabilir **klavye kısayolları iletişim kutusu**.
  - **Koyu Mod** sayfayı ve çevresindeki tuvali karartır, şeridi asla — Word'ün koyu düzenleme yüzeyi ile koyu pencere arasındaki ayrımı. Seçim hatırlanır ve her durumda uygulama temasına üstün gelir.
  - **Böl**, altta bağımsız kayan ve ilkini yansıtan ikinci bir bölme açar; kenarındaki × ile kapatın.

## Gezinti bölmesi

**Görünüm ▸ Gezinti Bölmesi**, belgenin başlık anahattını, tüm belge
üzerinde bir arama kutusunu ve her sayfa için bir küçük resmi tutan bir yan bölme açar.
Açık olup olmadığı başlatmalar arasında hatırlanır, böylece anahatla gezdiğiniz bir belge
gezilebilir kalır.

**Anahat**, başlık ağacıdır. Yalnızca gezinmek yerine katlamak ve yeniden yapılandırmak için içindeki bir başlığa sağ tıklayın:

- Bir başlıktaki **Daralt / Genişlet** o başlığın tüm alt ağacını katlar — bölüm
  görünmez olur, metni belgede kalır.
- **Tümünü Daralt / Tümünü Genişlet** hepsini bir kerede katlar veya açar. Uzun bir raporda
  bu, okunabilir bir anahat ile bir metin duvarı arasındaki farktır.
- **Başlık Düzeylerini Göster** ağacı önemsediğiniz derinliklere süzer, böylece
  _Başlık 1'i Göster_ size gerçekten göz gezdirebileceğiniz bir içindekiler bırakır.
- **Yükselt / Alçalt** başlığın düzeyini ve onunla birlikte altındaki her başlığın
  devraldığı düzeyi değiştirir — bir bölümün alt bölüm olma yolu.
- **Öncesine / Sonrasına Yeni Başlık**, bölmeden ayrılmadan imlecin olduğu yere bir tane ekler.
- **Sil** başlığı _ve altındaki her şeyi_ kaldırır; dikkatli olunması gereken budur:
  bir alt ağaç silme işlemidir, satır silme değil.
- **Başlığı ve İçeriği Seç**, başlıktan alt ağacının sonuna kadar seçer,
  bütün bir bölüm düzenlemesine hazır.

## Sağ tıklama menüsü

Gövdede herhangi bir yere sağ tıklayın — menü tıkladığınız şeye uyar. Ana gruplar:

- **Pano**: kes / kopyala / yapıştır / **düz metin olarak yapıştır**.
- **Yazı tipi, paragraf**: şeride gitmeden aile ve boyut değiştirme, kalın/italik/altı çizili, hizalama/girinti/aralık.
- **Eş anlamlılar**: seçili sözcüğün eş anlamlılarını listeler; değiştirmek için birine tıklayın.
- **Çevir** (AI): seçimi AI paneli üzerinden hedef dile (İngilizce, Basitleştirilmiş Çince, Japonca, Korece, Fransızca, Almanca, İspanyolca, ...) çevirir.
- **Yeni açıklama**: seçime bir açıklama iliştirir.
- **Yazım** (yanlış yazılmış bir sözcükte): önerilen değiştirmeler, tümünü yoksay, sözlüğe ekle, yazım denetimi dilini ayarla.
- **Köprü**: aç / düzenle / bağlantıyı kopyala / köprüyü kaldır.
- **Resim**: görseli görüntüle, görseli farklı kaydet, **metin kaydırma** (satır içi / kare sol ve sağ / üst ve alt / metnin arkasında / metnin önünde), sıralama düzeni.
- **Alanlar** (İçindekiler, sayfa numaraları): alanı güncelle / alan kodlarını aç-kapat / alanı düzenle.
- **Liste numaralandırma** (bir listenin içinde): numaralandırmayı yeniden başlat / numaralandırmaya devam et / liste düzeyini değiştir / numaralandırma değerini ayarla.
- **Tablo** (imleç bir tablonun içinde): satır/sütun ekleme, hücreleri birleştir / böl, tabloyu böl, otomatik sığdır, hücre hizalama, satırları/sütunları dağıt, tablo özellikleri, silme menüsü, seç.

## Düzenleme

- Bul ve değiştir (ctrl+F / ctrl+H): büyük/küçük harf duyarlı, tam sözcük, düzenli ifade.
- Biçim boyacısı; derin geri al/yinele; yapıştırma seçenekleri.
- Tablolar: hücreleri birleştirme/bölme, satır/sütun işlemleri, kenarlıklar ve gölgelendirme, sıralama, formüller. Satırlar gerektiğinde sayfalar arasında bölünür ve bir hücre içinde Ctrl+Enter sayfayı tam orada çevirir — yazdırma ve PDF bunu izler.
- Resimler: metin kaydırma, kırpma, sıkıştırma; çizim tuvali.

## CJK tipografisi

- Noktalama sıkıştırma ve kinsoku satır sonu kuralları Word ile eşleşir; tam/yarım genişlik dönüştürme.
- Yazı tipi adayları yaygın Windows ve macOS CJK aile adlarını kapsar.

## AI

- Şerit AI düğmesi ve yan panel: yeniden yazma, genişletme, çeviri, özetleme, tablo ekleme ön ayarları ve serbest biçimli yönergeler.
- Her AI turu önce anlık görüntü alır; sürüm listesinden geri dönün, geri dönüşün kendisi de geri alınabilir.

## Kaydetme ve dışa aktarma

- .docx'i yalnızca değişen paragrafları yeniden yazarak kaydeder — dokunulmamış içerik bayt bayt aynı kalır.
- PDF (sayfalandırıldığı gibi) ve sayfa başına görsel dışa aktarır.

## Yazdırma

Sistem iletişim kutusu üzerinden ctrl+P, WYSIWYG sayfalar.
