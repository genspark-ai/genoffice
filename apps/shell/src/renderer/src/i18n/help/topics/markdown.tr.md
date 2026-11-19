# Markdown düzenleyicisi

Markdown düzenleyicisi .md / .markdown dosyalarını — ve doğrudan kaynak görünümünde .txt / .json dosyalarını — kaynak + işlenmiş önizleme deneyimiyle açar.

- **Açma**: Giriş'ten veya Dosya ▸ Aç'tan; CLI de çalışır.
- **Düzenleme**: düz metin düzenleme; GFM uzantıları (tablolar, görev listeleri, üstü çizili, otomatik bağlantılar) önizlemede işlenir.
- **Önizleme**: canlı; görseller gibi göreli varlıklar belgenin yanında çözümlenir.
- **Kaydetme**: bayt düzeyinde sadık — BOM, CRLF ve sondaki satır sonunun varlığı korunur; değiştirilmemiş bir kaydetme dosyayı yeniden yazmaz.
- **Bul ve değiştir**: ctrl+F kaynakta arar; Tümünü Değiştir geri yazar.
- **AI**: hazır düğmeler, yardımcının belgeyi yeniden yazmasını, genişletmesini veya çevirmesini sağlar. Bir seçim AI'dan gizlenebilir (bkz. AI paneli bölümü).

## Araç çubuğu

Düzenleyicinin üstünde tek sıra düğme (ipuçları için üzerine gelin):

![Markdown araç çubuğu](img/md-toolbar.png)

- **Dosya ve geçmiş**: Kaydet, Farklı Kaydet, Geri Al, Yinele, Bul; sağdaki **Otomatik Kaydet** anahtarı değişiklikleri bir zamanlayıcıyla diske yazar.
- **AI düğmesi**: AI panelini açar; yanında yeniden yaz / genişlet / çevir hazır ayarları bulunur.
- **Paragraf stili** (açılır liste): gövde metni ile başlık düzeyleri arasında geçiş yapın.
- **Satır içi biçimlendirme**: **kalın**, _italik_, ~~üstü çizili~~, `satır içi kod`, bağlantı.
- **Listeler**: madde işaretli, numaralı, görev listesi.
- **Ekle**: tablo, görsel, yatay çizgi.
- **Özellikler**: dosyanın başındaki YAML ön madde bloğunu ekleyin veya ona atlayın.
- **Ana hat**: başlık hiyerarşisine göre atlayın.
- **Yazım denetimi**: bu belge için yazım denetimini açıp kapatın.
- **Görsel barındırma ayarları**: yapıştırılan ve sürüklenen resimleri önce kendi depolamanıza gönderin — S3 uyumlu (R2/OSS/COS/MinIO), SM.MS veya GitHub — yerel varlık kopyası yedek olarak kalır.

Üç hızlı örnek:

- **Başlık**: imleci satıra koyun ▸ paragraf stili açılır listesi ▸ "Başlık 1".
- **Tablo**: **Tablo ekle**'ye tıklayın ▸ satır/sütun sayısını sürükleyerek seçin ▸ hücrelere yazın; önizleme hemen işler.
- **Görev listesi**: birkaç satır seçin ▸ **Görev listesi**'ne tıklayın ▸ her satır `- [ ]` olur ve önizlemede onay kutusu olarak görünür.

## Dışa aktarma

Dosya menüsünde, hepsi yerel ve hepsi sonucu nereye koyacağını sorar:

- **Word olarak dışa aktar…** ve **PDF olarak dışa aktar…** gerçek bir .docx veya .pdf yazar.
- **Görsel olarak dışa aktar…** seçtiğiniz bir dizine içerik öğesi başına bir PNG yazar.
- **Dönüştür ve Docs'ta Aç** belgeyi .docx'e dönüştürüp burada, uygulamanın içindeki yerleşik Docs sekmesinde açar — buluttaki hiçbir şeye devredilmez ve dönüştürülmüş kopya, yaklaşık bir hafta sonra temizlenen bir önbellek klasöründe yaşar.

## Kaynak görünümü

Şerit bir **Kaynak** anahtarı taşır (uygulamayla birlikte yerelleştirilmiştir). Açtığınızda düzenleyicinin yerini ham Markdown alır: tam olarak kaydetmenin yazdığı metin; güzelleştirilmiş, altınızdan normalleştirilmiş hiçbir şey yok.

- **Düzenleme bayt düzeyinde sadıktır.** Kaynak görünümünden yapılan bir kaydetme, düzenleyiciden yapılan kaydetmeyle aynı baytları üretir — BOM, CRLF ve sondaki satır sonunun varlığı korunur.
- **Aynı belgedir.** İleri geri özgürce geçiş yapın; kaynak, düzenleyicinin kendi metnidir, birleştirilmesi gereken bir kopya değildir.
- **Biçimlendirme araç çubuğu kullanılamaz** görünüm açıkken, çünkü o düğmelerin çoğu yalnızca işlenmiş tarafta anlam taşıyan düzenleyici yapıları ekler. Görünümü kapattığınızda geri gelir.
- **.txt ve .json** burada doğrudan kaynak görünümünde açılır: işlenecek bir şey yoktur, dolayısıyla kaynak belgenin _kendisidir_.
