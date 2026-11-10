# Dosya işlemleri: yeniden adlandırma, silme, dışa aktarma

Bu bölüm her düzenleyicinin paylaştığı dosya işlemlerini kapsar; her düzenleyicinin kendi dışa aktarma seçenekleri kendi bölümündedir.

Satırın **⋯ menüsü** (bir dosya satırının üzerine gelin) bu eylemleri toplar:

![Bir dosya satırının ⋯ menüsü](img/file-ops.png)

## Yeniden adlandırma

İki giriş noktası, tek denetim kümesi:

- Giriş ekranındaki satırın **⋯ ▸ Yeniden Adlandır** öğesi.
- Satır içinde yeniden adlandırmak için **bir dosya sekmesine çift tıklayın** (bkz. [Sekmeler ve pencere yönetimi](help://tabs-and-windows)).

Kurallar: uzantı otomatik olarak korunur; geçersiz karakterler, sondaki noktalar ve ayrılmış adlar (CON/NUL ve benzerleri) bir iletiyle reddedilir; aynı klasörde ad çakışması da engellenir. Diskteki gerçek dosya yeniden adlandırılır; son kullanılanlar ve yıldızlar onu izler.

## Silme

- Giriş **⋯ ▸ Sil**: hangi dosyaların silineceğini sorar, ardından onları işletim sisteminin geri yükleyebileceği **sistem çöp kutusuna** taşır. GenOffice bunun için kendi geri alma mekanizmasını tutmaz — kurtarma çöp kutusunun işidir, bir bildirimin değil.

## Çoğaltma

**⋯ ▸ Çoğalt**, aynı klasörde <ad> kopyası oluşturur; çakışmalarda otomatik olarak bir sayaç eklenir. Kopya sizin için açılmak yerine **Son Kullanılanlar**'a eklenir, böylece önünüzde değil bir tık uzağınızdadır.

## Kaydet ve Farklı Kaydet

- **⌘S / ctrl+S** geçerli dosyayı kaydeder; adsız bir dosya önce konum ve ad sorar.
- **Farklı Kaydet** yeni bir dosya yazar ve özgün dosyaya dokunmaz; sonraki düzenlemeler yeni dosyayı hedefler.
- Her kayıt atomiktir (geçici dosya + yeniden adlandırma); yazma sırasında çıkış dosyayı bozamaz.
- Otomatik kaydetme yalnızca ilk elle kayıttan sonra devreye girer (bkz. [Hızlı başlangıç](help://getting-started)).

## PDF'ye dışa aktarma

- **Docs**: Dosya ▸ PDF Olarak Dışa Aktar (veya şerit düğmesi), yerleştirildiği gibi sayfalandırılmış.
- **Slides**: dışa aktarma sayfa sayfa tarar, büyük sunular için ilerleme gösterir.
- **Sheets**: dışa aktarma yazdırma sayfalandırmasını izler.
- Dışa aktarmalar gizli bir pencerede işlenir ve seçtiğiniz yere kaydedilir.

## Word / görsel olarak dışa aktarma

- **PDF ▸ Word'e Dönüştür**: PDF'yi .docx'e çevirir (yerel olarak dönüştürülür; karmaşık düzenler elden geldiğince).
- **Docs** sayfaları görsel olarak dışa aktarabilir (sayfa başına PNG).

## Yazdırma

Her düzenleyicide Dosya ▸ Yazdır (⌘P/ctrl+P) sistem yazdırma iletişim kutusunu açar; PDF'ler geçerli sayfa sırası ve döndürmeleriyle yazdırılır.

## Adsız dosyalar nerede yaşar

İlk kayıtta seçtiğiniz konum dosyanın evidir; ondan önce belge yalnızca bellekte vardır. Otomatik kaydetme ancak o ilk kayıttan sonra devreye girer.
