# HTML düzenleyicisi

HTML düzenleyicisi .html / .htm dosyalarını iki modla açar: **önizleme** (işlenmiş sayfa) ve **kaynak**.

- **Önizleme**: gerçek işleme; göreli stil sayfaları ve görseller dosyanın yanından yüklenir.
- **Önizleme denetçisi**: bir öğeyi seçmek için tıklayın, metnini yerinde düzenlemek için çift tıklayın, araç çubuğundan silin ve seçim üzerinde AI'ya Sor.
- **Kaynak modu**: HTML'i düzenleyin; ctrl+F bulma, Tümünü Değiştir yeniden yazılan işaretlemeyi kaydeder.
- **Kaydetme**: bayt sadık (BOM/CRLF/sondaki yeni satır korunur); değişmemiş kayıtlar dosyayı yeniden yazmaz.
- **Yakınlaştırma**: ctrl+tekerlek / kıstırma önizlemeyi ölçekler; önizleme içinde ctrl+Z son düzenlemeyi geri alır.
- **AI'dan gizle**: kaynak modunda bir seçime sağ tıklayın — öğe içerikleri, öznitelik değerleri, betik dizeleri — ve model bunun yerine bir yer tutucu görür; işlenen sayfa asla değişmez (ayrıntılar AI paneli bölümünde).

## Araç çubuğu

Önizlemede herhangi bir öğeye tıklayın, üstünde bir araç çubuğu belirir:

![Seçili bir öğenin üstündeki yüzen araç çubuğu](img/html-toolbar.png)

- **Dosya ve geçmiş**: Kaydet, Farklı Kaydet, Geri Al, Yinele, Bul; **Otomatik Kaydet** anahtarı değişiklikleri zamanlayıcıyla yazar.
- **Önizleme / Kaynak** anahtarı; **Sun** sayfayı tam ekran gösterir.
- **Biçimlendirme**: kalın, italik, yazı tipi boyutunu büyüt/küçült; seçili öğe için **stil paneli** (renkler ve daha fazlası).
- **Ekle**: başlık, paragraf, tablo, görsel (bağlantıyla), daha fazlası.
- **Görsel eylemleri** (bir görsel seçiliyken): kırp, **arka planı kaldır**, değiştir, en boy oranını kilitle.
- **Öğe eylemleri** (önizleme denetçisinde bir öğe seçiliyken): sil, çoğalt, yukarı/aşağı taşı.
- **AI düğmesi**: AI panelini açar; seçili öğe hakkında doğrudan sorun.

## Dışa aktarma

Dosya menüsü, hepsi yerel ve hepsi sonucu nereye koyacağını sorar:

- **Word Olarak Dışa Aktar…** ve **PDF Olarak Dışa Aktar…** gerçek bir .docx veya .pdf yazar.
- **Tek Dosyalık HTML Olarak Dışa Aktar…** görselleri içine gömülmüş tek bir .html yazar. Şu anda açık olan dosyanın üzerine yazmaz ve kaç görseli gömemediğini söyler.

## İskelet ekleme

Boş bir sayfa için **Ekle ▸ İskelet ekle** en küçük standart mod belgesini yazar:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Her parçanın bir nedeni vardır; bu yüzden elle yazılacak bir şey değil, bir komuttur:

- **doctype**, yoksa önizleme quirks modunda çalışır ve kutu boyutlandırma ile tablo düzeni beklediğinizden farklı kurallara uyar;
- **`lang`**, yoksa ekran okuyucunun sayfayı okuyacağı bir dil olmaz ve tarayıcı yanlış dil için yazı tipi ve yazım denetleyicisi seçer;
- **charset**, yoksa Latin dışı metin içeren bir sayfa bozuk karakterler olarak çözülebilir.

Viewport meta'sı bilinçli olarak yoktur: bu, etkileyeceği bir mobil görünüm alanı olmayan bir masaüstü bölmesinde işlenir.

`lang` uygulamanın arayüz dilini izler, böylece eklediğiniz iskelet araçlarınızın zaten ayarlı olduğu iskelettir. Sonrasında serbestçe düzenleyin.

Öğe yalnızca düzenleme modunda ve yalnızca belge boşken görünür — içerik varken iskeletin _içine_ ekleneceği bir şey yoktur.
