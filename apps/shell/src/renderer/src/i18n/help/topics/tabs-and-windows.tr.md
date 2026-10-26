# Sekmeler ve pencere yönetimi

Açık olan her dosya tek bir pencereyi paylaşır; üstteki sekme şeridi, tarayıcı tarzında aralarında geçiş yapar.

![Sekme şeridi: Giriş artı iki belge](img/tabs.png)

## Temeller

- **Geçiş**: bir sekmeye tıklayın ya da şeridin üzerinde tekerleği kaydırarak sırayla dolaşın.
- **Kapatma**: sekmedeki × veya ⌘W/ctrl+W. Kaydedilmemiş değişiklikler önce sorar.
- **Yeni**: şeridin sağ ucundaki +.
- **Etkinleştirme anındadır**: bir sekmeye basmak, tıklamanın bitmesini beklemeden basıldığı anda geçiş yapar.

## Yeniden adlandırma: sekmeye çift tıklayın

**Herhangi bir dosya sekmesine çift tıklayın**, başlık yerinde bir giriş alanına dönüşür: Enter onaylar, Escape iptal eder, odak kaybı onaylar; bir IME adayını onaylayan Enter, onay sanılmaz. Yeniden adlandırma, Giriş'teki satır yeniden adlandırmayla aynı güvenlik denetimlerinden geçer (geçersiz karakterler, ad çakışmaları), dosyayı diskte yeniden adlandırır ve son kullanılanlar listesini yeniler.

## Yeniden sıralamak için sürükleyin

Bir sekmeyi basılı tutup yana sürükleyerek yeniden sıralayın; komşular canlı olarak kenara kayar ve sıra bıraktığınızda yerleşir. 4 piksellik bir ölü bölge, düz tıklamaların sekmeleri kıpırdatmasını önler.

## Bağlam menüsü

Bir sekmeye sağ tıklayın: **Yeni Pencerede Aç** — ayrılabilen her belge sekmesinde sunulur — ve kapatılamayan bir sekmede gri görünen **Kapat**. İkisi de uygulama dilini izler.

## Pencereye ayırma

**Bir sekmeyi şeridin dışına sürükleyin** (ya da Yeni pencerede aç'ı kullanın), canlı belgeyi taşıyan kendi penceresi olur; şeride yeniden yerleştirmek için geri sürükleyin. Ayrılmış bir pencere tam yetkili bir eştir — orada düzenlemeye ve kaydetmeye devam edin.

## Tüm sekmeler listesi

Sekmeler taştığında, şeridin sağ ucundaki küçük sekme çubuğu simgesi tam listeyi açar (içerik alanının asla örtmediği yerel bir menü); ok tuşlarıyla seçin.

## Her düzenleyicinin içindeki araç çubuğu satırı

Her düzenleyici sekmesinin üstünde bir araç çubuğu satırı vardır (tam yerleşim düzenleyiciye göre biraz değişir):

- **Kaydet** (⌘S/ctrl+S) ve **Farklı Kaydet**.
- **Geri Al / Yinele**: çok düzeyli; her düzenleyici kendi geçmişini tutar.
- **Bul** (ctrl+F): o düzenleyicinin bul panelini açar.
- **Otomatik Kaydet** anahtarı: açıkken değişiklikler bir zamanlayıcıyla diske yazılır; kapalıyken yalnızca elle kaydetmeler yazar (sekmedeki/başlıktaki kaydedilmemiş işareti size hatırlatır).

## Giriş sekmesi

En soldaki Giriş sekmesi kapatılamaz; herhangi bir düzenleyiciden geri dönmek için ona tıklayın ya da Dosya ▸ **Giriş'e Dön**'ü kullanın.
