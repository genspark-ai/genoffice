# Yazı tipleri: sistem yüzleri ve indirilebilir aileler

## Yazı tipi listeleri

Yazı tipi denetimleri her düzenleyicinin Giriş sekmesindedir (aşağıda: Slides'taki yazı tipi grubu):

![Slides şeridindeki yazı tipi grubu](img/fonts.png)

Her düzenleyicinin yazı tipi seçicisi şunları birleştirir: yerel olarak kurulu yazı tipleri + platformda yaygın adaylar (olağan Windows ve macOS aileleri, Batı ve CJK, CJK yazı tiplerinin kendi sistemlerinde bildirdiği yerelleştirilmiş adlar dahil). Sistem yerel yazı tiplerini sayabildiğinde liste gerçekten var olanlara göre gruplanır; aksi halde tam aday listesine düşer.

## İndirilebilir yazı tipleri (Slides)

- Slides yazı tipi seçicisi bir **yazı tipi kataloğu** açar: birçok yazı sistemini kapsayan seçilmiş bir OFL koleksiyonu; her aile normal ve kalın ağırlıklarla gelir.
- Birini seçmek onu CDN'den (sağlama toplamına sabitlenmiş) indirip uygulamanın yazı tipi deposuna kurar — sonrasında her belgede kullanılabilir, sistem düzeyinde kurulum gerekmez.
- Kurulumlar uygulamaya özel bir dizinde yaşar ve uygulamayla birlikte gider.

## CJK tipografisi

- Docs'taki CJK yazı tipleri Word ile hizalı noktalama sıkıştırma ve kinsoku satır kuralları alır (bkz. Docs bölümü).
- Kataloğun CJK aileleri Basitleştirilmiş/Geleneksel Çince, Japonca ve Koreceyi, serif ve sans olarak kapsar.

## Yerel yazı tiplerini yönetme

- Yerel yazı tipi dosyalarını kurma girdisi, diskteki .ttf/.otf dosyalarını uygulamanın yazı tipi deposuna yükler.
- Depo aileye göre arşivler; bir ailenin birden çok ağırlığı bir arada bulunabilir.
