# AI asistan paneli

Her düzenleyici AI panelini çağırabilir: bir şey seçin, bir yönerge verin, akan sonucu izleyin.

## Açma ve kullanma

![Docs içindeki AI paneli](img/ai-panel.png)

- Girişler: her düzenleyicinin şeridindeki **AI düğmesi**, bağlam menülerindeki **AI'ya Sor** veya işaretleme çubuğundaki AI'ya Sor.
- Görevi düz bir dille tarif edin (bunu yeniden yaz / bu sütunu yüzdeye çevir / bu sayfayı yeniden düzenle...) ve Enter'a basın.
- Yanıtlar **akış halinde** gösterilir; AI'nın araçlara ihtiyacı olduğunda (belgeyi okuma, düzenleme) bunları çalıştırır ve iş bitene kadar devam eder.
- **Durdur**: geçerli turu istediğiniz an kesin.
- **Kuyruğa alma**: bir yanıt akarken Enter, iletinizi bunun yerine kuyruğa alır — girişin üstündeki bir şerit sayılarını gösterir; düzenlemek, kaldırmak, yeniden sıralamak, duraklatmak veya temizlemek için açın. Tur bitince sıradaki ileti kendiliğinden gönderilir; Esc yine turu durdurur.

## Neler yapabilir

- **Docs**: yeniden yazma/genişletme/çeviri/özetleme, tablo ve resim ekleme, biçimlendirme ayarı; her tur önce anlık görüntü alır.
- **Sheets**: formüller, veri doldurma, toplu dönüşümler, biçimlendirme.
- **Slides**: tüm sunu oluşturma, düzen ayarlama, metin yeniden yazma.
- **PDF**: seçili metin veya sayfalar üzerinde soru-cevap ve özetler.
- **Markdown / HTML**: yeniden yazma, genişletme, çeviri.
- **Giriş ▸ AI'ya Sor**, yıldızlı dosyalarınızdan yanıt verir ve bunlara atıfta bulunur — bilgi tabanı panelinin Giriş bölümünde kendi kısmı vardır ve Docs AI paneli aynı beceriyi taşır.

## Geri alma ve güvenlik

- Docs paneli bir **sürüm listesi** tutar: her tur için bir anlık görüntü, istediğinize geri dönebilirsiniz ve geri dönüşün kendisi de Ctrl+Z ile geri alınabilir. Anlık görüntüler belge yeniden açıldığında da kalır.
- AI düzenlemeleri elle yapılanlarla aynı düzenleme hattından geçer (geri alınabilir, kayıt onayına bağlı) — hiçbir şey kayıt onayınızı atlamaz.

## Bir seçimi AI'dan gizleme

- Slides, HTML ve Markdown'da bir seçim modelin gözünden uzak tutulabilir: sağ tıklayın ▸ **Seçimi AI'dan gizle**, yer tutucuya ad verin; model onun yerine `{{ad}}` okur. (Markdown bu öğeyi şeridindeki göz-kapalı düğmesiyle etkinleştirir; HTML'in öğesi kaynak panelindedir.)
- Dosya gerçek metni korur — kaydetme, dışa aktarma ve yazdırma etkilenmez ve işaret yeniden açıldığında kalır. Aynı öğe gizlemeyi kaldırır (Markdown'da ayrı bir **Seçimi AI'ya yeniden göster** girdisi vardır).
- HTML; öğe içeriklerini, öznitelik değerlerini ve betik dize sabitlerini gizler; işlenen sayfa asla değişmez.
- Slides ayrıca resimleri, videoları ve sesleri bütün olarak gizler; gizli içeriği olan bir slayt asla ekran görüntüsü olarak gönderilmez: otomatik düzen denetimi o sayfayı atlar ve bunu bildirir.
- Bu yalnızca uygulama içi AI içindir — CLI, MCP araçları ve başsız dışa aktarma dosyayı doğrudan okur ve metni yine görür.

## Gizlilik

- Yönergeler ve ilgili belge içeriği **yapılandırdığınız model hizmetine** gider (Genspark barındırmalı veya özel bir uç nokta, sonraki bölüm); yapılandırma yoksa gönderim de yoktur.
- Yerel dosyalar başka hiçbir yere yüklenmez; BYOK anahtarları bu makinedeki uygulama ayar dosyasında saklanır ve yalnızca istek başlıklarında gönderilir.
