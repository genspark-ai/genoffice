# Ayarlar, dil, tema ve MCP tümleştirmeleri

## Ayarları açma

Giriş'in sol altındaki hesap satırı ayarlar panelini açar (oturum kapalıyken Oturum aç yazar). Altı bölümü vardır: Hesap, AI Modeli, AI Medya ve Arama, Genel, Tümleştirmeler ve Hakkında.

![Ayarlar ▸ Genel: dil, tema, Otomatik Kaydet ve kullanım istatistikleri anahtarının bulunduğu yer](img/settings-general.png)

AI modeli yapılandırmasının kendi makalesi vardır; **AI Medya ve Arama**, görsel oluşturma, görsel analizi, video analizi, web araması ve yerel dosya aramasını sağlayıcı bazında açtığınız yerdir.

## Dil

- Ayarlar **22 arayüz dili** sunar: İngilizce, Basitleştirilmiş Çince, Japonca, Korece, Fransızca, Almanca, İspanyolca, Tayca, Endonezce, Rusça, Arapça, Portekizce, İtalyanca, Lehçe, Çekçe, Felemenkçe, Malayca, İbranice, Hintçe, Geleneksel Çince, Vietnamca, Türkçe.
- Geçiş hemen uygulanır ve kalıcıdır; yerel menü çubuğu dille birlikte yeniden oluşturulur.

## Tema

Açık / Koyu / Sistem. Sistem, işletim sisteminin görünümünü izler ve düzenleyiciler titreşmeden eşzamanlı olarak yeni görünüme geçer.

## Genel

- **Anonim kullanım istatistikleri gönder** — varsayılan olarak açık. Google Analytics 4 kullanır ve genel IP adresinizi ile aktarım meta verilerini gönderir; belge içerikleri ve dosya adları asla toplanmaz ve her olay yalnızca ".docx açıldı" gibi bir tür taşır. İstediğiniz zaman buradan kapatın.
- **AI kenar çubuğu konumu** (sol veya sağ), **AI paneli metin boyutu** ve **AI sohbetinde yazım denetimi**.
- **Yeni belgelerde AI panelini aç** — kapalıyken yeni bir belge, panel daraltılmış ve tek tık uzakta olarak başlar.
- **Tüm belgeleri otomatik kaydet** her düzenleyicide Otomatik Kaydet'i varsayılan olarak açar; tek bir pencere için yine de kapatabilirsiniz.
- **Değiştir** düğmesiyle **Kayıt Konumu** ve .docx / .xlsx / .pptx dosyalarını GenOffice'e atamak için **Office belgeleri için varsayılan uygulama**.

## AI Medya ve Arama

Anahtar değil — her yetenek ona hizmet eden satıcıyı seçer ve bir satıcının anahtarı ile temel URL'si bir kez girilip paylaşılır:

- **Web araması**, **Görsel oluşturma**, **Görsel analizi** ve **Video analizi**, her biri bir sağlayıcı, bir model, bir anahtar ve bir temel URL ile.
- **Yerel dosya araması** bu makinede çalışır. Altındaki **Jev arama yeniden sıralaması** **varsayılan olarak kapalıdır**. Açtığınızda, en iyi 20 yerel eşleşmenin alıntıları — her belgeden en çok 1.200 karakter, artı dosya ve klasör adları — seçtiğiniz karar uç noktasına (OpenRouter veya TypeSafe üzerinden Jev, Perplexity, Liquid AI d1, Cloudflare, yerel bir motor ya da özel bir URL) ilgiye göre yeniden sıralanmak üzere gider. Kapalıyken cihazdan hiçbir şey çıkmaz.

## Hakkında

- **Sürüm**, projenin GitHub bağlantısı ve bir **GitHub'da Yıldızla** düğmesi.
- **Güncelleme Kanalı**: Kararlı veya Beta. Değiştirmek hemen etkili olur ve güncelleme denetler; bir Beta kurulumunu Kararlı'ya geri düşürmez.

## Varsayılan uygulama bağlamaları

Ayarlar, GenOffice'i .docx / .xlsx / .pptx / .pdf ve benzerleri için işleyici olarak kaydedebilir (platform düzeyinde varsayılan uygulama kaydı; istendiğinde onaylayın).

## Üçüncü taraf bildirimleri ve güncellemeler

- Yardım ▸ Üçüncü taraf bildirimleri: uygulamayla gönderilen tam açık kaynak lisans envanteri.
- Yardım ▸ Güncellemeleri denetle: elle denetimi tetikler; daha yeni bir sürüm yükleme ister.

## Genspark oturum açma

- Oturum açma girişi (ayarlar veya bulut proje listesi) bir **cihaz kodu** akışı kullanır: GenOffice bir kod gösterir ve tarayıcıda oturum açma sayfasını açar; tamamlandığında otomatik olarak devam eder.
- Oturum açma yalnızca şunlar için kullanılır: bulut proje listesi ve Genspark'ın barındırılan modelleri. Onsuz her yerel özellik ve özel modeller çalışmaya devam eder.
- Oturumu kapatmak ayarlarda tek tıktır.

## MCP tümleştirmesi (ileri düzey kullanıcılar / AI istemcileri için)

**Tümleştirmeler**, GenOffice'i bir kodlama aracısına bağlayan bölmedir ve kendi makalesi vardır: Bir kodlama aracısını bağlama. Kısa sürüm — bir yol seçin (komut satırı veya MCP), o bölümü izleyin, sonra yeni bir sohbet başlatıp sorun.

![Ayarlar ▸ Tümleştirmeler: üç adım, ardından beceri satırları ve MCP seçenekleri](img/settings-integrations.png)

**Yerel HTTP sunucusu** altında uygulama sunucuyu kendisi çalıştırabilir — bir etkinleştirme anahtarı, bir bağlantı noktası — ve **Gelişmiş**, durum denetimi URL'si ile günlük dosyasını ekler. Yalnızca localhost üzerinde dinler.

## Komut satırı kopya kağıdı

| Komut              | Ne yapar                      |
| ------------------ | ----------------------------- |
| `genoffice <file>` | bir dosya açar                |
| `genoffice mcp`    | yerel MCP sunucusunu başlatır |
| `genoffice --help` | tüm komutlar ve bayraklar     |
