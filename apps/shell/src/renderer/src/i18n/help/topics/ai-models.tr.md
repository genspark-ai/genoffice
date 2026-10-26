# AI modelleri ve ayarları

## Sağlayıcılar ve modeller

Modeller ve anahtarlar Ayarlar'da yapılandırılır (Giriş ekranının sol altındaki hesap satırı):

![Ayarlar penceresi](img/settings-general.png)

- **Genspark barındırmalı**: oturum açın (cihaz kodu akışı) ve kullanın — sıfır yapılandırma.
- **Özel uç noktalar (BYOK)**: Ayarlar ▸ AI, her protokol için bir temel URL ve API anahtarı alır — OpenAI uyumlu, Anthropic, Gemini, DeepSeek, DashScope (qwen) ve daha fazlası. Anahtarlar bu makinedeki uygulama ayar dosyasında saklanır ve yalnızca istek başlıklarında gönderilir. Atlas Cloud yerleşik protokoller arasındadır.
- Her yetenek için farklı bir model seçilebilir: sohbet/üretim, görsel oluşturma, görsel analizi.
- **Bağlantı testi**: kaydetmeden önce uç noktanın erişilebilirliğini ve modellerin görünürlüğünü doğrulayın.
- Temel URL'ler yol ve sorgu dizesi taşıyabilir (ağ geçidi tarzı); uç nokta yolları doğru şekilde eklenir.

## CLI entegrasyonu (Codex sınıfı)

- Ayarlar yerel bir CLI program yolu kabul eder (ASCII dışı ev dizinleri ve ~ öneki çalışır; ~ otomatik olarak genişletilir); Modelleri algıla, CLI'ın kullanılabilir modellerini sorgular.
- Doğrulama yalnızca dosyanın varlığını denetler — karakter kümesi kısıtlaması yoktur.

## Değişiklikler ne zaman uygulanır

- Model ve uç nokta değişiklikleri hemen uygulanır; devam eden bir konuşma bir sonraki turuna kadar eski yapılandırmayı korur.
