# Giriş ekranı: dosyalarınızın yaşadığı yer

Giriş, GenOffice'in başlangıç sayfasıdır: solda bir gezinti kenar çubuğu, sağda dosya listeleri ve hızlı oluşturma kartları.

![Giriş ekranı](img/home-screen.png)

## Kenar çubuğu gezintisi

- **Son Kullanılanlar**: yakın zamanda açtığınız dosyalar. Her satır ne zaman açıldığıyla damgalanır — bugün, dün veya tarih.
- **Yıldızlı**: yıldızladığınız dosyalar. Bir dosya satırının üzerine gelip yıldıza tıklayarak ekleyin veya kaldırın. ⋯ menüsündeki **Gruba taşı**, yıldızlı dosyaları adlandırılmış gruplara ayırır; Yıldızlı görünümü bunlara göre süzer.
- **AI'ya Sor**: bilgi tabanı paneli — sorular sorun ve yıldızlı dosyalarınızdan atıflı yanıtlar alın (aşağıya bakın).
- **Genspark Projeleri**: Genspark hesabınızda oturum açtıktan sonra, web'de Genspark AI ile oluşturduğunuz projeleri gösterir; düzenlemeye tarayıcıda devam etmek için birine tıklayın. Arama, zamana göre sıralama, yenileme ve daha fazlasını yükleme desteklenir.
- **Klasörler**: **Klasör ekle…** ile kenar çubuğuna eklediğiniz veya dosya yöneticisinden sürüklediğiniz dizinler. Her biri açabileceğiniz, içinde alt klasör oluşturabileceğiniz, yeniden adlandırıp kaldırabileceğiniz bir kök olur; çevrimdışı kalan biri kullanılamaz olarak görünür ve listeden çıkarılabilir. **Yeni klasör** bir tane daha oluşturur.

Burada Çöp Kutusu girdisi yoktur. Silinen dosyalar sistem çöp kutusuna gider ve birini geri yüklemek işletim sisteminin işidir.

## Dosya listesi

Her satır bir simge, dosya adı, değiştirilme zamanı ve daha fazlasını gösterir. Satırın **⋯ menüsü** şunları sunar:

- **Aç** ve dosyayı dosya yöneticinizde bulmak için **Klasörde göster**.
- **Yolu kopyala**.
- **Klasöre taşı…**: bir klasör seçici açar ve dosyayı gerçekten taşır; hedefte aynı adda bir dosya varsa atla / üzerine yaz / yeniden adlandır seçenekleri sunar.
- **Yeniden Adlandır**: satır içinde, uzantı otomatik olarak korunur.
- **Yıldızla / Yıldızı kaldır** — yıldızlar yeniden başlatmalarda kalır ve dosyayı yeniden adlandırdığınızda onu izler.
- **Çoğalt**: aynı klasörde bir kopya oluşturur.
- **Sil**: dosyayı sistem çöp kutusuna taşır — kalıcı silme değildir.
- **Listeden kaldır**, üst düzey Son Kullanılanlar görünümünde, dosyaya dokunmadan bir girdiyi düşürmek için.

### Aynı anda birden çok dosya

Bir seçim oluşturmak için bir satırdaki onay kutusunu işaretleyin veya ⌘/ctrl ile tıklayın; üstbilgi onay kutusu o anda listelenen her şeyi seçer ve listenin üstündeki bir çubuk kaç tanesinin seçildiğini bildirir, tüm küme için **Klasöre taşı…** ve **Dosyaları sil** sunar. Çoklu seçimi kenar çubuğundaki bir klasöre de sürükleyebilirsiniz.

## Arama

Üstteki arama kutusu aynı anda iki şeyi eşleştirir:

- **Dosya adları**: ada göre hızlı süzme.
- **Dosya içerikleri**: GenOffice dosyalarınızı arka planda dizinler (docx/xlsx/pptx/pdf/md/html içindeki metin), böylece gövde metni aramak da dosya bulur. Kapsam ve anahtarlar arama ayarlarındadır.

## Hızlı başlangıç kartları

Listelerin üstündeki kartlar tek adımda yeni bir belge oluşturur. Bir karta tıklamak o türde bir dosya oluşturur ve düzenleyicisini açar — hemen yazmaya başlayın ya da taslağı AI'ya bırakın (her düzenleyicinin şeridinde bir **AI düğmesi**, seçim bağlam menüsünde **AI'ya Sor** vardır).

Yeni dosya kenar çubuğunda o anda seçili klasöre düşer; seçim yoksa varsayılan klasöre gider.

Her kartın yaptığı:

- **AI Docs** (.docx): Docs düzenleyicisinde boş bir metin belgesi. Dosya diske yalnızca **ilk kayıtta** yazılır; yeni belgeler AI paneli genişletilmiş açılır (Ayarlar → "Yeni belgelerde AI panelini aç" altından kapatın).
- **AI Sheets** (.xlsx): Sheets düzenleyicisinde boş bir elektronik tablo. Kaydedene kadar diskte dosya yoktur — ad ilk kayıt için ayrılır; ilk AI üretiminden sonra dosya içeriğinden otomatik olarak da yeniden adlandırılabilir.
- **AI Slides** (.pptx): Slides düzenleyicisinde boş bir sunu.
- **AI Markdown** (.md): Markdown düzenleyicisinde boş bir Markdown belgesi.
- **AI HTML** (.html): HTML düzenleyicisinde boş bir web sayfası.
- **AI PDF** (.pdf): diğerlerinden farklıdır — hedef klasörde **hemen** gerçek, boş, tek sayfalık bir PDF oluşturur ve onu sıradan bir dosya olarak açar (PDF düzenleyicisi gerçek dosyalar üzerinde çalışır). Ek açıklama, karartma veya metin eklemek için uygundur; dosya ilk kayıtta içeriğinden otomatik olarak yeniden adlandırılabilir.
- **Yerel dosyaları aç**: Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) ve web sayfaları (.html/.htm) için bir sistem dosya seçici. Çoklu seçim çalışır; her dosya kendi sekmesini alır. Düz metin dosyaları (.txt, .json) da buradan açılır — Markdown düzenleyicisinin kaynak görünümüne düşerler.

> İpucu: Menü çubuğundaki Dosya ▸ Yeni aynı belge türlerini oluşturur (⌘N/Ctrl+N varsayılan olarak metin belgesi); bir dosyayı pencereye sürüklemek onu açar.

## Bilgi tabanınıza sorun

- Önce dosyaları yıldızlayın (uygulamanın okuduğu herhangi bir tür); panel başlığı derlemi sayar ve boş bir bilgi tabanı size gidip bir şey yıldızlamanızı söyler.
- Yanıtlar geldikleri dosyalara atıfta bulunur; dosyayı düzenleyicisinde açmak için bir atıfa tıklayın.
- Genspark'ta oturum açmayı gerektirir. Konuşma yalnızca bellekte yaşar — uygulamayı kapatmak onu sonlandırır.
- Arama, Giriş araması gibi yerel dosya dizininde çalışır; kaybolan dosyalar sessizce derlemden düşer.
- Gruplar yalnızca Yıldızlı listesini düzenler — bilgi tabanı her zaman tüm yıldızlı dosyaları kapsar.

## Bulut projeleri (Genspark Projeleri)

- İlk kullanım Genspark hesabınızda oturum açmayı gerektirir (cihaz kodu akışı: GenOffice bir kod gösterir, oturum açmayı tarayıcıda tamamlarsınız).
- Proje listesi web ile eşitlenir; Tarayıcıda aç, devam etmek için oraya atlar.
- Oturum açmamak yerel özelliklerin hiçbirini etkilemez.
