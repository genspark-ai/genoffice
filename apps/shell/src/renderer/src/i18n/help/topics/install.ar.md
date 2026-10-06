# تثبيت GenOffice

كل نسخة منشورة على [صفحة الإصدارات](https://github.com/genspark-ai/genoffice/releases/latest). وجميعها مبنية من `main`؛ وحزّامات التثبيت لنظامي macOS وWindows موقّعة.

## اختر الملف المناسب لجهازك

| النظام | المتطلبات | الملف |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — معظم أجهزة PC | Windows 10+، Intel/AMD | حزمة `-x64.exe` |
| **Windows** على Arm | Windows 11 على Arm (Snapdragon X وما شابه) | حزمة `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64، glibc 2.34+ (Ubuntu 22.04 أو أحدث) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64، glibc 2.34+ (Fedora 35+، RHEL 9+، Leap 15.6+) | `.rpm` |
| **Linux** — أي شيء آخر | x86_64، glibc 2.34+، FUSE 2 | `.AppImage` |

تبقى الإصدارات الأقدم على صفحة الإصدارات.

## على macOS

افتح ملف `.dmg` واسحب **GenOffice** إلى مجلد التطبيقات. وفي نسخة Intel يطلب منك أول تشغيل تأكيد الفتح: انقر نقرًا يمينًا على التطبيق داخل التطبيقات ▸ فتح. هذا Gatekeeper أمام صورة قرص تبدو غير موقّعة، لا تنزيلًا تالفًا.

## على Windows

شغّل حزمة التثبيت `.exe`. فهي تضع GenOffice في قائمة ابدأ وتسجّل `.docx` و`.xlsx` و`.pptx` و`.pdf` وما شابهها، بحيث يفتح الملف بالنقر المزدوج.

## على Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — يعمل من مكانه، بلا خطوة تثبيت:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

يحتاج إلى بيئة تشغيل FUSE 2. وإن لم تكن ترغب في تثبيتها فشغّله مع `--appimage-extract-and-run`، وتعمل البيئة المعزولة من المجلد المستخرج.

## ثلاث طرق أخرى

تغلّف الطرق الثلاث الملفات المنشورة بدل إعادة بناء التطبيق، لذا فهي تتبع الملفات الثنائية الرسمية.

- **Flatpak** — يثبّت ملف `.deb` عبر آلية Flatpak المعروفة بـ `extra-data` ويشغّله تحت Zypak، الذي يهيّئ البيئة المعزولة الخاصة بـ Electron. وشجرة `/app` للقراءة فقط، لذا تمرّ التحديثات عبر Flatpak بدل نافذة التحديث داخل التطبيق.
- **Nix** — تعبير حزمة تبنيه من نسخة مصدرية عبر `nix-build packaging/nix`؛ انظر `packaging/nix`.
- **Docker** — صورة تحويل مجمّع بلا واجهة رسومية (`packaging/docker`). تحوّل شجرة كاملة من مستندات Office وMarkdown وHTML إلى PDF عبر مسار التصدير الخاص بالتطبيق، لذا تكون النتيجة هي نفسها التي تنتجها عملية ملف ▸ تصدير. إنها أداة CLI، لا خادمًا.

## سطر الأوامر يأتي معه

كل تثبيت يشحن أمر `genoffice` يتحدث إلى المحركات نفسها التي تستخدمها النافذة، فلا يتناقض الطرفان حول أي ملف. وعلى macOS وWindows يوجد داخل حزمة التطبيق؛ و`genoffice install-cli` يضعه في `PATH` لديك. انظر **سطر الأوامر والوكلاء** لمعرفة ما يمكنه فعله.