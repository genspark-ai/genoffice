# การติดตั้ง GenOffice

ทุกเวอร์ชันถูกเผยแพร่ไว้ที่ [หน้าเผยแพร่](https://github.com/genspark-ai/genoffice/releases/latest) ทุกเวอร์ชันมาจาก `main` ตัวติดตั้งสำหรับ macOS และ Windows มีลายเซ็นกำกับ

## เลือกไฟล์ให้ตรงกับเครื่องของคุณ

| แพลตฟอร์ม | ข้อกำหนด | ไฟล์ |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — คอมพิวเตอร์ส่วนใหญ่ | Windows 10+, Intel/AMD | ตัวติดตั้ง `-x64.exe` |
| **Windows** บน Arm | Windows 11 บน Arm (Snapdragon X และรุ่นใกล้เคียง) | ตัวติดตั้ง `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 หรือใหม่กว่า) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — อื่น ๆ | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

เวอร์ชันเก่ากว่ายังอยู่ในหน้าเผยแพร่

## บน macOS

เปิดไฟล์ `.dmg` แล้วลาก **GenOffice** ไปไว้ใน Applications ถ้าเป็นเวอร์ชัน Intel การเปิดครั้งแรกจะถามให้คุณยืนยันการเปิด คลิกขวาที่แอปใน Applications แล้วเลือก เปิด นี่คือ Gatekeeper ที่กำลังเจออิมเมจดิสก์ที่ดูเหมือนไม่ได้ลงลายเซ็น ไม่ใช่ไฟล์ที่ดาวน์โหลดเสีย

## บน Windows

รันตัวติดตั้ง `.exe` ตัวนี้จะวาง GenOffice ไว้ในเมนู Start และลงทะเบียน `.docx`, `.xlsx`, `.pptx`, `.pdf` และไฟล์อื่น ๆ ให้เป็นค่าเริ่มต้น ดังนั้นดับเบิลคลิกที่ไฟล์ก็จะเปิดขึ้นมา

## บน Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — รันจากที่วางไว้ได้เลย ไม่ต้องมีขั้นตอนติดตั้ง

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

ตัวนี้ต้องมีรันไทม์ FUSE 2 ถ้าไม่อยากติดตั้ง ให้รันด้วย `--appimage-extract-and-run` แซนด์บ็อกซ์จะทำงานได้จากโฟลเดอร์ที่แตกออกมา

## อีกสามทาง

ทั้งสามทางห่อหุ้มไฟล์ที่เผยแพร่แล้ว ไม่ได้ build แอปใหม่ จึงตามหลังไฟล์ไบนารีทางการไปเสมอ

- **Flatpak** — ติดตั้ง `.deb` ผ่านกลไก `extra-data` ของ Flatpak แล้วรันภายใต้ Zypak ซึ่งปรับแซนด์บ็อกซ์ของ Electron ให้เข้ากันได้ โครงสร้าง `/app` อ่านได้อย่างเดียว การอัปเดตจึงต้องไปที่ Flatpak ไม่ใช่กล่องอัปเดตในแอป
- **Nix** — นิยามแพ็กเกจที่ build จากโค้ดที่ clone มา โดยใช้ `nix-build packaging/nix` ดูที่ `packaging/nix`
- **Docker** — อิมเมจสำหรับแปลงไฟล์แบบไม่มีหน้าต่าง (`packaging/docker`) แปลงเอกสาร Office, Markdown และ HTML ทั้งโฟลเดอร์เป็น PDF ผ่านช่องทางส่งออกของตัวแอปเอง ผลลัพธ์จึงเหมือนกับที่ได้จาก ไฟล์ ▸ ส่งออก นี่คือยูทิลิตี้บนบรรทัดคำสั่ง ไม่ใช่เซิร์ฟเวอร์

## บรรทัดคำสั่งมาให้ด้วย

ทุกการติดตั้งมีคำสั่ง `genoffice` ที่คุยกับเอนจินตัวเดียวกับที่หน้าต่างใช้ ทั้งสองฝ่ายจึงไม่มีวันขัดแย้งกันเรื่องไฟล์ บน macOS และ Windows คำสั่งนี้อยู่ในแพ็กเกจของแอป ถ้าต้องการเรียกด้วยชื่อ ให้รัน `genoffice install-cli` ครั้งเดียว คำสั่งจะไปอยู่ใน `PATH` ของคุณ ดู **บรรทัดคำสั่งและเอเจนต์** ว่ามันทำอะไรได้บ้าง
