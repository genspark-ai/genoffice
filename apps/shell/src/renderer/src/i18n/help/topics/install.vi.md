# Cài đặt GenOffice

Mọi bản dựng đều được công bố trên [trang phát hành](https://github.com/genspark-ai/genoffice/releases/latest). Tất cả đều đến từ `main`; và trình cài đặt cho macOS và Windows đã được ký.

## Chọn tệp cho máy của bạn

| Nền tảng | Yêu cầu | Tệp |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — hầu hết PC | Windows 10+, Intel/AMD | trình cài đặt `-x64.exe` |
| **Windows** trên Arm | Windows 11 trên Arm (Snapdragon X và tương tự) | trình cài đặt `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 hoặc mới hơn) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — thứ khác | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Các bản cũ hơn vẫn nằm trên trang phát hành.

## Trên macOS

Mở tệp `.dmg` rồi kéo **GenOffice** vào Applications. Với bản dựng Intel, lần chạy đầu tiên sẽ yêu cầu bạn xác nhận việc mở: nhấp phải vào ứng dụng trong Applications ▸ Mở. Đó là Gatekeeper gặp một đĩa ảnh trông như không có chữ ký, chứ không phải tệp tải về bị hỏng.

## Trên Windows

Chạy trình cài đặt `.exe`. Nó đưa GenOffice vào menu Start và đăng ký `.docx`, `.xlsx`, `.pptx`, `.pdf` cùng các định dạng tương tự, để nhấp đúp vào một tệp là mở được tệp đó.

## Trên Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — chạy ngay tại chỗ, không cần bước cài đặt:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Nó cần môi trường chạy FUSE 2. Nếu bạn không muốn cài đặt, hãy chạy với `--appimage-extract-and-run`, và hộp cát sẽ hoạt động từ thư mục đã giải nén.

## Ba cách nữa

Cả ba đều bọc các tệp đã phát hành thay vì dựng lại ứng dụng, nên chúng bám theo tệp nhị phân chính thức.

- **Flatpak** — cài tệp `.deb` qua cơ chế `extra-data` của Flatpak và chạy dưới Zypak, bộ điều hợp hộp cát riêng của Electron. Cây `/app` là chỉ đọc, nên các bản cập nhật đi qua Flatpak chứ không qua hộp thoại cập nhật trong ứng dụng.
- **Nix** — một biểu thức gói mà bạn dựng từ một bản sao mã nguồn bằng `nix-build packaging/nix`; xem `packaging/nix`.
- **Docker** — một ảnh chuyển đổi hàng loạt không giao diện (`packaging/docker`). Nó chuyển cả một cây tài liệu Office, Markdown và HTML sang PDF qua đúng dãy xuất của ứng dụng, nên kết quả đúng bằng thứ mà File ▸ Xuất sẽ tạo ra. Đây là tiện ích CLI, không phải máy chủ.

## Dòng lệnh đi kèm luôn

Mọi bản cài đặt đều kèm lệnh `genoffice` nói chuyện với đúng những engine mà cửa sổ dùng, nên cả hai không bao giờ bất đồng về một tệp. Trên macOS và Windows nó nằm trong gói ứng dụng; và `genoffice install-cli` đưa nó vào `PATH` của bạn. Xem **Dòng lệnh và tác nhân** để biết nó làm được gì.