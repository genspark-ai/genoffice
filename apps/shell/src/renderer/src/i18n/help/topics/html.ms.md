# Editor HTML

Editor HTML membuka fail .html dan .htm dengan dua mod: **pratonton** untuk halaman yang dipaparkan, dan **sumber** untuk teks asalnya.

- **Pratonton**: memaparkan halaman seperti mana ia akan kelihatan, dan stylesheet serta imej relatif dimuatkan dari lokasi berhampiran fail.
- **Pemeriksa pratonton**: klik untuk memilih elemen, klik dua kali untuk menyunting teksnya di tempat, butang pada bar alat untuk memadam, dan Tanya AI untuk bertanya tentang pilihan semasa.
- **Mod sumber**: menyunting HTML secara langsung, dengan ctrl+F untuk mencari dan Ganti Semua untuk menyimpan penanda yang ditulis semula.
- **Simpanan** mengekalkan setiap bait dengan tepat, termasuk BOM, CRLF, dan baris akhir. Simpanan tanpa perubahan tidak menulis semula fail.
- **Zum**: ctrl dan roda tetikus, atau mencubit dengan dua jari, akan menukar skala pratonton, dan ctrl+Z di dalam pratonton membatalkan suntingan terakhir.

## Bar alat

Klik pada mana-mana elemen dalam pratonton, dan satu bar alat akan muncul di atasnya:

![Bar alat terapung di atas elemen yang dipilih](img/html-toolbar.png)

- **Fail dan sejarah**: Simpan, Simpan Sebagai, Buat Asal, Buat Semula, Cari, dan Togol **Autosimpan** yang menulis perubahan ke cakera secara berkala.
- Penukar **Pratonton / Sumber** menukar antara paparan halaman dan kodnya, manakala butang **Presentasikan** memaparkan halaman dalam skrin penuh.
- **Pemformatan** merangkumi teks tebal, condong, serta butang untuk membesarkan dan mengecilkan saiz fon. **Panel gaya** tersedia untuk elemen yang dipilih, yang menyediakan warna dan banyak lagi.
- **Sisip** membolehkan anda menambah tajuk, perenggan, jadual, imej melalui pautan, dan unsur-unsur lain.
- **Tindakan imej** apabila imej dipilih: memotong, **mengalih keluar latar belakang**, menggantikan imej, dan mengunci nisbah bidang.
- **Tindakan elemen** apabila elemen dipilih dalam pemeriksa pratonton: memadam, menyalin, dan memindahkan ke atas atau ke bawah.
- **Butang AI** membuka panel AI, dan anda boleh bertanya tentang elemen yang dipilih secara terus.

## Eksport

Menu Fail, semuanya setempat dan semuanya akan bertanya di mana hasilnya perlu disimpan:

- **Eksport sebagai Word…** dan **Eksport sebagai PDF…** menulis .docx atau .pdf yang sebenar.
- **Eksport sebagai HTML fail tunggal…** menulis satu .html dengan imej dibbenamkan di dalamnya. Ia tidak akan menimpa fail yang sedang anda buka, dan memberitahu anda berapa imej yang tidak dapat dibbenamkan.

## Sisip kerangka

Bagi halaman kosong, **Sisip ▸ Sisip kerangka** menulis dokumen minimum dalam mod piawaian:

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

Setiap bahagian itu ada sebabnya, dan sebab itulah ia sebuah arahan dan bukan sesuatu yang perlu taip:

- **doctype**, atau pratonton akan berjalan dalam mod quirks, yang menyebabkan saiz kotak dan susun atur jadual mengikut peraturan yang berbeza daripada yang anda jangka;
- **`lang`**, atau pembaca skrin tiada bahasa untuk membaca halaman itu, dan pelayar akan memilih fon serta pemeriksa ejaan untuk bahasa yang salah;
- **charset**, atau halaman yang mengandungi teks bukan Latin boleh dinyahkod sebagai mojibake.

Meta viewport sengaja tidak disertakan: ini dipaparkan dalam panel desktop tanpa sebarang viewport mudah alih yang perlu dipengaruhinya.

`lang` mengikut bahasa antara muka aplikasi, jadi kerangka yang anda sisip ialah kerangka yang alat anda sudah pun disediakan untuknya. Anda boleh menyuntingnya dengan bebas selepas itu.

Baris item itu hanya muncul dalam mod suntingan, dan hanya semasa dokumen itu kosong — tiada apa-apa untuk disisipkan kerangka _ke dalam_ sebaik sahaja ada kandungan.
