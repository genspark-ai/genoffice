# Editor Markdown

Editor Markdown membuka fail .md dan .markdown dengan pengalaman sumber dan pratonton yang dipaparkan.

- **Buka**: daripada Laman Utama atau melalui Fail ▸ Buka, dan baris arahan turut berfungsi.
- **Sunting**: penyuntingan teks biasa. Sambungan GFM seperti jadual, senarai tugasan, teks bergaris, dan pautan automatik dipaparkan dalam pratonton.
- **Pratonton**: dipaparkan secara langsung, dengan sumber relatif seperti imej diselesaikan berhampiran dokumen.
- **Simpanan** mengekalkan setiap bait dengan tepat. Tanda BOM, aksara CRLF, dan kehadiran baris akhir kosong semuanya dikekalkan, dan simpanan tanpa perubahan tidak menulis semula fail.
- **Cari dan ganti**: ctrl+F mencari dalam sumber, dan Ganti Semua menulis hasilnya ke fail.
- **AI**: Butang tetapan awal membolehkan pembantu menulis semula, melanjutkan, atau menterjemah dokumen tersebut.

## Bar alat

Satu baris butang di atas editor (halakan penuding ke atasnya untuk melihat tooltip):

![Bar alat Markdown](img/md-toolbar.png)

- **Fail dan sejarah**: Simpan, Simpan Sebagai, Buat Asal, Buat Semula, dan Cari. Togol **Autosimpan** di sebelah kanan menulis perubahan ke cakera secara berkala.
- **Butang AI** membuka panel AI, dan bersetunya terdapat tetapan awal untuk menulis semula, melanjutkan, dan menterjemah.
- **Gaya perenggan** dalam senarai lungsur membolehkan anda bertukar antara teks biasa dan pelbagai tahap tajuk.
- **Pemformatan dalam talian** merangkumi tebal, condong, teks bergaris, kod dalam talian, dan pautan.
- **Senarai** merangkumi senarai bertindan, bernombor, dan senarai tugasan.
- **Sisip** merangkumi jadual, imej, dan garis mendatar.
- **Sifat** memasukkan atau melompat ke blok YAML front matter di bahagian atas fail.
- **Kerangka** membolehkan anda melompat mengikut hierarki tajuk.
- **Semakan ejaan** menyalakan atau mematikan pemeriksaan ejaan bagi dokumen ini.

Tiga contoh pantas:

- **Tajuk**: letakkan kursor pada baris tersebut, pilih gaya perenggan daripada senarai lungsur, dan pilih "Tajuk 1".
- **Jadual**: klik **Sisip jadual**, seret untuk memilih bilangan baris dan lajur, kemudian taip ke dalam sel. Pratonton akan memaparkannya serta-merta.
- **Senarai tugasan**: pilih beberapa baris, klik **Senarai tugasan**, dan setiap baris akan menjadi `- [ ]` yang dipaparkan sebagai kotak semak dalam pratonton.

## Eksport

Menu Fail, semuanya setempat dan semuanya akan bertanya di mana hasilnya perlu disimpan:

- **Eksport sebagai Word…** dan **Eksport sebagai PDF…** menulis .docx atau .pdf yang sebenar.
- **Eksport sebagai imej…** menulis satu PNG bagi setiap halaman ke dalam folder yang anda pilih.
- **Tukar dan buka dalam Docs** menukar kepada .docx dan membukanya dalam tab Docs terbina dalam di dalam aplikasi ini — ia bukan serah kepada apa-apa dalam awan, dan salinan yang ditukar tinggal dalam folder cache yang dibersihkan selepas kira-kira seminggu.

## Paparan sumber

Reben itu membawa togol **Sumber** (dilocalkan bersama aplikasi). Hidupkan ia, dan editor akan digantikan oleh Markdown asal: betul-betul teks yang ditulis oleh sesuatu simpanan, tiada yang dicantikkan, dan tiada apa-apa yang dinormalkan di bawah anda.

- **Penyuntingan mengekalkan setiap bait dengan tepat.** Simpanan daripada paparan sumber menghasilkan bait yang sama seperti simpanan daripada editor — BOM, CRLF dan kehadiran baris akhir kosong semuanya kekal.
- **Ia dokumen yang sama.** Anda boleh bergantian antara paparan ini dan editor dengan bebas; sumber itu ialah teks editor itu sendiri, bukan salinan yang perlu digabungkan.
- **Bar alat pemformatan tidak tersedia** semasa ia terbuka, kerana kebanyakan butang itu memasukkan konstruksi editor yang hanya bermakna pada sisi yang dipaparkan. Ia kembali apabila anda menutup paparan tersebut.
- **Fail JSON dan fail mod sumber yang lain** dibuka di sini secara langsung: tiada apa-apa untuk dipaparkan, jadi sumber itu _ialah_ dokumennya.
