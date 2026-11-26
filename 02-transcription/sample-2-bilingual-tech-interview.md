# Audio Transcription Sample 2: Bilingual Tech Interview (ID — EN Code Switching)

**Subject:** Deep Dive Interview with Southeast Asian Fintech Co-Founder  
**Duration:** 04:30  
**Speakers:** 2 (Interviewer: Indonesian native; Founder: Bilingual English/Indonesian)  
**Language Environment:** Natural bilingual code-switching (Indonesian conversational syntax infused with English tech/venture capital terms)  
**Style Standard:** Standard Clean Verbatim; English domain terms italicized for editorial clarity.

---

## Transcript

**[00:00:00] Interviewer:**  
Selamat pagi, Mas Budi. Terima kasih banyak sudah meluangkan waktu di tengah jadwal yang padat. Mungkin sebagai pembuka, bisa diceritakan sedikit tentang *core problem* yang ingin diselesaikan oleh platform Anda di Indonesia?

**[00:00:15] Budi Santoso:**  
Pagi, terima kasih kembali sudah mengundang saya. Jadi, kalau kita bicara ekosistem UMKM di Asia Tenggara, khususnya di *second-tier* dan *third-tier cities*, *bottleneck* terbesarnya adalah *financial inclusion*. Bank konvensional umumnya memerlukan *collateral* dan *credit history* historis yang sering kali tidak dimiliki oleh pelaku usaha mikro.

**[00:00:39] Budi Santoso:**  
Pendekatan kami adalah memanfaatkan *alternative data*. Kami mengintegrasikan data transaksi harian, aktivitas *point-of-sale*, serta arus kas *supply chain* mereka. Dari situ, *underwriting engine* kami yang berbasis *machine learning* bisa menghasilkan *risk profile* secara *real-time*. Hasilnya, proses *loan approval* yang tadinya memakan waktu dua minggu bisa dipangkas menjadi hitungan jam.

**[00:01:06] Interviewer:**  
Menarik sekali. Dari sisi *engineering* dan arsitektur data, tantangan apa yang paling signifikan saat melakukan *scaling* sistem transaksi tersebut?

**[00:01:17] Budi Santoso:**  
Tantangan utamanya ada pada *network latency* dan *reliability* infrastruktur. Banyak *merchant* kami beroperasi di area dengan koneksi internet yang tidak stabil. Oleh karena itu, tim *engineering* kami merancang arsitektur aplikasi berbasis *offline-first*. Transaksi dicatat secara lokal dengan enkripsi *end-to-end*, lalu disinkronkan secara asinkron begitu perangkat mendapatkan sinyal seluler.

**[00:01:45] Budi Santoso:**  
Di sisi *backend*, kami beralih ke *event-driven microservices* menggunakan Apache Kafka untuk menangani *peak transaction volume*, terutama saat festival belanja tanggal kembar. *Data pipeline* kami memastikan tidak ada *duplicate reconciliation* antara catatan bank mitra dan sistem internal.

**[00:02:10] Interviewer:**  
Terkait regulasi dan *compliance*, bagaimana koordinasi dengan pihak OJK dan kepatuhan terhadap Undang-Undang Perlindungan Data Pribadi (UU PDP)?

**[00:02:22] Budi Santoso:**  
Sejak awal berdiri, kami masuk ke dalam *regulatory sandbox* OJK. Kepatuhan terhadap *data governance* adalah prioritas nomor satu. Semua data nasabah disimpan di dalam *data center* domestik, dengan sistem audit berkala dari pihak ketiga. Dalam konteks UU PDP, kami menerapkan prinsip *data minimization* dan *explicit consent* untuk setiap data yang diproses.

---

## Industry Terminology Glossary
| Term Spoken | Language | Definition & Industry Context |
|---|---|---|
| *Financial Inclusion* | English | Akses terhadap produk dan layanan keuangan formal yang terjangkau bagi semua lapisan masyarakat |
| *Collateral* | English | Agunan atau jaminan aset yang diserahkan peminjam kepada pemberi pinjaman |
| *Underwriting Engine* | English | Sistem algoritmik untuk menilai tingkat risiko kredit calon debitur |
| *Offline-First* | English | Pendekatan arsitektur perangkat lunak yang memprioritaskan fungsi lokal sebelum sinkronisasi cloud |
| *Data Minimization* | English | Prinsip kepatuhan hukum privasi di mana hanya data esensial yang boleh dikumpulkan |
| *Regulatory Sandbox* | English | Mekanisme uji coba regulator (OJK) untuk inovasi teknologi finansial sebelum izin penuh |
