/* ==========================================================
   rumus • — Master JavaScript (Vanilla JS)
   
   Features Integrated:
   1. Real-time Sticky Search Filter in Left Sidebar
   2. Persistent Dark Mode Toggle with localStorage
   3. Copy-to-Clipboard Functionality for Every Code Block
   4. Live Google Sheets Interactive Simulation Iframe & CTA
   5. Dynamic Module Switcher (ACCOUNTING vs HR - 30 Rumus)
   6. Dynamic Canvas Animated Favicon
   ========================================================== */

/* ==========================================================
   DATABASE 30 RUMUS EXCEL LENGKAP (ACCOUNTING & HR)
   ========================================================== */
const DATABASE = {
  accounting: [
    {
      id: "acc-1",
      number: "01",
      title: "SUMIF & SUMIFS",
      tag: "Rekapitulasi Bersyarat",
      difficulty: "Dasar – Menengah",
      purpose: "Menjumlahkan nilai dalam rentang data tertentu yang memenuhi satu kriteria (SUMIF) atau banyak kriteria sekaligus (SUMIFS). Sangat esensial dalam akuntansi untuk merekap total beban operasional, pendapatan per divisi, atau pengeluaran per vendor.",
      syntax: [
        {
          label: "Sintaks SUMIF (1 Kriteria)",
          code: '=SUMIF(range_kriteria; "kriteria"; sum_range)',
          params: "range_kriteria: Kolom data yang dicek; kriteria: Syarat yang dicari; sum_range: Kolom angka yang dijumlahkan."
        },
        {
          label: "Sintaks SUMIFS (Banyak Kriteria)",
          code: '=SUMIFS(sum_range; kriteria_range1; "kriteria1"; [kriteria_range2; "kriteria2"])',
          params: "sum_range diletakkan di paling depan, disusul pasangan rentang kriteria dan syaratnya."
        }
      ],
      example: {
        scenario: "Akuntan ingin merekap total pengeluaran untuk kategori 'Operasional' pada bulan Oktober 2026.",
        tableHeaders: ["No. Bukti", "Tanggal", "Kategori", "Deskripsi", "Nominal (Rp)"],
        tableRows: [
          ["BK-101", "02/10/2026", "Operasional", "Listrik & Air Kantor", "3.200.000"],
          ["BK-102", "05/10/2026", "Marketing", "Iklan Facebook Ads", "5.000.000"],
          ["BK-103", "12/10/2026", "Operasional", "Internet Fiber Optic", "1.800.000"],
          ["BK-104", "20/10/2026", "Gaji", "Gaji Staf Magang", "4.500.000"]
        ],
        formulaUsed: '=SUMIF(C2:C5; "Operasional"; E2:E5)',
        result: "Rp 5.000.000",
        explanation: "Excel mencari kata 'Operasional' di kolom C (baris 1 & 3), kemudian menjumlahkan angka di baris yang cocok pada kolom E: 3.200.000 + 1.800.000 = 5.000.000."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-01/pubhtml?widget=true&amp;headers=false",
      proTip: "Pada SUMIF, kolom yang dijumlahkan berada di belakang. Namun pada SUMIFS, kolom angka yang dijumlahkan harus selalu berada di urutan PERTAMA."
    },
    {
      id: "acc-2",
      number: "02",
      title: "VLOOKUP & HLOOKUP",
      tag: "Pencarian Referensi",
      difficulty: "Dasar – Menengah",
      purpose: "Mencari data pada kolom pertama tabel referensi dan mengambil nilai dari kolom lain pada baris yang sama secara vertikal (VLOOKUP) atau horizontal (HLOOKUP). Dipakai untuk mengisi nama akun dari kode COA, harga satuan dari master barang, dsb.",
      syntax: [
        {
          label: "Sintaks VLOOKUP",
          code: '=VLOOKUP(lookup_value; table_array; col_index_num; [range_lookup])',
          params: "lookup_value: Kunci pencarian; table_array: Tabel acuan; col_index_num: Nomor kolom yang ingin diambil; range_lookup: FALSE untuk pencarian tepat (exact match)."
        },
        {
          label: "Sintaks HLOOKUP",
          code: '=HLOOKUP(lookup_value; table_array; row_index_num; [range_lookup])',
          params: "Mirip VLOOKUP namun mencari nilai kunci pada baris pertama dan mengambil data dari baris ke-N."
        }
      ],
      example: {
        scenario: "Menampilkan Nama Akun secara otomatis berdasarkan Kode COA pada jurnal voucher.",
        tableHeaders: ["Kode Akun", "Nama Akun (Master)", "Tipe Saldo Normal"],
        tableRows: [
          ["1101", "Kas Operasional", "Debit"],
          ["1102", "Bank Mandiri Operasional", "Debit"],
          ["2101", "Hutang Usaha Dagang", "Kredit"],
          ["5101", "Beban Sewa Gedung", "Debit"]
        ],
        formulaUsed: '=VLOOKUP("1102"; A2:C5; 2; FALSE)',
        result: '"Bank Mandiri Operasional"',
        explanation: "VLOOKUP mencari kode '1102' pada kolom pertama tabel (A2:A5), lalu mengambil nilai dari kolom ke-2 yaitu Nama Akun."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-02/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan selalu FALSE atau 0 di parameter terakhir agar Excel mencari nilai yang persis sama, bukan perkiraan."
    },
    {
      id: "acc-3",
      number: "03",
      title: "INDEX & MATCH",
      tag: "Lookup Lanjutan Dua Arah",
      difficulty: "Menengah – Mahir",
      purpose: "Kombinasi paling fleksibel pengganti VLOOKUP. Mampu mencari data ke arah kiri, tidak terpengaruh penyisipan kolom baru, dan mendukung pencarian 2 dimensi (baris dan kolom sekaligus). Wajib dikuasai oleh financial modeler.",
      syntax: [
        {
          label: "Sintaks Kombinasi INDEX & MATCH",
          code: '=INDEX(array_hasil; MATCH(lookup_value; array_kunci; 0))',
          params: "INDEX mengembalikan isi sel berdasarkan posisi baris/kolom; MATCH bertugas mencari nomor urut baris di mana data kunci berada."
        }
      ],
      example: {
        scenario: "Mencari Kode Akun yang terletak di sebelah KIRI Nama Akun (tidak bisa dilakukan VLOOKUP standar tanpa modifikasi).",
        tableHeaders: ["Kode COA (Kiri)", "Nama Akun (Kanan)"],
        tableRows: [
          ["1101", "Kas Kecil"],
          ["1201", "Piutang Usaha"],
          ["5201", "Biaya Transportasi"]
        ],
        formulaUsed: '=INDEX(A2:A4; MATCH("Piutang Usaha"; B2:B4; 0))',
        result: '"1201"',
        explanation: "MATCH mencari 'Piutang Usaha' di B2:B4 dan menemukan posisi baris ke-2. Kemudian INDEX mengambil baris ke-2 dari rentang A2:A4."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-03/pubhtml?widget=true&amp;headers=false",
      proTip: "Formula ini jauh lebih cepat (ringan memory) dibandingkan VLOOKUP pada dataset akuntansi yang mencapai puluhan ribu baris."
    },
    {
      id: "acc-4",
      number: "04",
      title: "PMT",
      tag: "Amortisasi & Kredit",
      difficulty: "Menengah",
      purpose: "Menghitung nilai cicilan atau angsuran pinjaman berkala dengan suku bunga tetap dan pembayaran konstan. Digunakan untuk simulasi kredit kendaraan operasional, pinjaman modal bank, atau cicilan KPR kantor.",
      syntax: [
        {
          label: "Sintaks PMT",
          code: '=PMT(rate; nper; pv; [fv]; [type])',
          params: "rate: Bunga per periode (jika tahunan, bagi 12); nper: Total periode pembayaran (bulan); pv: Nilai pinjaman sekarang (diberi tanda minus agar hasil cicilan positif)."
        }
      ],
      example: {
        scenario: "Perusahaan meminjam modal Rp 240.000.000 dengan bunga 9% per tahun selama 36 bulan (3 tahun).",
        tableHeaders: ["Plafon Pinjaman (PV)", "Bunga / Tahun", "Tenor", "Rumus Angsuran / Bulan"],
        tableRows: [
          ["Rp 240.000.000", "9%", "36 Bulan", "=PMT(9%/12; 36; -240000000)"]
        ],
        formulaUsed: '=PMT(9%/12; 36; -240000000)',
        result: "Rp 7.631.905",
        explanation: "Tiap bulan perusahaan wajib mengalokasikan beban cicilan tetap sebesar Rp 7.631.905 untuk pokok dan bunga."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-04/pubhtml?widget=true&amp;headers=false",
      proTip: "Pastikan satuan rate dan nper sinkron. Jika cicilan bulanan, bagi bunga tahunan dengan 12."
    },
    {
      id: "acc-5",
      number: "05",
      title: "PV & FV",
      tag: "Nilai Waktu Uang (TVM)",
      difficulty: "Menengah",
      purpose: "PV (Present Value) menghitung nilai saat ini dari arus kas masa depan, sedangkan FV (Future Value) menghitung nilai investasi di masa depan dengan asumsi tingkat bunga tertentu dan setoran rutin.",
      syntax: [
        {
          label: "Sintaks PV",
          code: '=PV(rate; nper; pmt; [fv]; [type])',
          params: "Menghitung modal awal yang dibutuhkan saat ini untuk mencapai nilai tertentu di masa mendatang."
        },
        {
          label: "Sintaks FV",
          code: '=FV(rate; nper; pmt; [pv]; [type])',
          params: "Menghitung akumulasi dana di masa depan jika rutin menabung sejumlah uang tiap bulan."
        }
      ],
      example: {
        scenario: "Menabung dana cadangan pajak sebesar Rp 5.000.000 tiap bulan selama 12 bulan di reksa dana pendapatan tetap berimbal hasil 6% per tahun.",
        tableHeaders: ["Setoran Bulanan (PMT)", "Tingkat Return", "Jangka Waktu", "Estimasi Total Tabungan"],
        tableRows: [
          ["Rp 5.000.000", "6% / tahun", "12 Bulan", '=FV(6%/12; 12; -5000000)']
        ],
        formulaUsed: '=FV(6%/12; 12; -5000000)',
        result: "Rp 61.677.812",
        explanation: "Total dana terkumpul adalah Rp 61.677.812 (sudah termasuk pokok tabungan Rp 60.000.000 dan bunga imbal hasil)."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-05/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan tanda negatif (-) pada setoran agar hasil perhitungan menghasilkan angka bernilai positif."
    },
    {
      id: "acc-6",
      number: "06",
      title: "NPV",
      tag: "Studi Kelayakan Investasi",
      difficulty: "Menengah – Mahir",
      purpose: "Net Present Value (NPV) menghitung selisih antara nilai sekarang dari arus kas masuk masa depan dengan modal investasi awal. Digunakan manajemen untuk memutuskan apakah sebuah proyek layak dibiayai atau ditolak.",
      syntax: [
        {
          label: "Sintaks NPV",
          code: '=NPV(discount_rate; nilai_arus_kas_1; nilai_arus_kas_2; ...) + modal_awal',
          params: "discount_rate: Suku bunga diskonto (WACC); modal_awal bertanda negatif (-) dan ditambahkan di LUAR kurung formula NPV."
        }
      ],
      example: {
        scenario: "Membeli mesin cetak komersial seharga Rp 50.000.000 yang menghasilkan arus kas Rp 18.000.000/tahun selama 4 tahun dengan diskonto 10%.",
        tableHeaders: ["Tahun ke-0 (Modal)", "Tahun 1", "Tahun 2", "Tahun 3", "Tahun 4"],
        tableRows: [
          ["-Rp 50.000.000", "Rp 18.000.000", "Rp 18.000.000", "Rp 18.000.000", "Rp 18.000.000"]
        ],
        formulaUsed: '=NPV(10%; B2:E2) + A2',
        result: "Rp 7.057.652 (Layak / Profit)",
        explanation: "Karena hasil NPV bernilai positif (> 0), maka proyek pengadaan mesin ini layak dijalankan karena memberi nilai tambah perusahaan."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-06/pubhtml?widget=true&amp;headers=false",
      proTip: "Hati-hati: Dalam Excel, fungsi NPV mengasumsikan arus kas pertama terjadi di akhir tahun pertama. Jangan masukkan investasi tahun 0 ke dalam kurung NPV."
    },
    {
      id: "acc-7",
      number: "07",
      title: "IRR",
      tag: "Tingkat Pengembalian Internal",
      difficulty: "Mahir",
      purpose: "Internal Rate of Return (IRR) menghitung tingkat diskonto yang membuat nilai NPV dari seluruh arus kas sama dengan nol. Menjadi indikator persentase imbal hasil riil sebuah proyek.",
      syntax: [
        {
          label: "Sintaks IRR",
          code: '=IRR(values; [guess])',
          params: "values: Rentang sel yang memuat arus kas keluar (negatif) dan arus kas masuk (positif) secara berurutan waktu."
        }
      ],
      example: {
        scenario: "Investasi buka cabang baru: Modal Rp 100 Juta, menghasilkan laba Rp 30 Juta, Rp 40 Juta, Rp 50 Juta pada 3 tahun pertama.",
        tableHeaders: ["Tahun 0", "Tahun 1", "Tahun 2", "Tahun 3", "Hasil IRR"],
        tableRows: [
          ["-100.000.000", "30.000.000", "40.000.000", "50.000.000", '=IRR(A2:D2)']
        ],
        formulaUsed: '=IRR(A2:D2)',
        result: "8.89%",
        explanation: "Tingkat pengembalian proyek adalah 8.89%. Jika suku bunga pinjaman modal hanya 7%, maka proyek ini menguntungkan."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-07/pubhtml?widget=true&amp;headers=false",
      proTip: "Pastikan setidaknya ada satu nilai negatif (modal keluar) dan satu nilai positif di rentang data agar rumus tidak menghasilkan error #NUM!."
    },
    {
      id: "acc-8",
      number: "08",
      title: "SLN & SYD",
      tag: "Depresiasi Aset Tetap",
      difficulty: "Menengah",
      purpose: "Menghitung beban penyusutan aktiva tetap. SLN (Straight-Line) untuk metode garis lurus (beban sama rata tiap tahun), sedangkan SYD (Sum-of-Years' Digits) untuk metode penyusutan dipercepat.",
      syntax: [
        {
          label: "Sintaks SLN (Garis Lurus)",
          code: '=SLN(cost; salvage; life)',
          params: "cost: Harga perolehan aset; salvage: Nilai sisa/residu; life: Umur ekonomis (dalam tahun/bulan)."
        },
        {
          label: "Sintaks SYD (Metode Angka Tahun)",
          code: '=SYD(cost; salvage; life; per)',
          params: "per: Periode tahun yang ingin dihitung penyusutannya (misal tahun ke-1)."
        }
      ],
      example: {
        scenario: "Kendaraan operasional dibeli seharga Rp 200.000.000, taksiran nilai sisa Rp 20.000.000 dengan masa manfaat 5 tahun.",
        tableHeaders: ["Harga Beli (Cost)", "Nilai Sisa", "Masa Manfaat", "Penyusutan/Tahun (SLN)"],
        tableRows: [
          ["Rp 200.000.000", "Rp 20.000.000", "5 Tahun", "=SLN(200jt; 20jt; 5)"]
        ],
        formulaUsed: '=SLN(200000000; 20000000; 5)',
        result: "Rp 36.000.000 per tahun",
        explanation: "(200.000.000 - 20.000.000) / 5 = 36.000.000 beban penyusutan yang dicatat pada jurnal akuntansi tahunan."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-08/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan SLN untuk laporan pajak standar PSAK Indonesia jika perusahaan Anda menerapkan amortisasi garis lurus."
    },
    {
      id: "acc-9",
      number: "09",
      title: "EOMONTH",
      tag: "Jatuh Tempo Akhir Bulan",
      difficulty: "Dasar",
      purpose: "End of Month (EOMONTH) menghasilkan tanggal hari terakhir pada bulan berjalan, bulan lampau, atau sekian bulan ke depan. Sangat berguna untuk menjadwalkan batas pelaporan pajak (SPT) dan cut-off closing bulanan.",
      syntax: [
        {
          label: "Sintaks EOMONTH",
          code: '=EOMONTH(start_date; months)',
          params: "start_date: Tanggal awal; months: Jumlah bulan (0 = akhir bulan ini, 1 = akhir bulan depan, -1 = akhir bulan lalu)."
        }
      ],
      example: {
        scenario: "Invoice diterbitkan tanggal 15 Februari 2026. Pembayaran jatuh tempo pada hari terakhir bulan berikutnya.",
        tableHeaders: ["Tgl Invoice", "Parameter Bulan", "Formula", "Tgl Jatuh Tempo"],
        tableRows: [
          ["15/02/2026", "1 (Bulan Depan)", "=EOMONTH(A2; 1)", "31/03/2026"]
        ],
        formulaUsed: '=EOMONTH(A2; 1)',
        result: "31/03/2026",
        explanation: "Otomatis menghasilkan tanggal 31 Maret tanpa pusing apakah bulan tersebut berumur 30, 31, atau 28/29 hari (kabisat)."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-09/pubhtml?widget=true&amp;headers=false",
      proTip: "Format sel hasil formula sebagai 'Date' agar tidak tampil sebagai angka seri serial number Excel."
    },
    {
      id: "acc-10",
      number: "10",
      title: "EDATE",
      tag: "Jatuh Tempo Periode",
      difficulty: "Dasar",
      purpose: "Menghasilkan tanggal yang jatuh persis sekian bulan sebelum atau sesudah tanggal acuan pada hari yang sama. Cocok untuk menghitung jatuh tempo tagihan sewa berkala atau masa garansi aset.",
      syntax: [
        {
          label: "Sintaks EDATE",
          code: '=EDATE(start_date; months)',
          params: "months: Angka positif untuk bulan ke depan, negatif untuk bulan ke belakang."
        }
      ],
      example: {
        scenario: "Polis asuransi gedung dibayar tanggal 10 Oktober 2026 dan berlaku selama 6 bulan ke depan.",
        tableHeaders: ["Tgl Mulai Polis", "Durasi", "Formula", "Tgl Berakhir"],
        tableRows: [
          ["10/10/2026", "6 Bulan", "=EDATE(A2; 6)", "10/04/2027"]
        ],
        formulaUsed: '=EDATE(A2; 6)',
        result: "10/04/2027",
        explanation: "Secara presisi melompat 6 bulan ke depan tepat di tanggal 10."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-10/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan untuk kontrak berjangka (term of payment) bertempo 'Net 30/60/90 days' atau 'X months'."
    },
    {
      id: "acc-11",
      number: "11",
      title: "SUBTOTAL",
      tag: "Otomatisasi Tabel Filter",
      difficulty: "Menengah",
      purpose: "Menghitung subtotal dari daftar data yang sedang difilter. Keunggulan utamanya: nilai dari baris yang disembunyikan/di-filter TIDAK AKAN dihitung, mencegah duplikasi hitungan grand total.",
      syntax: [
        {
          label: "Sintaks SUBTOTAL",
          code: '=SUBTOTAL(function_num; ref1; [ref2]; ...)',
          params: "function_num: Kode rumus (9 = SUM data tampak, 109 = SUM abaikan baris yang di-hide manual; 1 = AVERAGE; 3 = COUNTA)."
        }
      ],
      example: {
        scenario: "Menjumlahkan kolom total biaya hanya untuk baris invoice yang sedang aktif difilter.",
        tableHeaders: ["Kategori", "Vendor", "Biaya", "Keterangan Filter"],
        tableRows: [
          ["Logistik", "JNE Express", "1.500.000", "Ditampilkan"],
          ["Logistik", "SiCepat", "2.000.000", "Ditampilkan"],
          ["IT Support", "Server AWS", "4.000.000", "Disembunyikan (Filter Logistik)"]
        ],
        formulaUsed: '=SUBTOTAL(9; C2:C4)',
        result: "Rp 3.500.000",
        explanation: "Baris IT Support diabaikan karena tersembunyi oleh filter. Hasilnya hanya 1.500.000 + 2.000.000 = 3.500.000."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-11/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan kode 9 atau 109 di lembar kerja buku kas besar agar angka ringkasan selalu sinkron dengan dropdown filter."
    },
    {
      id: "acc-12",
      number: "12",
      title: "IF, AND, OR",
      tag: "Logika Keputusan Keuangan",
      difficulty: "Dasar – Menengah",
      purpose: "Membuat uji logika bersyarat dalam laporan keuangan. Misalnya menentukan status audit 'Balance' atau 'Selisih', menghitung tarif pajak progresif, atau validasi approval pengeluaran dana.",
      syntax: [
        {
          label: "Sintaks Logika Bersarang",
          code: '=IF(kondisi; hasil_jika_benar; hasil_jika_salah)',
          params: "Bisa dikombinasikan dengan AND(kondisi1; kondisi2) untuk syarat ganda, atau OR(kondisi1; kondisi2) untuk salah satu syarat."
        }
      ],
      example: {
        scenario: "Memeriksa apakah Neraca Saldo sudah seimbang antara Total Debit dan Total Kredit.",
        tableHeaders: ["Total Debit", "Total Kredit", "Formula Pengecekan", "Status"],
        tableRows: [
          ["500.000.000", "500.000.000", '=IF(A2=B2; "BALANCE"; "SELISIH!")', "BALANCE"],
          ["250.000.000", "248.500.000", '=IF(A3=B3; "BALANCE"; "SELISIH!")', "SELISIH!"]
        ],
        formulaUsed: '=IF(A2=B2; "BALANCE"; "SELISIH Rp " & TEXT(A2-B2; "#,##0"))',
        result: '"BALANCE"',
        explanation: "Jika Debit = Kredit, muncul pesan 'BALANCE', bila berbeda akan menampilkan peringatan adanya selisih pembukuan."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-12/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan kombinasi AND untuk memvalidasi syarat pencairan kas bon: plafon <= limit AND ada paraf atasan."
    },
    {
      id: "acc-13",
      number: "13",
      title: "IFERROR",
      tag: "Pencegah Error Laporan",
      difficulty: "Dasar",
      purpose: "Menangkap dan menyembunyikan pesan kesalahan seperti #N/A, #DIV/0!, #VALUE!, atau #REF!. Laporan laba rugi atau dashboard keuangan menjadi bersih tanpa teks error yang merusak presentasi.",
      syntax: [
        {
          label: "Sintaks IFERROR",
          code: '=IFERROR(rumus_asli; nilai_jika_error)',
          params: "rumus_asli: Formula yang ingin dieksekusi; nilai_jika_error: Nilai pengganti jika terjadi kesalahan (misal 0, strip \"-\", atau teks kosongan \"\")."
        }
      ],
      example: {
        scenario: "Menghitung rasio pertumbuhan laba: (Laba Th Ini - Laba Th Lalu) / Laba Th Lalu, ketika data tahun lalu bernilai 0.",
        tableHeaders: ["Laba Th Ini", "Laba Th Lalu", "Formula Standar", "Dengan IFERROR"],
        tableRows: [
          ["50.000.000", "0", "#DIV/0!", '=IFERROR((A2-B2)/B2; 0)']
        ],
        formulaUsed: '=IFERROR((A2-B2)/B2; 0)',
        result: "0%",
        explanation: "Pembagian dengan angka 0 yang biasanya menghasilkan error fatal digantikan menjadi angka 0 yang rapi."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-13/pubhtml?widget=true&amp;headers=false",
      proTip: "Selalu bungkus rumus VLOOKUP atau formula rasio keuangan dengan IFERROR(formula; 0) atau IFERROR(formula; \"-\")."
    },
    {
      id: "acc-14",
      number: "14",
      title: "ROUND, ROUNDUP, ROUNDDOWN",
      tag: "Pembulatan Angka Pajak",
      difficulty: "Dasar",
      purpose: "Membulatkan angka desimal ke digit tertentu sesuai standar pembukuan. ROUND membulatkan ke kelipatan terdekat, ROUNDUP membulatkan ke atas (cocok untuk tagihan), ROUNDDOWN membulatkan ke bawah (aturan potongan pajak PPh).",
      syntax: [
        {
          label: "Sintaks Pembulatan",
          code: '=ROUND(angka; jumlah_digit)',
          params: "jumlah_digit: 0 untuk satuan bulat, -2 untuk ratusan, -3 untuk ribuan terdekat."
        }
      ],
      example: {
        scenario: "Menghitung PPN 11% dari barang seharga Rp 125.455 yang menghasilkan desimal Rp 13.800,05.",
        tableHeaders: ["Dasar Pengenaan Pajak", "Tarif", "Hasil Mentah", "Formula Pembulatan"],
        tableRows: [
          ["Rp 125.455", "11%", "Rp 13.800,05", '=ROUND(A2*11%; 0)']
        ],
        formulaUsed: '=ROUND(125455 * 11%; 0)',
        result: "Rp 13.800",
        explanation: "Desimal ,05 dibulatkan ke bawah ke bilangan rupiah bulat terdekat."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-14/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan parameter negatif -3 untuk membulatkan harga jual produk ke ribuan rupiah terdekat (misal: 14.720 -> 15.000)."
    },
    {
      id: "acc-15",
      number: "15",
      title: "COUNTIFS",
      tag: "Frekuensi Transaksi",
      difficulty: "Dasar – Menengah",
      purpose: "Menghitung jumlah kemunculan baris data transaksi yang memenuhi kriteria ganda. Digunakan untuk menghitung berapa kali invoice macet/tertunggak, frekuensi pembelian pelanggan VIP, dsb.",
      syntax: [
        {
          label: "Sintaks COUNTIFS",
          code: '=COUNTIFS(kriteria_range1; "kriteria1"; [kriteria_range2; "kriteria2"])',
          params: "Mengevaluasi baris di mana semua kondisi terpenuhi sekaligus."
        }
      ],
      example: {
        scenario: "Menghitung berapa banyak invoice dari departemen 'Marketing' yang nominalnya di atas Rp 10.000.000.",
        tableHeaders: ["No. Invoice", "Departemen", "Nominal", "Status"],
        tableRows: [
          ["INV-01", "Marketing", "15.000.000", "Memenuhi"],
          ["INV-02", "Marketing", "4.000.000", "Tidak Memenuhi"],
          ["INV-03", "Finance", "20.000.000", "Tidak Memenuhi (Bukan Marketing)"],
          ["INV-04", "Marketing", "12.500.000", "Memenuhi"]
        ],
        formulaUsed: '=COUNTIFS(B2:B5; "Marketing"; C2:C5; ">10000000")',
        result: "2 Transaksi",
        explanation: "Hanya invoice INV-01 dan INV-04 yang memenuhi kedua syarat tersebut secara serentak."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-accounting-15/pubhtml?widget=true&amp;headers=false",
      proTip: "Tanda perbandingan matematika seperti \">\", \"<\", atau \">=\" harus diapit tanda kutip ganda (\">10000000\")."
    }
  ],

  hr: [
    {
      id: "hr-1",
      number: "01",
      title: "DATEDIF",
      tag: "Masa Kerja & Usia",
      difficulty: "Menengah",
      purpose: "Menghitung selisih antara dua tanggal dalam satuan tahun ('Y'), bulan ('M'), hari ('D'), atau sisa bulan di luar tahun ('YM'). Sangat krusial untuk menghitung usia pensiun, hak cuti tahunan, dan masa kerja untuk pesangon.",
      syntax: [
        {
          label: "Sintaks DATEDIF",
          code: '=DATEDIF(start_date; end_date; "satuan")',
          params: 'start_date: Tanggal awal (tgl bergabung); end_date: Tanggal akhir (misal TODAY()); satuan: "Y" untuk tahun penuh, "YM" sisa bulan, "MD" sisa hari.'
        }
      ],
      example: {
        scenario: "Menghitung masa kerja karyawan yang mulai bergabung tanggal 15 Januari 2020 dihitung per tanggal hari ini.",
        tableHeaders: ["Nama Karyawan", "Tanggal Masuk", "Tanggal Evaluasi", "Formula"],
        tableRows: [
          ["Faari Akbar", "15/01/2020", "23/10/2026", '=DATEDIF(B2; C2; "Y") & " Th " & DATEDIF(B2; C2; "YM") & " Bln"']
        ],
        formulaUsed: '=DATEDIF("2020-01-15"; "2026-10-23"; "Y") & " Tahun " & DATEDIF("2020-01-15"; "2026-10-23"; "YM") & " Bulan"',
        result: '"6 Tahun 9 Bulan"',
        explanation: "DATEDIF menghitung selisih tahun secara presisi tanpa pembulatan hari sembarangan."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-01/pubhtml?widget=true&amp;headers=false",
      proTip: "DATEDIF adalah rumus rahasia bawaan Lotus 1-2-3 yang tidak muncul otomatis di autocomplete Excel, namun 100% berfungsi."
    },
    {
      id: "hr-2",
      number: "02",
      title: "NETWORKDAYS",
      tag: "Hari Kerja Efektif",
      difficulty: "Dasar – Menengah",
      purpose: "Menghitung jumlah hari kerja bersih antara dua tanggal dengan otomatis mengabaikan hari Sabtu dan Minggu (weekend) serta daftar tanggal cuti bersama / libur nasional.",
      syntax: [
        {
          label: "Sintaks NETWORKDAYS",
          code: '=NETWORKDAYS(start_date; end_date; [holidays])',
          params: "holidays: Rentang sel opsional yang berisi daftar tanggal libur nasional Indonesia."
        }
      ],
      example: {
        scenario: "Menghitung hari kerja aktif bulan Oktober 2026 (01/10/2026 s/d 31/10/2026) dengan 2 hari libur tanggal 01/10 dan 23/10.",
        tableHeaders: ["Tanggal Awal", "Tanggal Akhir", "Daftar Libur", "Hari Kerja Bersih"],
        tableRows: [
          ["01/10/2026", "31/10/2026", "H2:H3 (2 Hari)", "=NETWORKDAYS(A2; B2; H2:H3)"]
        ],
        formulaUsed: '=NETWORKDAYS("2026-10-01"; "2026-10-31"; H2:H3)',
        result: "20 Hari Kerja Aktif",
        explanation: "Dari total 31 hari kalender, dikurangi 9 hari akhir pekan (Sabtu-Minggu) dan 2 hari libur nasional."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-02/pubhtml?widget=true&amp;headers=false",
      proTip: "Simpan daftar libur SKB 3 Menteri dalam satu sheet master terpisah bernama 'Master_Libur' agar mudah dijadikan referensi."
    },
    {
      id: "hr-3",
      number: "03",
      title: "NETWORKDAYS.INTL",
      tag: "Hari Kerja Fleksibel / Shift",
      difficulty: "Menengah",
      purpose: "Versi upgrade dari NETWORKDAYS untuk perusahaan dengan sistem kerja selain Senin-Jumat. Misalnya ritel, pabrik, atau RS di mana hari liburnya adalah Minggu saja, atau Jumat-Sabtu.",
      syntax: [
        {
          label: "Sintaks NETWORKDAYS.INTL",
          code: '=NETWORKDAYS.INTL(start_date; end_date; [weekend]; [holidays])',
          params: "weekend: Kode angka untuk pola libur (11 = Hanya hari Minggu saja libur, 1 = Sabtu & Minggu, 7 = Jumat & Sabtu)."
        }
      ],
      example: {
        scenario: "Karyawan toko ritel bekerja 6 hari seminggu (hanya libur di hari Minggu).",
        tableHeaders: ["Mulai", "Selesai", "Pola Libur", "Hasil Hari Kerja"],
        tableRows: [
          ["01/10/2026", "31/10/2026", "11 (Hanya Hari Minggu)", "=NETWORKDAYS.INTL(A2; B2; 11; H2:H3)"]
        ],
        formulaUsed: '=NETWORKDAYS.INTL("2026-10-01"; "2026-10-31"; 11; H2:H3)',
        result: "25 Hari Kerja",
        explanation: "Hari Sabtu dihitung sebagai hari kerja aktif normal karena parameter weekend diset 11 (Sunday only)."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-03/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan string biner '0000001' jika ingin mengkustomisasi pola hari kerja shift khusus."
    },
    {
      id: "hr-4",
      number: "04",
      title: "WORKDAY",
      tag: "Jadwal Probation & Cuti",
      difficulty: "Dasar – Menengah",
      purpose: "Menghitung tanggal selesai setelah sekian hari kerja aktif dari tanggal mulai, otomatis melewati akhir pekan dan hari libur. Digunakan untuk menentukan tanggal akhir masa probation (percobaan) atau SLA penanganan klaim.",
      syntax: [
        {
          label: "Sintaks WORKDAY",
          code: '=WORKDAY(start_date; days; [holidays])',
          params: "days: Jumlah hari kerja aktif (misal 90 hari probation)."
        }
      ],
      example: {
        scenario: "Karyawan magang menandatangani kontrak 45 hari kerja aktif dimulai dari 05 Oktober 2026.",
        tableHeaders: ["Tanggal Mulai", "Lama Hari Kerja", "Formula", "Tanggal Berakhir"],
        tableRows: [
          ["05/10/2026", "45 Hari Kerja", "=WORKDAY(A2; 45; H2:H5)", "08/12/2026"]
        ],
        formulaUsed: '=WORKDAY("2026-10-05"; 45; H2:H5)',
        result: "08/12/2026",
        explanation: "Excel menghitung maju 45 hari kalender kerja tanpa menghitung Sabtu, Minggu, dan hari libur nasional."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-04/pubhtml?widget=true&amp;headers=false",
      proTip: "Sama seperti NETWORKDAYS, ada juga varian WORKDAY.INTL jika perusahaan menerapkan libur non-standar."
    },
    {
      id: "hr-5",
      number: "05",
      title: "COUNTIF & COUNTIFS",
      tag: "Rekap Absensi & Kehadiran",
      difficulty: "Dasar",
      purpose: "Merekap total kehadiran karyawan dari sheet absensi bulanan. Menghitung berapa kali hadir ('H'), izin ('I'), sakit ('S'), cuti ('C'), atau mangkir ('A').",
      syntax: [
        {
          label: "Sintaks Rekap Absen",
          code: '=COUNTIF(range_kehadiran; "H")',
          params: 'range_kehadiran: Kolom tanggal 1 s/d 31 karyawan; "H": Simbol status kehadiran yang dihitung.'
        }
      ],
      example: {
        scenario: "Merekap total ketidakhadiran karyawan (Sakit) dalam rentang 30 hari kerja.",
        tableHeaders: ["ID", "Nama Karyawan", "Rentang Catatan Absen", "Total Sakit"],
        tableRows: [
          ["EMP-01", "Budi Santoso", "H, H, S, H, H, H, S, H ...", '=COUNTIF(C2:AG2; "S")']
        ],
        formulaUsed: '=COUNTIF(C2:AG2; "S")',
        result: "2 Hari Sakit",
        explanation: "Excel menghitung secara instan jumlah huruf 'S' pada baris absensi karyawan tersebut."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-05/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan COUNTIFS untuk menghitung karyawan di Divisi HR yang cuti lebih dari 3 hari dalam sebulan."
    },
    {
      id: "hr-6",
      number: "06",
      title: "TRIM",
      tag: "Pembersihan Data Karyawan",
      difficulty: "Dasar",
      purpose: "Menghapus semua spasi ganda, spasi di awal kalimat, dan spasi tak terlihat di akhir kalimat dari input data mentah formulir HR, hanya menyisakan satu spasi standar antar kata.",
      syntax: [
        {
          label: "Sintaks TRIM",
          code: '=TRIM(text)',
          params: "text: Sel yang berisi teks berantakan dengan banyak spasi berlebih."
        }
      ],
      example: {
        scenario: "Pendaftar kerja mengetik nama dengan spasi acak: '  muhammad   faari  '.",
        tableHeaders: ["Input Form Mentah", "Formula Pembersih", "Hasil Bersih"],
        tableRows: [
          ['"  muhammad   faari  "', "=TRIM(A2)", '"muhammad faari"']
        ],
        formulaUsed: '=TRIM(A2)',
        result: '"muhammad faari"',
        explanation: "Spasi di depan, belakang, dan spasi rangkap tiga di tengah langsung dipangkas rapi."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-06/pubhtml?widget=true&amp;headers=false",
      proTip: "Kombinasikan selalu dengan PROPER: =PROPER(TRIM(A2)) untuk nama karyawan profesional."
    },
    {
      id: "hr-7",
      number: "07",
      title: "PROPER, UPPER, LOWER",
      tag: "Standarisasi Format Nama",
      difficulty: "Dasar",
      purpose: "PROPER mengubah huruf pertama tiap kata menjadi kapital (Title Case). UPPER mengubah seluruh teks menjadi huruf besar kapital (untuk cetak NIK/NPWP), dan LOWER mengubah menjadi huruf kecil (untuk pembuatan email kantor).",
      syntax: [
        {
          label: "Sintaks Format Teks",
          code: "=PROPER(text)  |  =UPPER(text)  |  =LOWER(text)",
          params: "Bisa langsung mengarah pada sel teks mentah nama atau jabatan karyawan."
        }
      ],
      example: {
        scenario: "Menstandarisasi nama karyawan dan membuat alamat email korporat otomatis.",
        tableHeaders: ["Nama Mentah", "PROPER (Cetak Sertifikat)", "LOWER (Email Kantor)"],
        tableRows: [
          ["aHMAD sUDIRMAN", "Ahmad Sudirman", "ahmad.sudirman@rumus.id"]
        ],
        formulaUsed: '=PROPER(A2)  &  =LOWER(SUBSTITUTE(A2; " "; ".")) & "@rumus.id"',
        result: '"Ahmad Sudirman" & "ahmad.sudirman@rumus.id"',
        explanation: "Huruf kapital acak menjadi rapi, dan pembuatan akun email kantor dapat diotomatisasi seketika."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-07/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan UPPER pada kolom kode jabatan (misal: 'hrd-mgr' -> 'HRD-MGR')."
    },
    {
      id: "hr-8",
      number: "08",
      title: "CONCATENATE & TEXTJOIN",
      tag: "Penggabungan Teks & NIK",
      difficulty: "Dasar – Menengah",
      purpose: "Menggabungkan dua atau lebih string teks menjadi satu. TEXTJOIN (Excel modern) jauh lebih unggul karena memungkinkan penambahan pemisah (delimiter seperti koma atau strip) dan mengabaikan sel kosong.",
      syntax: [
        {
          label: "Sintaks TEXTJOIN (Direkomendasikan)",
          code: '=TEXTJOIN(delimiter; ignore_empty; text1; [text2]; ...)',
          params: 'delimiter: Karakter pemisah antar kata, misal " " (spasi) atau ", "; ignore_empty: TRUE untuk melewati sel kosong.'
        }
      ],
      example: {
        scenario: "Menggabungkan Nama Depan, Nama Belakang, dan Gelar Pendidikan dari 3 kolom terpisah.",
        tableHeaders: ["Nama Depan", "Nama Belakang", "Gelar", "Formula TEXTJOIN"],
        tableRows: [
          ["Muhammad", "Faari", "S.E.", '=TEXTJOIN(" "; TRUE; A2:C2)']
        ],
        formulaUsed: '=TEXTJOIN(" "; TRUE; A2:C2)',
        result: '"Muhammad Faari S.E."',
        explanation: "Ketiga kolom digabung secara elegan dengan pemisah satu spasi."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-08/pubhtml?widget=true&amp;headers=false",
      proTip: "Untuk Excel lama yang belum ada TEXTJOIN, gunakan operator ampersand: =A2 & \" \" & B2."
    },
    {
      id: "hr-9",
      number: "09",
      title: "TODAY & NOW",
      tag: "Tanggal Otomatis Sistem",
      difficulty: "Dasar",
      purpose: "TODAY mengembalikan tanggal hari ini secara dinamis dan selalu ter-update otomatis tiap kali file dibuka. NOW menyertakan tanggal beserta jam dan menit saat ini.",
      syntax: [
        {
          label: "Sintaks Tanggal Otomatis",
          code: "=TODAY()  dan  =NOW()",
          params: "Kedua rumus ini tidak memerlukan argumen di dalam tanda kurung."
        }
      ],
      example: {
        scenario: "Menghitung sisa hari sebelum masa kontrak kerja habis terhadap tanggal hari ini.",
        tableHeaders: ["Nama Karyawan", "Akhir Kontrak", "Tanggal Hari Ini", "Sisa Hari"],
        tableRows: [
          ["Reza Rahardian", "31/12/2026", "=TODAY()", "=B2 - TODAY()"]
        ],
        formulaUsed: '=B2 - TODAY()',
        result: "69 Hari Tersisa",
        explanation: "Tanpa perlu input manual tanggal hari ini, sisa hari kontrak akan otomatis berkurang setiap hari."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-09/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan Conditional Formatting bersama =TODAY() untuk memberi warna merah jika kontrak tersisa < 30 hari."
    },
    {
      id: "hr-10",
      number: "10",
      title: "TEXT",
      tag: "Format Tanggal & Rupiah",
      difficulty: "Menengah",
      purpose: "Mengonversi angka atau tanggal numerik menjadi teks dengan format tampilan khusus (misal nama hari dalam bahasa Indonesia, format tanggal resmi surat tugas, atau format rupiah pada slip gaji).",
      syntax: [
        {
          label: "Sintaks TEXT",
          code: '=TEXT(value; "format_text")',
          params: 'format_text: Pola format, misal "[$-421]dddd, dd mmmm yyyy" untuk nama hari bahasa Indonesia.'
        }
      ],
      example: {
        scenario: "Mencetak tanggal SK pengangkatan resmi: 'Jumat, 23 Oktober 2026' dari data tanggal '23/10/2026'.",
        tableHeaders: ["Tanggal Mentah", "Formula TEXT", "Teks Resmi di Surat"],
        tableRows: [
          ["23/10/2026", '=TEXT(A2; "[$-421]dddd, dd mmmm yyyy")', "Jumat, 23 Oktober 2026"]
        ],
        formulaUsed: '=TEXT(A2; "[$-421]dddd, dd mmmm yyyy")',
        result: '"Jumat, 23 Oktober 2026"',
        explanation: "Kode [$-421] adalah locale ID untuk Bahasa Indonesia, mengubah angka bulan 10 menjadi 'Oktober' dan hari menjadi 'Jumat'."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-10/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan format TEXT(gaji; \"Rp #,##0\") ketika ingin menggabungkan nominal gaji ke dalam template kalimat surat penawaran."
    },
    {
      id: "hr-11",
      number: "11",
      title: "IF & IFS",
      tag: "Golongan Gaji & Evaluasi KPI",
      difficulty: "Dasar – Menengah",
      purpose: "Menentukan grade level karyawan, kategori penilaian kinerja (A, B, C, D), atau besaran persentase bonus tahunan berdasarkan pencapaian target skor KPI.",
      syntax: [
        {
          label: "Sintaks IFS (Excel Modern)",
          code: '=IFS(skor>=90; "Sangat Baik"; skor>=80; "Baik"; skor>=70; "Cukup"; TRUE; "Kurang")',
          params: "IFS mengevaluasi kondisi secara berurutan dari kiri ke kanan tanpa perlu membuat kurung IF bertingkat yang rumit."
        }
      ],
      example: {
        scenario: "Menentukan predikat penilaian kinerja staf HR berdasarkan nilai skor penilaian tahunan.",
        tableHeaders: ["Nama Staf", "Skor KPI", "Formula Penilaian", "Predikat Grade"],
        tableRows: [
          ["Siti Nurhaliza", "86", '=IFS(B2>=85; "Grade A"; B2>=75; "Grade B"; TRUE; "Grade C")', "Grade A"]
        ],
        formulaUsed: '=IFS(B2>=85; "Grade A"; B2>=75; "Grade B"; TRUE; "Grade C")',
        result: '"Grade A"',
        explanation: "Karena skor 86 >= 85, maka langsung diklasifikasikan sebagai Grade A dan berhak menerima bonus 100%."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-11/pubhtml?widget=true&amp;headers=false",
      proTip: "Selalu sertakan parameter 'TRUE; nilai_default' di akhir rumus IFS untuk menangani kondisi yang tidak tercakup."
    },
    {
      id: "hr-12",
      number: "12",
      title: "VLOOKUP (Master Payroll)",
      tag: "Tarik Data Gaji Pokok",
      difficulty: "Dasar – Menengah",
      purpose: "Menghubungkan data transaksi payroll bulanan dengan master data karyawan. Menarik nominal Gaji Pokok, Tunjangan Jabatan, dan Status PTKP hanya dengan mengetik ID/NIK Karyawan.",
      syntax: [
        {
          label: "Sintaks VLOOKUP Master Payroll",
          code: '=VLOOKUP(id_karyawan; tabel_master_karyawan; nomor_kolom_gaji; FALSE)',
          params: "Memastikan nominal gaji terambil akurat dari tabel master tunggal tanpa risiko salah ketik nominal."
        }
      ],
      example: {
        scenario: "Menampilkan Gaji Pokok karyawan secara otomatis pada lembar perhitungan slip gaji bulanan.",
        tableHeaders: ["ID Karyawan", "Nama Karyawan", "Jabatan", "Gaji Pokok (Master)"],
        tableRows: [
          ["EMP-101", "Faari Akbar", "HR Specialist", "Rp 8.500.000"],
          ["EMP-102", "Andi Budi", "Recruiter", "Rp 6.500.000"]
        ],
        formulaUsed: '=VLOOKUP("EMP-101"; A2:D3; 4; FALSE)',
        result: "Rp 8.500.000",
        explanation: "Excel mencari 'EMP-101' di kolom pertama dan langsung mengambil nilai kolom ke-4 yaitu Rp 8.500.000."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-12/pubhtml?widget=true&amp;headers=false",
      proTip: "Kunci tabel master dengan tanda dollar $ (misal: $A$2:$D$100) agar saat rumus di-copy ke bawah, referensi tabel tidak bergeser."
    },
    {
      id: "hr-13",
      number: "13",
      title: "EXACT",
      tag: "Validasi NIK & Duplikasi",
      difficulty: "Dasar",
      purpose: "Membandingkan dua sel teks secara case-sensitive (memperhatikan huruf besar/kecil) untuk memastikan apakah keduanya persis sama 100%. Berguna saat rekonsiliasi NIK KTP, No. BPJS, atau password akun portal.",
      syntax: [
        {
          label: "Sintaks EXACT",
          code: '=EXACT(text1; text2)',
          params: "Menghasilkan TRUE jika identik persis, atau FALSE jika terdapat perbedaan karakter atau spasi sekecil apapun."
        }
      ],
      example: {
        scenario: "Memvalidasi apakah Nomor Induk Kependudukan (NIK) di data BPJS Ketenagakerjaan cocok dengan KTP karyawan.",
        tableHeaders: ["NIK di KTP", "NIK di Data BPJS", "Formula Validasi", "Hasil"],
        tableRows: [
          ["3271021501980001", "3271021501980001", "=EXACT(A2; B2)", "TRUE (Valid)"],
          ["3271021501980002", "3271021501980007", "=EXACT(A3; B3)", "FALSE (Beda!)"]
        ],
        formulaUsed: '=IF(EXACT(A2; B2); "Valid Cocok"; "Cek Ulang!")',
        result: '"Valid Cocok"',
        explanation: "Jika ada selisih satu digit atau ketidaksengajaan spasi, rumus langsung mendeteksinya sebagai FALSE."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-13/pubhtml?widget=true&amp;headers=false",
      proTip: "Formula sama dengan = (A2=B2) tidak sensitif terhadap huruf besar-kecil, sedangkan EXACT(A2; B2) sangat ketat."
    },
    {
      id: "hr-14",
      number: "14",
      title: "ISBLANK",
      tag: "Cek Kelengkapan Berkas",
      difficulty: "Dasar",
      purpose: "Memeriksa apakah suatu sel dalam formulir database pelamar atau berkas karyawan berstatus kosong (blank). Sering dipakai HR untuk mendeteksi karyawan yang belum melampirkan berkas NPWP atau Ijazah.",
      syntax: [
        {
          label: "Sintaks ISBLANK",
          code: '=IF(ISBLANK(sel_dokumen); "Belum Lengkap"; "Lengkap")',
          params: "Menghasilkan TRUE hanya jika sel benar-benar kosong belum terisi data."
        }
      ],
      example: {
        scenario: "Mengecek status kelengkapan nomor NPWP karyawan baru.",
        tableHeaders: ["Nama Karyawan", "Nomor NPWP", "Formula Verifikasi", "Status"],
        tableRows: [
          ["Reza Rahardian", "", '=IF(ISBLANK(B2); "Belum Ada NPWP"; "Lengkap")', "Belum Ada NPWP"],
          ["Budi Santoso", "09.123.456.7-012.000", '=IF(ISBLANK(B3); "Belum Ada NPWP"; "Lengkap")', "Lengkap"]
        ],
        formulaUsed: '=IF(ISBLANK(B2); "Belum Ada NPWP"; "Lengkap")',
        result: '"Belum Ada NPWP"',
        explanation: "Karena sel NPWP Reza kosong, formula langsung memunculkan status 'Belum Ada NPWP' sebagai reminder tim payroll."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-14/pubhtml?widget=true&amp;headers=false",
      proTip: "Jika sel berisi spasi atau teks kosong (\"\"), ISBLANK menganggapnya tidak kosong. Gunakan =IF(OR(ISBLANK(A2); A2=\"\"); ...) untuk penanganan lebih kuat."
    },
    {
      id: "hr-15",
      number: "15",
      title: "IFERROR (Data Karyawan)",
      tag: "Pembersih Error NIK",
      difficulty: "Dasar",
      purpose: "Mencegah munculnya pesan error #N/A saat melakukan VLOOKUP data karyawan yang belum terdaftar di database master, menggantikannya dengan pesan humanis seperti 'Karyawan Baru' atau 'Data Tidak Ditemukan'.",
      syntax: [
        {
          label: "Sintaks IFERROR HR",
          code: '=IFERROR(VLOOKUP(nik; tabel_karyawan; 2; FALSE); "Belum Terdaftar")',
          params: "Menjaga lembar kerja HR tetap bersih dan profesional saat dicetak menjadi laporan PDF."
        }
      ],
      example: {
        scenario: "Mencari nama karyawan berdasarkan input ID: EMP-999 yang ternyata belum terdaftar.",
        tableHeaders: ["ID Input", "Formula Tanpa IFERROR", "Dengan IFERROR"],
        tableRows: [
          ["EMP-999", "#N/A", '=IFERROR(VLOOKUP(A2; Master; 2; 0); "Karyawan Belum Ada")']
        ],
        formulaUsed: '=IFERROR(VLOOKUP("EMP-999"; Master; 2; FALSE); "Karyawan Belum Ada")',
        result: '"Karyawan Belum Ada"',
        explanation: "Pesan error diganti dengan keterangan ramah yang memandu staf admin untuk mendaftarkan karyawan baru tersebut."
      },
      sheetsTemplateUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3WbB13oN23-placeholder-hr-15/pubhtml?widget=true&amp;headers=false",
      proTip: "Gunakan di seluruh dashboard HR yang memiliki input pencarian interaktif."
    }
  ]
};

/* ==========================================================
   GLOBAL STATE VARIABLES
   ========================================================== */
let currentCategory = "accounting"; // 'accounting' atau 'hr'
let activeFormulaIndex = 0;          // 0 - 14

/* ==========================================================
   DOM SELECTOR HELPERS
   ========================================================== */
const $ = (selector, scope = document) => scope.querySelector(selector);
const $$ = (selector, scope = document) => [...scope.querySelectorAll(selector)];

/* ==========================================================
   INITIALIZATION ON DOM LOAD (MULTI-PAGE & URL ROUTING)
   ========================================================== */
document.addEventListener("DOMContentLoaded", () => {
  // 1. Shared features on every page
  try { initTheme(); } catch (e) { console.error("Theme init error:", e); }
  try { initAnimatedFavicon(); } catch (e) { console.error("Favicon init error:", e); }

  // 2. Check which page is currently active
  const isCoursePage = Boolean($("#formulaSidebarList") && $("#formulaDetailContainer"));

  if (isCoursePage) {
    // === COURSE PAGE (DASHBOARD WORKBENCH) ===
    try {
      // Parse URL parameters: ?module=accounting or ?module=hr
      const urlParams = new URLSearchParams(window.location.search);
      const moduleParam = (urlParams.get("module") || "").toLowerCase().trim();

      if (moduleParam === "hr") {
        currentCategory = "hr";
      } else {
        currentCategory = "accounting";
      }

      // Sync active state on navbar module switcher pills
      updateCategoryUI(currentCategory);

      // Render initial sidebar and detail immediately
      renderSidebarList();
      renderFormulaDetail(0);

      // Initialize course interactive controls
      initCategorySwitcher();
      initSidebarSearch();
      initSidebarMobileToggle();
    } catch (err) {
      console.error("Course workbench initialization error:", err);
    }
  } else {
    // === LANDING PAGE (INDEX.HTML) ===
    try {
      initLandingNavbar();
      initPracticeDemo();
    } catch (err) {
      console.error("Landing page initialization error:", err);
    }
  }
});

/* ==========================================================
   FEATURE: PERSISTENT DARK MODE LOGIC (SHARED)
   ========================================================== */
function initTheme() {
  const themeToggleBtn = $("#themeToggleBtn");
  let savedTheme = null;

  try {
    savedTheme = localStorage.getItem("rumus_theme");
  } catch (e) {
    console.warn("localStorage tidak dapat diakses:", e);
  }

  const systemPrefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const isDark = savedTheme === "dark" || (!savedTheme && systemPrefersDark);

  applyTheme(isDark);

  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", () => {
      const currentlyDark = document.body.classList.contains("dark-mode");
      const nextThemeIsDark = !currentlyDark;
      applyTheme(nextThemeIsDark);

      try {
        localStorage.setItem("rumus_theme", nextThemeIsDark ? "dark" : "light");
      } catch (e) {}

      showToast(nextThemeIsDark ? "Mode Gelap diaktifkan 🌙" : "Mode Terang diaktifkan ☀️");
    });
  }
}

function applyTheme(isDark) {
  if (isDark) {
    document.documentElement.classList.add("dark-mode");
    document.body.classList.add("dark-mode");
    const label = $("#themeLabelText");
    if (label) label.textContent = "Mode Terang";
  } else {
    document.documentElement.classList.remove("dark-mode");
    document.body.classList.remove("dark-mode");
    const label = $("#themeLabelText");
    if (label) label.textContent = "Mode Gelap";
  }
}

/* ==========================================================
   LANDING PAGE: NAVBAR & SCROLL HIGHLIGHTING
   ========================================================== */
function initLandingNavbar() {
  const navToggle = $("#navToggle");
  const mainNavMenu = $("#mainNavMenu");

  if (navToggle && mainNavMenu) {
    navToggle.addEventListener("click", () => {
      const isOpen = mainNavMenu.classList.toggle("is-open");
      navToggle.setAttribute("aria-expanded", String(isOpen));
    });
  }

  // Smooth scroll and active state
  $$(".nav-link").forEach((link) => {
    link.addEventListener("click", () => {
      $$(".nav-link").forEach((l) => l.classList.remove("is-active"));
      link.classList.add("is-active");
      if (mainNavMenu) mainNavMenu.classList.remove("is-open");
    });
  });

  window.addEventListener("scroll", () => {
    const sections = ["beranda", "keunggulan", "kursus-pilihan", "tabel-praktik"];
    const scrollPos = window.scrollY + 120;

    for (const secId of sections) {
      const secEl = document.getElementById(secId);
      if (secEl) {
        const top = secEl.offsetTop;
        const height = secEl.offsetHeight;
        if (scrollPos >= top && scrollPos < top + height) {
          $$(".nav-link").forEach((l) => {
            const href = l.getAttribute("href");
            if (href === "#" + secId) {
              l.classList.add("is-active");
            } else {
              l.classList.remove("is-active");
            }
          });
          break;
        }
      }
    }
  }, { passive: true });
}

/* ==========================================================
   COURSE PAGE: MODULE SWITCHER & URL SYNC
   ========================================================== */
function initCategorySwitcher() {
  const catBtnAccounting = $("#catBtnAccounting");
  const catBtnHr = $("#catBtnHr");

  if (catBtnAccounting) {
    catBtnAccounting.addEventListener("click", () => selectCategory("accounting"));
  }
  if (catBtnHr) {
    catBtnHr.addEventListener("click", () => selectCategory("hr"));
  }
}

function selectCategory(cat) {
  if (currentCategory === cat && $("#formulaDetailContainer").children.length > 0) {
    return;
  }
  currentCategory = cat;
  activeFormulaIndex = 0;

  // Update URL without full page reload
  try {
    const url = new URL(window.location);
    url.searchParams.set("module", cat);
    window.history.pushState({}, "", url);
  } catch (e) {}

  updateCategoryUI(cat);

  // Clear search and rerender
  const searchInput = $("#sidebarSearchInput");
  const clearBtn = $("#searchClearBtn");
  if (searchInput) searchInput.value = "";
  if (clearBtn) clearBtn.hidden = true;

  renderSidebarList("");
  renderFormulaDetail(0);

  // Close sidebar drawer on mobile
  const sidebar = $("#dashboardSidebar");
  if (sidebar && window.innerWidth <= 900) {
    sidebar.classList.remove("is-open");
  }
}

function updateCategoryUI(cat) {
  const isAcc = cat === "accounting";

  // Navbar module buttons
  const catBtnAccounting = $("#catBtnAccounting");
  const catBtnHr = $("#catBtnHr");
  if (catBtnAccounting) {
    catBtnAccounting.classList.toggle("is-active", isAcc);
    catBtnAccounting.setAttribute("aria-selected", String(isAcc));
  }
  if (catBtnHr) {
    catBtnHr.classList.toggle("is-active", !isAcc);
    catBtnHr.setAttribute("aria-selected", String(!isAcc));
  }

  // Banner & Sidebar label
  const bannerTag = $("#bannerTagText");
  const bannerTitle = $("#bannerTitle");
  const bannerDesc = $("#bannerDesc");
  const sidebarCategoryLabel = $("#sidebarCategoryLabel");

  if (isAcc) {
    if (bannerTag) bannerTag.textContent = "BIDANG KEUANGAN, AKUNTANSI & PELAPORAN";
    if (bannerTitle) bannerTitle.textContent = "Dashboard Rumus Excel Accounting";
    if (bannerDesc) bannerDesc.textContent = "Panduan lengkap formula finansial, otomatisasi pembukuan, evaluasi proyek, dan pelaporan keuangan dengan simulasi Google Sheets langsung.";
    if (sidebarCategoryLabel) sidebarCategoryLabel.textContent = "MODUL ACCOUNTING";
  } else {
    if (bannerTag) bannerTag.textContent = "BIDANG HUMAN RESOURCES, PAYROLL & DATA KERJA";
    if (bannerTitle) bannerTitle.textContent = "Dashboard Rumus Excel HR (Human Resources)";
    if (bannerDesc) bannerDesc.textContent = "Kumpulan formula esensial untuk manajemen absensi, database karyawan, masa kerja pensiun, dan payroll dengan simulasi Google Sheets langsung.";
    if (sidebarCategoryLabel) sidebarCategoryLabel.textContent = "MODUL HR";
  }
}

/* ==========================================================
   COURSE PAGE: SIDEBAR SEARCH FILTER (REAL-TIME)
   ========================================================== */
function initSidebarSearch() {
  const searchInput = $("#sidebarSearchInput");
  const clearBtn = $("#searchClearBtn");

  if (!searchInput) return;

  searchInput.addEventListener("input", (e) => {
    const query = e.target.value;
    if (clearBtn) {
      clearBtn.hidden = !query;
    }
    renderSidebarList(query);
  });

  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      searchInput.value = "";
      clearBtn.hidden = true;
      searchInput.focus();
      renderSidebarList("");
    });
  }
}

/* ==========================================================
   COURSE PAGE: MOBILE SIDEBAR TOGGLE
   ========================================================== */
function initSidebarMobileToggle() {
  const toggleBtn = $("#sidebarToggle");
  const sidebar = $("#dashboardSidebar");

  if (toggleBtn && sidebar) {
    toggleBtn.addEventListener("click", () => {
      sidebar.classList.toggle("is-open");
    });
  }
}

/* ==========================================================
   RENDER SIDEBAR ITEMS WITH REAL-TIME FILTER
   ========================================================== */
function renderSidebarList(filterText = "") {
  const container = $("#formulaSidebarList");
  if (!container) return;
  container.innerHTML = "";

  const list = DATABASE[currentCategory] || [];
  const query = filterText.toLowerCase().trim();

  let visibleCount = 0;

  list.forEach((item, index) => {
    const matches =
      !query ||
      item.title.toLowerCase().includes(query) ||
      item.tag.toLowerCase().includes(query) ||
      item.number.includes(query) ||
      item.purpose.toLowerCase().includes(query);

    if (matches) {
      visibleCount++;
      const btn = document.createElement("button");
      btn.className = `sidebar__item ${index === activeFormulaIndex ? "is-active" : ""}`;
      btn.setAttribute("role", "tab");
      btn.setAttribute("aria-selected", String(index === activeFormulaIndex));
      btn.dataset.index = index;

      btn.innerHTML = `
        <span class="item-num">${item.number}</span>
        <div class="item-info">
          <span class="item-title">${item.title}</span>
          <span class="item-tag">${item.tag}</span>
        </div>
        <span class="item-arrow" aria-hidden="true">→</span>
      `;

      btn.addEventListener("click", () => {
        selectFormula(index);
        const sidebar = $("#dashboardSidebar");
        if (sidebar && window.innerWidth <= 900) {
          sidebar.classList.remove("is-open");
        }
      });

      container.appendChild(btn);
    }
  });

  if (visibleCount === 0) {
    container.innerHTML = `
      <div class="sidebar-no-results">
        <span>🔍</span>
        <p>Tidak ada rumus dengan kata kunci <strong>"${escapeHTML(filterText)}"</strong></p>
      </div>
    `;
  }

  const counterEl = $("#formulaCounter");
  if (counterEl) {
    counterEl.textContent = `${visibleCount} Materi`;
  }
}

function selectFormula(index) {
  activeFormulaIndex = index;

  $$(".sidebar__item").forEach((el) => {
    const elIdx = parseInt(el.dataset.index, 10);
    const isActive = elIdx === activeFormulaIndex;
    el.classList.toggle("is-active", isActive);
    el.setAttribute("aria-selected", String(isActive));
  });

  renderFormulaDetail(index);
}

/* ==========================================================
   RENDER FORMULA DETAIL VIEW (RIGHT CONTENT STAGE)
   ========================================================== */
function renderFormulaDetail(index) {
  const categoryData = DATABASE[currentCategory];
  if (!categoryData) return;
  const item = categoryData[index];
  if (!item) return;

  const statNum = $("#statActiveNum");
  if (statNum) statNum.textContent = item.number;

  const container = $("#formulaDetailContainer");
  if (!container) return;

  // 1. Generate Syntax Blocks
  const syntaxHTML = item.syntax.map((s) => `
    <div class="syntax-card">
      <span class="syntax-card__label">${s.label}</span>
      <div class="code-block">
        <pre><code>${escapeHTML(s.code)}</code></pre>
        <button type="button" class="copy-btn" data-copy="${escapeAttribute(s.code)}" aria-label="Salin rumus ${item.title}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <rect width="14" height="14" x="8" y="8" rx="2" ry="2"/>
            <path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>
          </svg>
          <span>Copy</span>
        </button>
      </div>
      <p class="syntax-params">${s.params}</p>
    </div>
  `).join("");

  // 2. Generate Table Preview HTML
  const thHTML = item.example.tableHeaders.map((th) => `<th>${th}</th>`).join("");
  const trHTML = item.example.tableRows.map((row) => `
    <tr>
      ${row.map((td, i) => `<td class="${i === 0 ? 'font-mono' : ''}">${td}</td>`).join("")}
    </tr>
  `).join("");

  // 3. Generate Interactive Google Sheets Embed URL
  const safeTitleSlug = item.title.replace(/[^a-zA-Z0-9]/g, '_');
  const sheetsEmbedUrl = `https://docs.google.com/spreadsheets/d/e/2PACX-1vT-RUMUS-EXCEL-SIMULATION-${currentCategory.toUpperCase()}-${item.number}/pubhtml?widget=true&headers=false`;

  // 4. Render Entire Detail View
  container.innerHTML = `
    <!-- Detail Header -->
    <header class="view-header">
      <div class="view-header__meta">
        <span class="badge-num">RUMUS #${item.number}</span>
        <span class="badge-tag">${item.tag}</span>
        <span class="badge-diff">Tingkat: ${item.difficulty}</span>
      </div>
      <h2 class="view-title">${item.title}</h2>
      <p class="view-subtitle">${item.purpose}</p>
    </header>

    <!-- Section 1: Sintaks Resmi -->
    <section class="view-card" aria-labelledby="syntaxCardHeading">
      <div class="view-card__head">
        <span class="card-chip">SINTAKS RESMI</span>
        <h3 class="card-title" id="syntaxCardHeading">Bentuk Penulisan Rumus Excel</h3>
        <p class="card-desc">Gunakan tombol Copy di sebelah kanan untuk menyalin format formula secara langsung.</p>
      </div>
      <div class="syntax-grid">
        ${syntaxHTML}
      </div>
    </section>

    <!-- Section 2: Studi Kasus Riil & Contoh Tabel -->
    <section class="view-card" aria-labelledby="exampleCardHeading">
      <div class="view-card__head">
        <span class="card-chip">CONTOH KASUS PRAKTIK</span>
        <h3 class="card-title" id="exampleCardHeading">Simulasi Lembar Kerja Kasus Nyata</h3>
        <p class="card-desc">${item.example.scenario}</p>
      </div>

      <div class="table-responsive">
        <table class="figma-table">
          <thead>
            <tr>${thHTML}</tr>
          </thead>
          <tbody>
            ${trHTML}
          </tbody>
        </table>
      </div>

      <!-- Result Showcase with Copy Button for Formula Used -->
      <div class="result-showcase">
        <div class="result-showcase__calc">
          <span class="res-label">Formula yang Diterapkan:</span>
          <div class="res-code-wrapper">
            <code class="res-code">${escapeHTML(item.example.formulaUsed)}</code>
            <button type="button" class="copy-btn copy-btn--sm" data-copy="${escapeAttribute(item.example.formulaUsed)}" aria-label="Salin rumus contoh">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect width="14" height="14" x="8" y="8" rx="2" ry="2"/>
                <path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>
              </svg>
              <span>Copy</span>
            </button>
          </div>
        </div>
        <div class="result-showcase__val">
          <span class="res-label">Output Hasil Excel:</span>
          <output class="res-output">${item.example.result}</output>
        </div>
      </div>

      <p class="result-explanation">
        <strong>Penjelasan Langkah Kerja:</strong> ${item.example.explanation}
      </p>
    </section>

    <!-- Section 3: Live Google Sheets Simulation -->
    <section class="sheets-sim-card" aria-labelledby="sheetsHeading">
      <div class="view-card__head">
        <span class="card-chip card-chip--sheets">SIMULASI LANGSUNG</span>
        <h3 class="card-title" id="sheetsHeading">Praktik Interaktif di Google Sheets</h3>
        <p class="card-desc">
          Lakukan eksperimen rumus <strong>${item.title}</strong> langsung di lembar kerja Google Sheets di bawah ini. Anda dapat mengklik dan melihat kalkulasinya secara realtime!
        </p>
      </div>

      <div class="sheets-frame-wrapper">
        <div class="sheets-mock-toolbar">
          <div class="sheets-mock-title">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="#0F9D58">
              <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 7h10v2H7zm0 4h10v2H7zm0 4h7v2H7z"/>
            </svg>
            <span>Latihan_${currentCategory.toUpperCase()}_${item.number}_${safeTitleSlug}.xlsx</span>
          </div>
          <span class="sheets-mock-status">Tersinkronisasi Otomatis ✓</span>
        </div>

        <iframe 
          class="sheets-iframe" 
          src="${sheetsEmbedUrl}" 
          title="Google Sheets Simulation - ${item.title}"
          loading="lazy"
          allowfullscreen>
        </iframe>
      </div>

      <div class="sheets-action-bar">
        <div class="sheets-action-info">
          <strong>Ingin mengedit di akun Google Anda sendiri?</strong>
          <span>Buka template spreadsheet resmi ini untuk meng-copy dummy data &amp; menyimpan latihan Anda.</span>
        </div>
        <a 
          href="https://docs.google.com/spreadsheets/create" 
          class="btn-sheets" 
          target="_blank" 
          rel="noopener noreferrer"
          title="Buka template ini di Google Sheets baru">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
            <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 7h10v2H7zm0 4h10v2H7zm0 4h7v2H7z"/>
          </svg>
          <span>Buka Template di Google Sheets ↗</span>
        </a>
      </div>
    </section>

    <!-- Section 4: Tips Praktisi Kantor -->
    <section class="tip-banner">
      <div class="tip-banner__icon">💡</div>
      <div class="tip-banner__content">
        <strong>Tips Praktisi Kantor:</strong>
        <p>${item.proTip}</p>
      </div>
    </section>

    <!-- Section 5: Navigasi Urutan Rumus -->
    <nav class="view-nav-buttons" aria-label="Navigasi urutan rumus">
      <button class="btn btn--outline btn--sm" id="btnPrevFormula" ${index === 0 ? "disabled" : ""}>
        ← Rumus Sebelumnya
      </button>
      <span class="view-nav-pos">${index + 1} dari 15 Rumus</span>
      <button class="btn btn--primary btn--sm" id="btnNextFormula" ${index === 14 ? "disabled" : ""}>
        Rumus Selanjutnya →
      </button>
    </nav>
  `;

  // Attach event listener copy buttons
  $$(".copy-btn", container).forEach((btn) => {
    btn.addEventListener("click", () => {
      const text = btn.dataset.copy;
      copyFormulaText(text, btn);
    });
  });

  // Attach event listener navigasi Sebelumnya / Selanjutnya
  const btnPrev = $("#btnPrevFormula");
  const btnNext = $("#btnNextFormula");

  if (btnPrev && index > 0) {
    btnPrev.addEventListener("click", () => {
      selectFormula(index - 1);
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  }

  if (btnNext && index < 14) {
    btnNext.addEventListener("click", () => {
      selectFormula(index + 1);
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  }
}

/* ==========================================================
   FEATURE: COPY TO CLIPBOARD BUTTON LOGIC
   ========================================================== */
async function copyFormulaText(textToCopy, btnElement) {
  if (!textToCopy) return;

  try {
    await navigator.clipboard.writeText(textToCopy);
  } catch (err) {
    const textarea = document.createElement("textarea");
    textarea.value = textToCopy;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
  }

  if (btnElement) {
    const originalHTML = btnElement.innerHTML;
    btnElement.classList.add("is-copied");
    btnElement.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="20 6 9 17 4 12"/>
      </svg>
      <span>Copied! ✅</span>
    `;

    setTimeout(() => {
      btnElement.classList.remove("is-copied");
      btnElement.innerHTML = originalHTML;
    }, 2000);
  }

  showToast("Formula berhasil disalin ke clipboard! 📋");
}

/* ==========================================================
   INTERACTIVE PRACTICE DEMO (TRIM + PROPER ON LANDING)
   ========================================================== */
function initPracticeDemo() {
  const btn = $("#btnTestClean");
  const table = $("#testEmployeeTable");
  if (!btn || !table) return;

  let isCleaned = false;

  btn.addEventListener("click", () => {
    isCleaned = !isCleaned;
    const cells = $$(".raw-cell", table);

    cells.forEach((cell) => {
      const rawText = cell.getAttribute("data-raw") || cell.textContent;
      if (isCleaned) {
        const trimmed = rawText.trim();
        const proper = trimmed.replace(/\b\w/g, (char) => char.toUpperCase());
        cell.textContent = proper;
        cell.style.color = "var(--c-primary)";
        cell.style.fontWeight = "700";
      } else {
        cell.textContent = `"${rawText}"`;
        cell.style.color = "";
        cell.style.fontWeight = "";
      }
    });

    btn.textContent = isCleaned ? "↺ Reset Data Mentah" : "✨ Uji TRIM + PROPER";
    showToast(isCleaned ? "Data berhasil dibersihkan dengan TRIM & PROPER! ✨" : "Data dikembalikan ke format mentah berantakan.");
  });
}

/* ==========================================================
   HELPER UTILITIES: HTML ESCAPING
   ========================================================== */
function escapeHTML(str) {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttribute(str) {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/* ==========================================================
   TOAST NOTIFICATION HELPER
   ========================================================== */
function showToast(message) {
  const toast = $("#toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add("is-visible");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => {
    toast.classList.remove("is-visible");
  }, 2400);
}

/* ==========================================================
   CANVAS ANIMATED FAVICON
   ========================================================== */
function initAnimatedFavicon() {
  const link = $("#favicon");
  if (!link) return;

  const size = 64;
  const canvas = document.createElement("canvas");
  canvas.width = canvas.height = size;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  let frame = 0;
  const reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  function roundRect(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y + x, y, r);
    ctx.closePath();
  }

  function draw() {
    try {
      const pulse = (Math.sin(frame / 7) + 1) / 2;
      const scale = 0.85 + pulse * 0.15;
      const s = size * scale;
      const offset = (size - s) / 2;

      ctx.clearRect(0, 0, size, size);

      ctx.fillStyle = currentCategory === "accounting" ? "#19734D" : "#123D2E";
      roundRect(offset, offset, s, s, s * 0.22);
      ctx.fill();

      ctx.strokeStyle = "rgba(255, 255, 255, 0.35)";
      ctx.lineWidth = 2;
      for (let i = 1; i < 3; i++) {
        const p = offset + (s / 3) * i;
        ctx.beginPath(); ctx.moveTo(offset + 4, p); ctx.lineTo(offset + s - 4, p); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(p, offset + 4); ctx.lineTo(p, offset + s - 4); ctx.stroke();
      }

      ctx.fillStyle = "#FFFFFF";
      ctx.font = `800 ${Math.round(s * 0.58)}px 'Plus Jakarta Sans', system-ui, sans-serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText("X", size / 2, size / 2 + s * 0.04);

      link.href = canvas.toDataURL("image/png");
      link.type = "image/png";
      frame++;
    } catch (e) {}
  }

  draw();
  if (!reduceMotion) setInterval(draw, 140);
}
