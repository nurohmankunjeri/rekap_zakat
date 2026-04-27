<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Pembagian Zakat - Full Version</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Inter', sans-serif; }
        body { background: #f0f0f0; padding: 20px; }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { background: white; padding: 20px 24px; border-radius: 16px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
        .header h1 { font-size: 22px; font-weight: 700; display: flex; align-items: center; gap: 10px; }
        .badge-date { background: #f0f0f0; padding: 6px 16px; border-radius: 30px; font-size: 13px; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
        .stat-card { background: white; border-radius: 16px; padding: 16px; border: 1px solid #e8e8e8; }
        .stat-card .label { font-size: 12px; color: #888; margin-bottom: 8px; }
        .stat-card .value { font-size: 28px; font-weight: 700; color: #1a1a1a; }
        .tab-menu { display: flex; gap: 8px; margin-bottom: 20px; border-bottom: 1px solid #ddd; flex-wrap: wrap; }
        .tab-btn { background: none; border: none; padding: 10px 20px; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 10px 10px 0 0; color: #666; }
        .tab-btn.active { color: #333; border-bottom: 2px solid #333; background: white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .two-columns { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
        @media (max-width: 900px) { .two-columns { grid-template-columns: 1fr; } .stats-grid { grid-template-columns: 1fr 1fr; } }
        .card { background: white; border-radius: 16px; border: 1px solid #e8e8e8; overflow: hidden; margin-bottom: 20px; }
        .card-header { padding: 14px 18px; border-bottom: 1px solid #efefef; font-weight: 600; font-size: 15px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }
        .card-content { padding: 16px 18px; }
        .form-group { margin-bottom: 14px; display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
        .form-group input, .form-group select { flex: 1; padding: 10px 14px; border: 1px solid #ddd; border-radius: 10px; font-size: 14px; }
        .btn-primary { background: #2c2c2c; color: white; border: none; padding: 10px 20px; border-radius: 10px; font-weight: 500; cursor: pointer; }
        .btn-success { background: #27ae60; color: white; border: none; padding: 10px 20px; border-radius: 10px; cursor: pointer; }
        .btn-danger { background: #e74c3c; color: white; border: none; padding: 8px 14px; border-radius: 10px; cursor: pointer; }
        .btn-outline { background: transparent; border: 1px solid #ccc; color: #555; padding: 6px 12px; border-radius: 8px; cursor: pointer; }
        .btn-sm { padding: 4px 10px; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { text-align: left; padding: 10px 8px; background: #fafafa; border-bottom: 1px solid #eee; font-weight: 600; }
        td { padding: 10px 8px; border-bottom: 1px solid #f0f0f0; }
        .badge { background: #f0f0f0; padding: 4px 10px; border-radius: 20px; font-size: 11px; }
        .master-item { display: flex; justify-content: space-between; align-items: center; padding: 10px; background: #fafafa; border-radius: 10px; margin-bottom: 8px; flex-wrap: wrap; gap: 8px; }
        .export-buttons { display: flex; gap: 12px; justify-content: flex-end; margin-bottom: 16px; }
        .info-box { background: #e8f5e9; padding: 12px; border-radius: 12px; margin-bottom: 16px; font-size: 13px; }
        footer { text-align: center; margin-top: 30px; padding: 15px; color: #999; font-size: 11px; }
        
        @media print {
            body { background: white; padding: 0; margin: 0; }
            .no-print, .tab-menu, .btn-primary, .btn-success, .stats-grid, .header .badge-date, footer, .export-buttons { display: none !important; }
            .card { border: 1px solid #ddd; margin-bottom: 20px; page-break-inside: avoid; }
        }
    </style>
</head>
<body>
<div class="container">
    <div class="header no-print">
        <h1><i class="fas fa-hand-holding-heart"></i> Baitulmaal · Zakat Management</h1>
        <div class="badge-date"><i class="far fa-calendar-alt"></i> Ramadan 1447 H</div>
    </div>

    <div class="stats-grid no-print">
        <div class="stat-card"><div class="label">TOTAL UANG MASUK</div><div class="value" id="totalUangValue">Rp 0</div></div>
        <div class="stat-card"><div class="label">TOTAL BERAS MASUK</div><div class="value" id="totalBerasValue">0 kg</div></div>
        <div class="stat-card"><div class="label">TOTAL UANG TERSALUR</div><div class="value" id="totalTersalurValue">Rp 0</div></div>
        <div class="stat-card"><div class="label">TOTAL BERAS TERSALUR</div><div class="value" id="totalBerasTersalurValue">0 kg</div></div>
    </div>

    <div class="tab-menu no-print">
        <button class="tab-btn" data-tab="master">⚙️ Master Data</button>
        <button class="tab-btn active" data-tab="penerimaan">📥 Data Penerimaan</button>
        <button class="tab-btn" data-tab="penyaluran">📤 Data Penyaluran</button>
        <button class="tab-btn" data-tab="export">📄 Export Laporan</button>
    </div>

    <!-- TAB MASTER DATA -->
    <div id="tab-master" class="tab-content">
        <div class="two-columns">
            <!-- Master Beras -->
            <div class="card">
                <div class="card-header">
                    <span><i class="fas fa-box"></i> Master Data Beras</span>
                    <button class="btn-primary btn-sm" id="tambahBerasBtn">+ Tambah</button>
                </div>
                <div class="card-content" id="masterBerasList">
                    <!-- akan diisi javascript -->
                </div>
                <div class="card-content" style="border-top:1px solid #eee; background:#fafafa;">
                    <small><i class="fas fa-info-circle"></i> Zakat Fitrah = 2,5 kg beras per jiwa</small>
                </div>
            </div>
            
            <!-- Master Uang -->
            <div class="card">
                <div class="card-header">
                    <span><i class="fas fa-money-bill-wave"></i> Master Konversi Uang</span>
                    <button class="btn-primary btn-sm" id="tambahUangBtn">+ Tambah</button>
                </div>
                <div class="card-content" id="masterUangList">
                    <!-- akan diisi javascript -->
                </div>
                <div class="card-content" style="border-top:1px solid #eee; background:#fafafa;">
                    <small><i class="fas fa-info-circle"></i> Digunakan untuk konversi zakat fitrah per jiwa</small>
                </div>
            </div>
        </div>
    </div>

    <!-- TAB PENERIMAAN -->
    <div id="tab-penerimaan" class="tab-content active">
        <div class="card">
            <div class="card-header">
                <span><i class="fas fa-plus-circle"></i> Form Penerimaan Zakat</span>
                <button class="btn-danger btn-sm" id="resetPenerimaanBtn">Reset Semua</button>
            </div>
            <div class="card-content">
                <div class="form-group">
                    <input type="text" id="namaMuzakki" placeholder="Nama Muzakki">
                    <select id="jenisZakat">
                        <option value="fitrah_uang">Zakat Fitrah (Uang)</option>
                        <option value="fitrah_beras">Zakat Fitrah (Beras)</option>
                        <option value="maal">Zakat Maal</option>
                    </select>
                </div>
                <div class="form-group" id="dynamicFieldPenerimaan"></div>
                <div class="form-group">
                    <input type="number" id="jumlahJiwa" placeholder="Jumlah Jiwa" value="1" min="1">
                    <button class="btn-primary" id="btnSimpanPenerimaan">Simpan</button>
                </div>
            </div>
        </div>
        <div class="card">
            <div class="card-header">Daftar Penerimaan Zakat</div>
            <div class="card-content table-wrapper" style="overflow-x:auto;">
                <table id="tabelPenerimaan">
                    <thead><tr><th>No</th><th>Nama</th><th>Jenis</th><th>Detail</th><th>Uang (Rp)</th><th>Beras (kg)</th><th>Tanggal</th><th>Aksi</th></tr></thead>
                    <tbody id="tbodyPenerimaan"></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- TAB PENYALURAN -->
    <div id="tab-penyaluran" class="tab-content">
        <div class="card">
            <div class="card-header">
                <span><i class="fas fa-hand-holding-heart"></i> Form Penyaluran Zakat</span>
                <button class="btn-danger btn-sm" id="resetPenyaluranBtn">Reset Semua</button>
            </div>
            <div class="card-content">
                <div class="form-group">
                    <select id="asnafPilih">
                        <option value="">Pilih Asnaf</option>
                        <option value="Fakir">Fakir</option><option value="Miskin">Miskin</option><option value="Amil">Amil</option>
                        <option value="Muallaf">Muallaf</option><option value="Riqab">Riqab</option><option value="Gharimin">Gharimin</option>
                        <option value="Fisabilillah">Fisabilillah</option><option value="Ibnu Sabil">Ibnu Sabil</option>
                    </select>
                    <input type="text" id="namaPenerima" placeholder="Nama Penerima">
                </div>
                <div class="form-group">
                    <input type="number" id="jumlahUang" placeholder="Uang (Rp)" step="1000">
                    <input type="number" id="jumlahBeras" placeholder="Beras (kg)" step="0.5">
                    <button class="btn-primary" id="btnSimpanPenyaluran">Simpan Penyaluran</button>
                </div>
            </div>
        </div>
        <div class="card">
            <div class="card-header">Daftar Penyaluran Zakat</div>
            <div class="card-content table-wrapper" style="overflow-x:auto;">
                <table id="tabelPenyaluran">
                    <thead><tr><th>No</th><th>Tanggal</th><th>Asnaf</th><th>Nama Penerima</th><th>Uang (Rp)</th><th>Beras (kg)</th><th>Aksi</th></tr></thead>
                    <tbody id="tbodyPenyaluran"></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- TAB EXPORT -->
    <div id="tab-export" class="tab-content">
        <div class="card">
            <div class="card-header">Export Laporan</div>
            <div class="card-content">
                <div class="info-box">
                    <i class="fas fa-info-circle"></i> <strong>Data saat ini:</strong><br>
                    - Penerimaan: <span id="infoJmlPenerimaan">0</span> data<br>
                    - Penyaluran: <span id="infoJmlPenyaluran">0</span> data
                </div>
                <div style="display: flex; gap: 20px; flex-wrap: wrap; margin-top: 20px;">
                    <button class="btn-primary" id="exportExcelBtn" style="background: #27ae60; padding: 12px 24px;">
                        <i class="fas fa-file-excel"></i> Export ke Excel (XLSX)
                    </button>
                    <button class="btn-primary" id="exportPDFBtn" style="background: #e74c3c; padding: 12px 24px;">
                        <i class="fas fa-file-pdf"></i> Export ke PDF (Print)
                    </button>
                </div>
            </div>
        </div>
    </div>
    <footer class="no-print">Sistem Pengelolaan Zakat - Master data bisa diedit di tab Master Data</footer>
</div>

<script>
    // ==================== MASTER DATA ====================
    let masterBeras = [
        { id: 1, nama: "Beras Premium", hargaPerKg: 16000 },
        { id: 2, nama: "Beras Medium", hargaPerKg: 14000 },
        { id: 3, nama: "Beras Standar", hargaPerKg: 12500 }
    ];
    
    let masterUang = [
        { id: 1, nama: "Konversi Premium", nominalPerJiwa: 40000 },
        { id: 2, nama: "Konversi Medium", nominalPerJiwa: 35000 },
        { id: 3, nama: "Konversi Standar", nominalPerJiwa: 31250 }
    ];

    // ==================== DATA TRANSAKSI ====================
    let penerimaan = [];
    let penyaluran = [];

    // Data contoh demo
    function loadDemoData() {
        penerimaan = [
            { id: 1, nama: "H. Ahmad Fauzi", jenis: "Zakat Fitrah (Uang)", detail: "Konversi Standar (4 jiwa)", uang: 125000, beras: 0, tanggal: "25/04/2026" },
            { id: 2, nama: "Ibu Siti Rahmah", jenis: "Zakat Fitrah (Beras)", detail: "Beras Premium - 7.5kg (3 jiwa)", uang: 0, beras: 7.5, tanggal: "25/04/2026" },
            { id: 3, nama: "PT. Berkah Abadi", jenis: "Zakat Maal", detail: "Zakat Perusahaan", uang: 15000000, beras: 0, tanggal: "24/04/2026" },
            { id: 4, nama: "Bpk. Joko Widodo", jenis: "Zakat Fitrah (Uang)", detail: "Konversi Medium (5 jiwa)", uang: 175000, beras: 0, tanggal: "24/04/2026" },
            { id: 5, nama: "Kel. Bambang", jenis: "Zakat Fitrah (Beras)", detail: "Beras Standar - 12.5kg (5 jiwa)", uang: 0, beras: 12.5, tanggal: "23/04/2026" }
        ];
        penyaluran = [
            { id: 1, tanggal: "25/04/2026", asnaf: "Fakir", penerima: "Ibu Fatimah", uang: 2500000, beras: 50 },
            { id: 2, tanggal: "25/04/2026", asnaf: "Miskin", penerima: "Bpk. Slamet", uang: 3000000, beras: 60 },
            { id: 3, tanggal: "24/04/2026", asnaf: "Amil", penerima: "Ustadz Abdul", uang: 1500000, beras: 20 },
            { id: 4, tanggal: "24/04/2026", asnaf: "Muallaf", penerima: "Saudara Ahmad", uang: 1000000, beras: 15 },
            { id: 5, tanggal: "23/04/2026", asnaf: "Fisabilillah", penerima: "Santri Tahfidz", uang: 2000000, beras: 30 }
        ];
    }
    loadDemoData();

    let nextIdPenerimaan = 6;
    let nextIdPenyaluran = 6;

    // Helper functions
    function formatRupiah(angka) { return "Rp " + (angka || 0).toLocaleString('id-ID'); }
    function totalUangMasuk() { return penerimaan.reduce((a,b)=> a + (b.uang||0), 0); }
    function totalBerasMasuk() { return penerimaan.reduce((a,b)=> a + (b.beras||0), 0); }
    function totalUangTersalur() { return penyaluran.reduce((a,b)=> a + (b.uang||0), 0); }
    function totalBerasTersalur() { return penyaluran.reduce((a,b)=> a + (b.beras||0), 0); }

    function updateStats() {
        document.getElementById("totalUangValue").innerHTML = formatRupiah(totalUangMasuk());
        document.getElementById("totalBerasValue").innerHTML = totalBerasMasuk().toFixed(1) + " kg";
        document.getElementById("totalTersalurValue").innerHTML = formatRupiah(totalUangTersalur());
        document.getElementById("totalBerasTersalurValue").innerHTML = totalBerasTersalur().toFixed(1) + " kg";
        document.getElementById("infoJmlPenerimaan").innerHTML = penerimaan.length;
        document.getElementById("infoJmlPenyaluran").innerHTML = penyaluran.length;
    }

    // ==================== RENDER MASTER BERAS ====================
    function renderMasterBeras() {
        const container = document.getElementById("masterBerasList");
        if (!container) return;
        container.innerHTML = "";
        masterBeras.forEach(b => {
            container.innerHTML += `
                <div class="master-item">
                    <div><strong>${b.nama}</strong><br><small>Rp ${b.hargaPerKg.toLocaleString()}/kg</small></div>
                    <div>
                        <button class="btn-outline btn-sm" onclick="editBeras(${b.id})">Edit</button>
                        <button class="btn-danger btn-sm" onclick="hapusBeras(${b.id})">Hapus</button>
                    </div>
                </div>
            `;
        });
        if (masterBeras.length === 0) container.innerHTML = "<div class='info-box'>Belum ada data beras. Silakan tambah.</div>";
        updateFormPenerimaan();
    }

    window.editBeras = (id) => {
        let b = masterBeras.find(x => x.id === id);
        if (!b) return;
        let newNama = prompt("Nama beras:", b.nama);
        let newHarga = prompt("Harga per kg (Rp):", b.hargaPerKg);
        if (newNama && newNama.trim()) b.nama = newNama.trim();
        if (newHarga && !isNaN(parseFloat(newHarga))) b.hargaPerKg = parseFloat(newHarga);
        renderMasterBeras();
    };
    window.hapusBeras = (id) => {
        if (confirm("Hapus beras ini?")) {
            masterBeras = masterBeras.filter(b => b.id !== id);
            renderMasterBeras();
        }
    };
    document.getElementById("tambahBerasBtn")?.addEventListener("click", () => {
        let nama = prompt("Nama beras baru:");
        let harga = prompt("Harga per kg (Rp):");
        if (nama && harga && !isNaN(parseFloat(harga))) {
            let newId = Math.max(0, ...masterBeras.map(b => b.id), 0) + 1;
            masterBeras.push({ id: newId, nama: nama, hargaPerKg: parseFloat(harga) });
            renderMasterBeras();
        }
    });

    // ==================== RENDER MASTER UANG ====================
    function renderMasterUang() {
        const container = document.getElementById("masterUangList");
        if (!container) return;
        container.innerHTML = "";
        masterUang.forEach(u => {
            container.innerHTML += `
                <div class="master-item">
                    <div><strong>${u.nama}</strong><br><small>Rp ${u.nominalPerJiwa.toLocaleString()}/jiwa</small></div>
                    <div>
                        <button class="btn-outline btn-sm" onclick="editUang(${u.id})">Edit</button>
                        <button class="btn-danger btn-sm" onclick="hapusUang(${u.id})">Hapus</button>
                    </div>
                </div>
            `;
        });
        if (masterUang.length === 0) container.innerHTML = "<div class='info-box'>Belum ada data konversi. Silakan tambah.</div>";
        updateFormPenerimaan();
    }

    window.editUang = (id) => {
        let u = masterUang.find(x => x.id === id);
        if (!u) return;
        let newNama = prompt("Nama konversi:", u.nama);
        let newNominal = prompt("Nominal per jiwa (Rp):", u.nominalPerJiwa);
        if (newNama && newNama.trim()) u.nama = newNama.trim();
        if (newNominal && !isNaN(parseFloat(newNominal))) u.nominalPerJiwa = parseFloat(newNominal);
        renderMasterUang();
    };
    window.hapusUang = (id) => {
        if (confirm("Hapus konversi ini?")) {
            masterUang = masterUang.filter(u => u.id !== id);
            renderMasterUang();
        }
    };
    document.getElementById("tambahUangBtn")?.addEventListener("click", () => {
        let nama = prompt("Nama konversi uang:");
        let nominal = prompt("Nominal per jiwa (Rp):");
        if (nama && nominal && !isNaN(parseFloat(nominal))) {
            let newId = Math.max(0, ...masterUang.map(u => u.id), 0) + 1;
            masterUang.push({ id: newId, nama: nama, nominalPerJiwa: parseFloat(nominal) });
            renderMasterUang();
        }
    });

    // ==================== FORM PENERIMAAN DINAMIS ====================
    const jenisSelect = document.getElementById("jenisZakat");
    const dynamicDiv = document.getElementById("dynamicFieldPenerimaan");
    function updateFormPenerimaan() {
        let jenis = jenisSelect?.value;
        if (!dynamicDiv) return;
        if (jenis === "fitrah_beras") {
            let options = masterBeras.map(b => `<option value="${b.id}" data-harga="${b.hargaPerKg}">${b.nama} (Rp ${b.hargaPerKg.toLocaleString()}/kg)</option>`).join('');
            dynamicDiv.innerHTML = `<select id="kategoriBeras" style="flex:1">${options}</select>
                                    <input type="number" id="nominalBeras" placeholder="Total Beras (kg) - kosongkan auto">`;
        } else if (jenis === "fitrah_uang") {
            let options = masterUang.map(u => `<option value="${u.id}" data-nominal="${u.nominalPerJiwa}">${u.nama} (Rp ${u.nominalPerJiwa.toLocaleString()}/jiwa)</option>`).join('');
            dynamicDiv.innerHTML = `<select id="kategoriUang" style="flex:1">${options}</select>
                                    <input type="number" id="nominalUang" placeholder="Nominal Uang - kosongkan auto">`;
        } else {
            dynamicDiv.innerHTML = `<input type="number" id="nominalMaal" placeholder="Zakat Maal (Rp)">`;
        }
    }
    jenisSelect?.addEventListener("change", updateFormPenerimaan);
    updateFormPenerimaan();

    // ==================== SIMPAN PENERIMAAN ====================
    document.getElementById("btnSimpanPenerimaan")?.addEventListener("click", () => {
        let nama = document.getElementById("namaMuzakki")?.value.trim();
        if (!nama) { alert("Masukkan nama muzakki"); return; }
        let jenis = jenisSelect.value;
        let jiwa = parseInt(document.getElementById("jumlahJiwa")?.value) || 1;
        let totalUang = 0, totalBeras = 0, detail = "";
        let tgl = new Date().toLocaleDateString('id-ID');
        let jenisTeks = "";

        if (jenis === "fitrah_beras") {
            let selectedId = document.getElementById("kategoriBeras")?.value;
            let beras = masterBeras.find(b => b.id == selectedId);
            let manualKg = parseFloat(document.getElementById("nominalBeras")?.value);
            totalBeras = (!isNaN(manualKg) && manualKg > 0) ? manualKg : jiwa * 2.5;
            detail = `${beras?.nama || 'Beras'} - ${totalBeras}kg (${jiwa} jiwa)`;
            jenisTeks = "Zakat Fitrah (Beras)";
        } else if (jenis === "fitrah_uang") {
            let nominalManual = parseFloat(document.getElementById("nominalUang")?.value);
            if (!isNaN(nominalManual) && nominalManual > 0) {
                totalUang = nominalManual;
                detail = `Bayar langsung Rp${totalUang.toLocaleString()}`;
            } else {
                let selectedId = document.getElementById("kategoriUang")?.value;
                let konversi = masterUang.find(u => u.id == selectedId);
                totalUang = jiwa * (konversi?.nominalPerJiwa || 31250);
                detail = `${konversi?.nama || 'Konversi'} (${jiwa} jiwa)`;
            }
            jenisTeks = "Zakat Fitrah (Uang)";
        } else {
            let nominal = parseFloat(document.getElementById("nominalMaal")?.value);
            if (isNaN(nominal) || nominal <= 0) { alert("Masukkan nominal maal"); return; }
            totalUang = nominal;
            detail = `Zakat Maal`;
            jenisTeks = "Zakat Maal";
        }
        penerimaan.push({ id: nextIdPenerimaan++, nama, jenis: jenisTeks, detail, uang: totalUang, beras: totalBeras, tanggal: tgl });
        renderPenerimaan();
        document.getElementById("namaMuzakki").value = "";
        if(document.getElementById("nominalBeras")) document.getElementById("nominalBeras").value = "";
        if(document.getElementById("nominalUang")) document.getElementById("nominalUang").value = "";
        if(document.getElementById("nominalMaal")) document.getElementById("nominalMaal").value = "";
        alert(`✅ ${nama} tercatat`);
    });

    // ==================== RENDER PENERIMAAN ====================
    function renderPenerimaan() {
        let tbody = document.getElementById("tbodyPenerimaan");
        if (!tbody) return;
        tbody.innerHTML = "";
        penerimaan.forEach((p, idx) => {
            tbody.innerHTML += `<tr><td>${idx+1}</td><td>${p.nama}</td><td>${p.jenis}</td><td>${p.detail}</td><td>${formatRupiah(p.uang)}</td><td>${p.beras.toFixed(1)} kg</td><td>${p.tanggal}</td><td><button class="btn-danger btn-sm" onclick="hapusPenerimaan(${p.id})">Hapus</button></td></tr>`;
        });
        if(penerimaan.length === 0) tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;">Belum ada data penerimaan</td></tr>';
        updateStats();
    }

    // ==================== SIMPAN PENYALURAN ====================
    document.getElementById("btnSimpanPenyaluran")?.addEventListener("click", () => {
        let asnaf = document.getElementById("asnafPilih")?.value;
        let nama = document.getElementById("namaPenerima")?.value.trim();
        let uang = parseFloat(document.getElementById("jumlahUang")?.value) || 0;
        let beras = parseFloat(document.getElementById("jumlahBeras")?.value) || 0;
        if (!asnaf || !nama) { alert("Pilih asnaf dan isi nama penerima"); return; }
        if (uang === 0 && beras === 0) { alert("Masukkan nominal uang atau beras"); return; }
        let tgl = new Date().toLocaleDateString('id-ID');
        penyaluran.push({ id: nextIdPenyaluran++, tanggal: tgl, asnaf, penerima: nama, uang, beras });
        renderPenyaluran();
        document.getElementById("namaPenerima").value = "";
        document.getElementById("jumlahUang").value = "";
        document.getElementById("jumlahBeras").value = "";
        alert(`✅ ${nama} (${asnaf}) tercatat`);
    });

    // ==================== RENDER PENYALURAN ====================
    function renderPenyaluran() {
        let tbody = document.getElementById("tbodyPenyaluran");
        if (!tbody) return;
        tbody.innerHTML = "";
        penyaluran.forEach((p, idx) => {
            tbody.innerHTML += `<tr><td>${idx+1}</td><td>${p.tanggal}</td><td>${p.asnaf}</td><td>${p.penerima}</td><td>${formatRupiah(p.uang)}</td><td>${p.beras.toFixed(1)} kg</td><td><button class="btn-danger btn-sm" onclick="hapusPenyaluran(${p.id})">Hapus</button></td></tr>`;
        });
        if(penyaluran.length === 0) tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;">Belum ada data penyaluran</td></tr>';
        updateStats();
    }

    window.hapusPenerimaan = (id) => { penerimaan = penerimaan.filter(p => p.id !== id); renderPenerimaan(); };
    window.hapusPenyaluran = (id) => { penyaluran = penyaluran.filter(p => p.id !== id); renderPenyaluran(); };

    document.getElementById("resetPenerimaanBtn")?.addEventListener("click", () => {
        if(confirm("Hapus semua data penerimaan?")) { penerimaan = []; renderPenerimaan(); }
    });
    document.getElementById("resetPenyaluranBtn")?.addEventListener("click", () => {
        if(confirm("Hapus semua data penyaluran?")) { penyaluran = []; renderPenyaluran(); }
    });

    // ==================== EXPORT EXCEL ====================
    document.getElementById("exportExcelBtn")?.addEventListener("click", () => {
        let dataPenerimaan = penerimaan.map((p, i) => ({
            "No": i+1, "Nama Muzakki": p.nama, "Jenis Zakat": p.jenis, "Detail": p.detail,
            "Uang (Rp)": p.uang, "Beras (kg)": p.beras, "Tanggal": p.tanggal
        }));
        let dataPenyaluran = penyaluran.map((p, i) => ({
            "No": i+1, "Tanggal": p.tanggal, "Asnaf": p.asnaf, "Nama Penerima": p.penerima,
            "Uang (Rp)": p.uang, "Beras (kg)": p.beras
        }));
        let ringkasan = [
            { "Keterangan": "Total Uang Masuk", "Nilai": totalUangMasuk() },
            { "Keterangan": "Total Beras Masuk (kg)", "Nilai": totalBerasMasuk() },
            { "Keterangan": "Total Uang Tersalur", "Nilai": totalUangTersalur() },
            { "Keterangan": "Total Beras Tersalur (kg)", "Nilai": totalBerasTersalur() },
            { "Keterangan": "Sisa Uang", "Nilai": totalUangMasuk() - totalUangTersalur() },
            { "Keterangan": "Sisa Beras (kg)", "Nilai": totalBerasMasuk() - totalBerasTersalur() }
        ];
        let wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ringkasan), "Ringkasan");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataPenerimaan), "Data Penerimaan");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataPenyaluran), "Data Penyaluran");
        XLSX.writeFile(wb, `laporan_zakat_${new Date().toISOString().slice(0,10)}.xlsx`);
    });

    // ==================== EXPORT PDF ====================
    document.getElementById("exportPDFBtn")?.addEventListener("click", () => {
        let printContent = `
        <!DOCTYPE html>
        <html>
        <head><title>Laporan Zakat</title><meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            h1, h2 { color: #333; }
            .summary { display: flex; gap: 15px; flex-wrap: wrap; margin-bottom: 20px; }
            .summary-item { flex: 1; background: #f5f5f5; padding: 12px; border-radius: 8px; text-align: center; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background: #f0f0f0; }
            .footer { text-align: center; margin-top: 30px; font-size: 12px; color: #999; }
            @media print { body { margin: 0; padding: 10px; } }
        </style>
        </head>
        <body>
            <h1>BAITULMAAL - LAPORAN ZAKAT</h1>
            <p style="text-align:center">Tanggal: ${new Date().toLocaleDateString('id-ID')}</p>
            <div class="summary">
                <div class="summary-item"><strong>Total Uang Masuk</strong><br>${formatRupiah(totalUangMasuk())}</div>
                <div class="summary-item"><strong>Total Beras Masuk</strong><br>${totalBerasMasuk().toFixed(1)} kg</div>
                <div class="summary-item"><strong>Total Uang Tersalur</strong><br>${formatRupiah(totalUangTersalur())}</div>
                <div class="summary-item"><strong>Total Beras Tersalur</strong><br>${totalBerasTersalur().toFixed(1)} kg</div>
            </div>
            <h2>📋 DATA PENERIMAAN (${penerimaan.length})</h2>
            <table><thead><tr><th>No</th><th>Nama</th><th>Jenis</th><th>Detail</th><th>Uang (Rp)</th><th>Beras (kg)</th><th>Tgl</th></tr></thead>
            <tbody>${penerimaan.map((p,i)=>`<tr><td>${i+1}</td><td>${p.nama}</td><td>${p.jenis}</td><td>${p.detail}</td><td>${p.uang.toLocaleString()}</td><td>${p.beras.toFixed(1)}</td><td>${p.tanggal}</td></tr>`).join('')}${penerimaan.length===0?'<tr><td colspan="7">Belum ada data</td></tr>':''}</tbody></table>
            <h2>📦 DATA PENYALURAN (${penyaluran.length})</h2>
            <table><thead><tr><th>No</th><th>Tgl</th><th>Asnaf</th><th>Penerima</th><th>Uang (Rp)</th><th>Beras (kg)</th></tr></thead>
            <tbody>${penyaluran.map((p,i)=>`<tr><td>${i+1}</td><td>${p.tanggal}</td><td>${p.asnaf}</td><td>${p.penerima}</td><td>${p.uang.toLocaleString()}</td><td>${p.beras.toFixed(1)}</td></tr>`).join('')}${penyaluran.length===0?'<tr><td colspan="6">Belum ada data</td></tr>':''}</tbody></table>
            <div class="footer">Laporan dibuat oleh Sistem Baitulmaal | ${new Date().toLocaleString('id-ID')}</div>
        </body>
        </html>`;
        let printWindow = window.open('', '_blank');
        printWindow.document.write(printContent);
        printWindow.document.close();
        printWindow.print();
    });

    // Tab switching
    document.querySelectorAll(".tab-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            document.querySelectorAll(".tab-content").forEach(tc => tc.classList.remove("active"));
            document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
            if(btn.dataset.tab === "master") { renderMasterBeras(); renderMasterUang(); }
        });
    });

    renderMasterBeras();
    renderMasterUang();
    renderPenerimaan();
    renderPenyaluran();
    updateStats();
</script>
</body>
</html>