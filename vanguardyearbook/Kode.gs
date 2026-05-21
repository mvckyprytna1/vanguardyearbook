function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Buku Kenangan Pondok Pesantren')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsInfo = [
    {
      name: "Siswa",
      headers: ["id", "nama", "angkatan", "quotes", "foto_url", "ttl", "alamat", "nohp", "ig"],
      data: [
        ["S-001", "Ahmad Faris", "2026", "Ilmu adalah cahaya.", "https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?w=400&q=80", "Jambi, 15 Mei 2008", "Jl. Sudirman No. 12, Jambi", "081234567890", "@ahmad.faris"],
        ["S-002", "Siti Aisyah", "2026", "Sebaik-baik manusia adalah yang bermanfaat.", "https://images.unsplash.com/photo-1544005313-94ddf0286df2?w=400&q=80", "Padang, 20 Agustus 2008", "Jl. Merpati No. 5, Padang", "089876543210", "@aisyah.siti"]
      ]
    },
    {
      name: "Guru",
      headers: ["id", "nama", "jabatan", "foto_url", "nip", "bio"],
      data: [
        ["G-001", "Ust. Abdul Somad, Lc.", "Wali Kelas A", "https://images.unsplash.com/photo-1472099645785-5658abf4ff4e?w=400&q=80", "198005152005011003", "Lulusan Universitas Al-Azhar Kairo. Mengampu mata pelajaran Tafsir dan Hadits."],
        ["G-002", "Ustzh. Halimah, S.Pd.", "Wali Kelas B", "https://images.unsplash.com/photo-1531123897727-8f129e1bf08c?w=400&q=80", "198508202010012005", "Ahli dalam bidang Fiqih Wanita dan pengasuh asrama putri."]
      ]
    },
    {
      name: "PesanKesan",
      headers: ["id", "pengirim", "pesan", "tanggal", "status"],
      data: [
        ["P-001", "Alumni 2020", "Terima kasih atas ilmu yang diberikan. Jaya selalu pondokku!", new Date().toISOString(), "Approved"]
      ]
    },
    {
      name: "Galeri",
      headers: ["id", "judul", "kategori", "foto_url", "deskripsi"],
      data: [
        ["GL-001", "Gedung Utama", "Infrastruktur", "https://images.unsplash.com/photo-1584551246679-0daf3d275d0f?w=800&q=80", "Gedung utama tempat berlangsungnya kajian akbar dan pusat kegiatan santri."],
        ["GL-002", "Kegiatan Mengaji", "Aktivitas", "https://images.unsplash.com/photo-1574345311409-1a48c6a084bc?w=800&q=80", "Kegiatan rutinan setoran hafalan Al-Qur'an setiap ba'da Subuh."]
      ]
    }
  ];

  sheetsInfo.forEach(info => {
    let sheet = ss.getSheetByName(info.name);
    if (!sheet) {
      sheet = ss.insertSheet(info.name);
      sheet.getRange(1, 1, 1, info.headers.length).setValues([info.headers]).setFontWeight("bold");
      if (info.data.length > 0) {
        sheet.getRange(2, 1, info.data.length, info.headers.length).setValues(info.data);
      }
    }
  });
  return "Database berhasil disetup!";
}

function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    siswa: getSheetData(ss.getSheetByName("Siswa")),
    guru: getSheetData(ss.getSheetByName("Guru")),
    pesan: getSheetData(ss.getSheetByName("PesanKesan")),
    galeri: getSheetData(ss.getSheetByName("Galeri"))
  };
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function submitPesan(formData) {
  try {
    if (!formData.pengirim || !formData.pesan) throw new Error("Data tidak lengkap!");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PesanKesan");
    const id = "P-" + new Date().getTime();
    const tanggal = new Date().toISOString();
    
    sheet.appendRow([id, formData.pengirim, formData.pesan, tanggal, "Approved"]);
    return { success: true, message: "Pesan berhasil dikirim!" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}