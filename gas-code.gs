function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    
    // Auto-create Header jika sheet masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Nama",
        "Email",
        "Perusahaan",
        "Role / Jabatan",
        "Overall Score",
        "Overall Level",
        "Evidence Orientation",
        "Data Seeking & Curiosity",
        "Cognitive Flexibility",
        "Reflection & Learning",
        "Kekuatan Utama",
        "Prioritas Pengembangan"
      ]);
    }
    
    // Tambah Baris Data Peserta Baru
    sheet.appendRow([
      new Date(),
      data.nama || "-",
      data.email || "-",
      data.company || "-",
      data.role || "-",
      data.overallScore || "0.00",
      data.overallLevel || "-",
      data.evidenceOrientation || "0.00",
      data.dataSeekingCuriosity || "0.00",
      data.cognitiveFlexibility || "0.00",
      data.reflectionLearning || "0.00",
      data.kekuatanUtama || "-",
      data.prioritasPengembangan || "-"
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
