function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Data Responses");
    
    if (!sheet) {
      const firstSheet = ss.getSheets()[0];
      if (firstSheet.getName() === "Sheet1") {
        firstSheet.setName("Data Responses");
        sheet = firstSheet;
      } else {
        sheet = ss.insertSheet("Data Responses");
      }
    }
    
    const data = JSON.parse(e.postData.contents);
    
    // Auto-create Header jika sheet data masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Nama",
        "Email",
        "Perusahaan",
        "Role / Job Level",
        "Tahap Evaluasi",
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
      data.stage || "Pre-Test",
      Number(data.overallScore) || 0,
      data.overallLevel || "-",
      Number(data.evidenceOrientation) || 0,
      Number(data.dataSeekingCuriosity) || 0,
      Number(data.cognitiveFlexibility) || 0,
      Number(data.reflectionLearning) || 0,
      data.kekuatanUtama || "-",
      data.prioritasPengembangan || "-"
    ]);
    
    // Update Cohort Analytics Dashboard
    updateCohortDashboard(ss);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function updateCohortDashboard(ss) {
  let dash = ss.getSheetByName("Cohort Analytics");
  if (!dash) {
    dash = ss.insertSheet("Cohort Analytics");
  }
  
  dash.clear(); // Refresh layout & formula
  
  // Header Title
  dash.getRange("A1:G1").merge().setValue("🏢 MINDSHIFT COHORT & ORGANIZATIONAL INSIGHT DASHBOARD")
      .setFontWeight("bold").setFontSize(16).setBackground("#172554").setFontColor("#ffffff")
      .setHorizontalAlignment("center");

  // Executive Cards
  dash.getRange("A3").setValue("Total Responden").setFontWeight("bold");
  dash.getRange("B3").setValue("=COUNTA('Data Responses'!B2:B)");
  dash.getRange("A4").setValue("Rata-rata Skor Cohort").setFontWeight("bold");
  dash.getRange("B4").setValue("=AVERAGE('Data Responses'!G2:G)").setNumberFormat("0.00");

  // Pre vs Post Comparison Section
  dash.getRange("A6:D6").merge().setValue("🔄 PRE vs POST PROGRAM COMPARISON")
      .setFontWeight("bold").setBackground("#e0e7ff").setFontColor("#1e3a8a");
  dash.getRange("A7").setValue("Dimensi Mindset").setFontWeight("bold");
  dash.getRange("B7").setValue("Pre-Test Avg").setFontWeight("bold");
  dash.getRange("C7").setValue("Post-Test Avg").setFontWeight("bold");
  dash.getRange("D7").setValue("Mindset Shift (Delta)").setFontWeight("bold");

  const dims = [
    ["Overall Score", "G"],
    ["Evidence Orientation", "I"],
    ["Data Seeking & Curiosity", "J"],
    ["Cognitive Flexibility", "K"],
    ["Reflection & Learning", "L"]
  ];

  dims.forEach((d, idx) => {
    let row = 8 + idx;
    dash.getRange(row, 1).setValue(d[0]);
    dash.getRange(row, 2).setValue(`=AVERAGEIFS('Data Responses'!${d[1]}2:${d[1]}, 'Data Responses'!F2:F, "Pre-Test")`).setNumberFormat("0.00");
    dash.getRange(row, 3).setValue(`=AVERAGEIFS('Data Responses'!${d[1]}2:${d[1]}, 'Data Responses'!F2:F, "Post-Test")`).setNumberFormat("0.00");
    dash.getRange(row, 4).setValue(`=IF(ISBLANK(C${row}), 0, C${row} - B${row})`).setNumberFormat("+0.00;-0.00;0.00");
  });

  // Maturity Level Distribution Section
  dash.getRange("F6:G6").merge().setValue("🏆 LEVEL DISTRIBUTION")
      .setFontWeight("bold").setBackground("#e0e7ff").setFontColor("#1e3a8a");
  dash.getRange("F7").setValue("Level 1 — Data Novice");
  dash.getRange("G7").setValue('=COUNTIF(\'Data Responses\'!H2:H, "*Data Novice*")');
  dash.getRange("F8").setValue("Level 2 — Data Aware");
  dash.getRange("G8").setValue('=COUNTIF(\'Data Responses\'!H2:H, "*Data Aware*")');
  dash.getRange("F9").setValue("Level 3 — Practitioner");
  dash.getRange("G9").setValue('=COUNTIF(\'Data Responses\'!H2:H, "*Practitioner*")');
  dash.getRange("F10").setValue("Level 4 — Data-Driven Leader");
  dash.getRange("G10").setValue('=COUNTIF(\'Data Responses\'!H2:H, "*Leader*")');
}
