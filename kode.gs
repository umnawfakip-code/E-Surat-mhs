const SPREADSHEET_ID = '1mlVSLb1BS_YDkynqHIWrfXnaGk5zOoiom1KkZlS-jeM';
const OUTPUT_FOLDER_ID = '1Qh5hmlDDR_lvGJYT04h9oQ3kxjTy7-7X';
const ADMIN_PIN = 'adminfkip'; // PIN LOGIN ADMIN

const TEMPLATE_IDS = {
  'AKTIF KULIAH': '1c9XCdUspIi6CTPDckIFGpyqoqlBOly-kwIUD0XOWovo',
  'OBSERVASI': '1NLg8ajmWwLaRDM-FCb4p-cDalVrBCAY203Jvgiw4LUc',
  'IZIN PENELITIAN': '1ifjBgwGEaghhVLJQdQ1yo56hfCmVBV_ha1sTEsg1THQ',
  'UNDANGAN SEMINAR': '1wQcnaTCVa8qVKN9JxUHay2PoIgrTYBQ_s90bkuRjgkQ',
  'TANDA TERIMA SKRIPSI': '1skU9S-QTI_tQA6qSpxpX38boX8cKdGf0F83sYNIn9B4',
  'SKBB': '1gwe_uqYMVjKxTyMiKMhCTD6rIqYtzC3uLoPspPrDL0s'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('E-Surat FKIP UMN Al-Washliyah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSystemInfo() {
  return { limit: MailApp.getRemainingDailyQuota() };
}

function verifyAdmin(pin) {
  return pin === ADMIN_PIN;
}

function getAdminData(letterType) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = null;
    for (let s of ss.getSheets()) {
      if (s.getName().toUpperCase() === letterType.toUpperCase()) {
        sheet = s;
        break;
      }
    }
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    let results = [];
    let start = Math.max(1, data.length - 100); 
    
    for (let i = data.length - 1; i >= start; i--) {
      let cleanData = data[i].map(cell => {
        if (cell instanceof Date) {
          return cell.getFullYear() + '-' + String(cell.getMonth() + 1).padStart(2, '0') + '-' + String(cell.getDate()).padStart(2, '0');
        }
        return cell !== null && cell !== undefined ? cell.toString().trim() : '';
      });
      results.push({ row: i + 1, data: cleanData });
    }
    return results;
  } catch (error) {
    Logger.log("Error Admin Data: " + error.toString());
    return [];
  }
}

function deleteAdminData(letterType, rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = null;
    for (let s of ss.getSheets()) {
      if (s.getName().toUpperCase() === letterType.toUpperCase()) {
        sheet = s;
        break;
      }
    }
    if (!sheet) throw new Error("Sheet tidak ditemukan");
    
    sheet.deleteRow(rowNumber);
    return { success: true, message: "Data baris ke-" + rowNumber + " berhasil dihapus." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function formatTanggalIndo(tanggalMasuk) {
  if (!tanggalMasuk) return '-';
  const bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const d = new Date(tanggalMasuk);
  if (isNaN(d.getTime())) return tanggalMasuk;
  return d.getDate() + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
}

function processSubmission(letterType, formData, filesData, editRowNumber = null) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = null;
    for (let s of ss.getSheets()) { 
      if (s.getName().toUpperCase() === letterType.toUpperCase()) { 
        sheet = s; 
        break;
      } 
    }
    if (!sheet) throw new Error("Sheet tidak ditemukan: " + letterType);
    
    const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    const uploadedUrls = {};
    let oldData = [];
    
    if (editRowNumber) {
      oldData = sheet.getRange(editRowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    }

    if (filesData && Object.keys(filesData).length > 0) {
      for (let key in filesData) {
        if (filesData[key]) {
          const fileBlob = Utilities.newBlob(Utilities.base64Decode(filesData[key].base64), filesData[key].mimeType, filesData[key].name);
          uploadedUrls[key] = folder.createFile(fileBlob).getUrl();
        }
      }
    }

    let rowData = [];
    let replacements = {};
    let recipientEmail = '';
    let fileNameBase = '';
    const tanggalSuratIndo = formatTanggalIndo(formData.tanggal);

    // ======================== SUSUNAN FIELD SPREADSHEET KONSISTEN ========================
    if (letterType === 'AKTIF KULIAH') {
      recipientEmail = formData.email;
      fileNameBase = `Surat_Aktif_Kuliah_${formData.npm}_${formData.nama}`;
      const urlSpp = uploadedUrls['slipSpp'] || (editRowNumber ? oldData[10] : '');
      const urlKtm = uploadedUrls['ktm'] || (editRowNumber ? oldData[11] : '');
      
      rowData = [formData.nama, formData.nohp, formData.email, formData.npm, formData.ttl, formData.jurusan, formData.prodi, formData.semester, formData.alamat, formData.tanggal, urlSpp, urlKtm];
      replacements = {'Nama': formData.nama, 'No. HP': formData.nohp, 'Email Aktif': formData.email, 'NPM': formData.npm, 'Tempat Tanggal Lahir': formData.ttl, 'ttl': formData.ttl, 'Jurusan': formData.jurusan, 'Program Studi': formData.prodi, 'prodi': formData.prodi, 'Semester': formData.semester, 'Alamat': formData.alamat, 'Tanggal Surat': tanggalSuratIndo};
    } 
    else if (letterType === 'OBSERVASI') {
      recipientEmail = formData.email;
      fileNameBase = `Surat_Observasi_${formData.npmKoor}_${formData.namaKoor}`;
      
      rowData = [formData.namaKoor, formData.npmKoor, formData.nama2, formData.npm2, formData.nama3, formData.npm3, formData.nama4, formData.npm4, formData.nama5, formData.npm5, formData.nama6, formData.npm6, formData.nama7, formData.npm7, formData.nama8, formData.npm8, formData.nama9, formData.npm9, formData.nama10, formData.npm10, formData.nohpKoor, formData.email, formData.tujuan, formData.tanggal, formData.tugas, formData.dosen];
      replacements = {'Nama Koordinator': formData.namaKoor, 'NPM Koordinator': formData.npmKoor, 'No. HP Koordinator': formData.nohpKoor, 'Email Aktif': formData.email, 'Tujuan Surat': formData.tujuan, 'Tanggal Surat': tanggalSuratIndo, 'Tugas Observasi': formData.tugas, 'Dosen Pembimbing': formData.dosen};
      for(let i=2; i<=10; i++) { 
        replacements[`Nama ${i}`] = formData[`nama${i}`] ? formData[`nama${i}`] : ''; 
        replacements[`NPM ${i}`] = formData[`npm${i}`] ? formData[`npm${i}`] : '';
      }
    } 
    else if (letterType === 'IZIN PENELITIAN') {
      recipientEmail = formData.email;
      fileNameBase = `Surat_Izin_Penelitian_${formData.npm}_${formData.nama}`;
      
      rowData = [formData.nama, formData.npm, formData.jurusan, formData.prodi, formData.tujuan, formData.tanggal, formData.judul, formData.nohp, formData.email];
      replacements = {'Nama': formData.nama, 'NPM': formData.npm, 'Jurusan': formData.jurusan, 'Program Studi': formData.prodi, 'Tujuan Surat': formData.tujuan, 'Tanggal Surat': tanggalSuratIndo, 'Judul': formData.judul, 'No. HP': formData.nohp, 'Email Aktif': formData.email};
    } 
    else if (letterType === 'UNDANGAN SEMINAR') {
      recipientEmail = formData.email;
      fileNameBase = `Undangan_Seminar_${formData.npm}_${formData.nama}`;
      const urlSlip = uploadedUrls['slipSeminar'] || (editRowNumber ? oldData[11] : '');
      const urlForm = uploadedUrls['formF'] || (editRowNumber ? oldData[12] : '');
      
      rowData = [formData.nama, formData.nowa, formData.email, formData.npm, formData.jurusan, formData.prodi, formData.kaprodi, formData.judul, formData.pembimbing, formData.penguji1, formData.penguji2, urlSlip, urlForm, formData.tanggal];
      replacements = {'Nama Mahasiswa': formData.nama, 'No. WhatsApp': formData.nowa, 'Email Aktif': formData.email, 'NPM': formData.npm, 'Jurusan': formData.jurusan, 'Program Studi': formData.prodi, 'Ka.Prodi': formData.kaprodi, 'Judul Skripsi': formData.judul, 'Pembimbing': formData.pembimbing, 'Penguji 1': formData.penguji1, 'Penguji 2': formData.penguji2, 'Tanggal Surat': tanggalSuratIndo};
    } 
    else if (letterType === 'TANDA TERIMA SKRIPSI') {
      recipientEmail = formData.email; 
      fileNameBase = `Tanda_Terima_Skripsi_${formData.npm}_${formData.nama}`;
      
      rowData = [formData.nama, formData.npm, formData.prodi, formData.pembimbing, formData.judul, formData.tglSidang, formData.judulArtikel, formData.jenisArtikel, formData.linkArtikel, formData.tglTerbit, formData.tanggal, formData.email];
      replacements = {'Nama Mahasiswa': formData.nama, 'NPM': formData.npm, 'Program Studi': formData.prodi, 'Pembimbing': formData.pembimbing, 'Judul Skripsi': formData.judul, 'Tanggal Sidang': formatTanggalIndo(formData.tglSidang), 'Judul Artikel Ilmiah': formData.judulArtikel, 'Jenis Artikel Ilmiah': formData.jenisArtikel, 'Link Artikel Ilmiah': formData.linkArtikel, 'Tgl. Terbit Artikel Ilmiah': formatTanggalIndo(formData.tglTerbit), 'Tanggal Surat': tanggalSuratIndo, 'Email Aktif': formData.email};
    }
    else if (letterType === 'SKBB') {
      recipientEmail = formData.email;
      fileNameBase = `Surat_SKBB_${formData.npm}_${formData.nama}`;
      
      rowData = [formData.nama, formData.email, formData.npm, formData.semester, formData.prodi, formData.kaprodi, formData.nidn, formData.tanggal];
      replacements = {'Nama': formData.nama, 'Email Aktif': formData.email, 'NPM': formData.npm, 'Semester': formData.semester, 'Program Studi': formData.prodi, 'Ka. Prodi': formData.kaprodi, 'NIDN': formData.nidn, 'Tanggal Surat': tanggalSuratIndo};
    }

    // ======================== PROSES MANIPULASI TEMPLATE DOKUMEN ========================
    const templateId = TEMPLATE_IDS[letterType];
    const copiedDocFile = DriveApp.getFileById(templateId).makeCopy(fileNameBase, folder);
    const doc = DocumentApp.openById(copiedDocFile.getId());
    const body = doc.getBody();
    
    // PENGGANTIAN TEKS BERSIH: Mencari ekspresi persis "{Placeholder}" bawaan template Docs
    for (let placeholder in replacements) {
      let clearText = replacements[placeholder];
      if (Array.isArray(clearText)) clearText = clearText.join(', ');
      if (clearText === null || clearText === undefined) clearText = '';
      
      // Bersihkan isi variabel dari karakter kurung kurawal yang mungkin tidak sengaja terinput
      clearText = clearText.toString().replace(/[\{\}\[\]]/g, "").trim();
      
      // Eksekusi replaceText langsung dengan menyertakan pembungkus kurung kurawalnya agar ikut terhapus resmi
      body.replaceText("{" + placeholder + "}", clearText);
    }
    
    // Amankan sisa penanda huruf kecil khusus RPS (japa/prodi jika ada)
    if(formData.prodi) body.replaceText("{prodi}", formData.prodi);
    if(formData.ttl) body.replaceText("{ttl}", formData.ttl);

    // KUNCI UTAMA: Simpan dan tutup dokumen untuk memaksa Google Docs me-refresh cache memori internal dokumennya
    doc.saveAndClose();

    // ======================== KONVERSI DAN OUTPUT BERKAS ========================
    let finalFileUrl = "";
    if (letterType === 'UNDANGAN SEMINAR') {
      finalFileUrl = copiedDocFile.getUrl();
    } else {
      // Buka ulang file salinan dalam mode blob PDF setelah dipastikan memori tersimpan bersih
      const pdfBlob = DriveApp.getFileById(copiedDocFile.getId()).getBlob().getAs(MimeType.PDF);
      const pdfFile = folder.createFile(pdfBlob).setName(fileNameBase + '.pdf');
      copiedDocFile.setTrashed(true); // Hapus Doc dummy transaksi
      finalFileUrl = pdfFile.getUrl();
    }

    let labelStatus = (letterType === 'UNDANGAN SEMINAR') ? "Tersimpan" : "Terkirim";

    // Manajemen Pengiriman Email Otomatis
    if (MailApp.getRemainingDailyQuota() > 5) {
      MailApp.sendEmail({
        to: recipientEmail,
        subject: `E-Surat FKIP UMN Al-Washliyah: ${fileNameBase.replace(/_/g, ' ')}`,
        body: `Halo,\n\nBerikut berkas pengajuan surat Anda yang berhasil diproses secara otomatis oleh sistem.\n\nTerima Kasih.\nFKIP UMN Al-Washliyah`,
        attachments: (letterType === 'UNDANGAN SEMINAR') ? [copiedDocFile.getBlob()] : [DriveApp.getFileById(finalFileUrl.match(/[-\w]{25,}/)[0]).getBlob()]
      });
    } else {
      catatAntrean(recipientEmail, letterType, finalFileUrl);
      labelStatus = (letterType === 'UNDANGAN SEMINAR') ? "Antrean Simpan (Limit)" : "Antrean Kirim (Limit)";
    }

    // Dorong data pelaporan ke Google Sheets
    rowData.push(finalFileUrl); 
    rowData.push(labelStatus);  

    if (editRowNumber) {
      const targetRange = sheet.getRange(editRowNumber, 1, 1, rowData.length);
      targetRange.clearContent();
      targetRange.setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, url: finalFileUrl, emailStatus: labelStatus };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function catatAntrean(email, letterType, fileUrl) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let queueSheet = ss.getSheetByName("AntreanEmail");
  if (!queueSheet) { 
    queueSheet = ss.insertSheet("AntreanEmail"); 
    queueSheet.appendRow(["Tanggal Antrean", "Email", "Jenis Surat", "URL File", "Status"]); 
  }
  queueSheet.appendRow([new Date(), email, letterType, fileUrl, "Pending"]);
}

function prosesAntreanEmail() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const queueSheet = ss.getSheetByName("AntreanEmail");
  if (!queueSheet) return;
  const data = queueSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === "Pending" && MailApp.getRemainingDailyQuota() > 5) {
      try {
        const fileId = data[i][3].match(/[-\w]{25,}/); 
        const fileBlob = DriveApp.getFileById(fileId).getBlob();
        MailApp.sendEmail({
          to: data[i][1],
          subject: `[Antrean] E-Surat FKIP UMN Al-Washliyah: ${data[i][2]}`,
          body: `Halo, berikut berkas surat Anda yang sebelumnya tertunda. Terima kasih.`,
          attachments: [fileBlob]
        });
        queueSheet.getRange(i + 1, 5).setValue("Berhasil");
      } catch (e) {
        queueSheet.getRange(i + 1, 5).setValue("Gagal: " + e.toString());
      }
    }
  }
}
