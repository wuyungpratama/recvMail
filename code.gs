const ss = SpreadsheetApp.openById(
  "1pbKWbsD8wlxcAX-iXnP3Pez25Kaw7Wp9ZccwwkE52SU"
);
const sheet = ss.getSheetByName("Data");
const FOLDER_NAME = "1Svdp0DXFOns_S3D5yyUe-4k5I4hC1h9_";

// Akses halaman utama
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index.html")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

// Ambil semua data dari Sheet
function getData() {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows
    .filter((row) => row[0]) // abaikan baris kosong
    .map((row) => ({
      id: row[0] || "",
      penerima: row[1] || "",
      noUrut: row[2] || "",
      asalSurat: row[3] || "",
      nomorSurat: row[4] || "",
      tglSurat: row[5]
        ? Utilities.formatDate(
            new Date(row[5]),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          )
        : "",
      tglTerimaSurat: row[6]
        ? Utilities.formatDate(
            new Date(row[6]),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          )
        : "",
      perihal: row[7] || "",
      disposisi: row[8] || "",
      fileUrls: row[9] || "",
    }));
}

// Tambah data baru
function addData(formData, files) {
  // Generate ID otomatis
  const lastRow = sheet.getLastRow();
  const lastId =
    lastRow > 1 ? parseInt(sheet.getRange(lastRow, 1).getValue()) : 0;
  const newId = lastId + 1;

  // Upload file dan simpan link
  const fileLinks = uploadFiles(newId, files) || [];

  // Susun data sesuai header
  const row = [
    newId,
    formData.penerima,
    formData.noUrut,
    formData.asalSurat,
    formData.nomorSurat,
    formData.tglSurat,
    formData.tglTerimaSurat,
    formData.perihal,
    formData.disposisi,
    fileLinks.join(", "), // gabungkan semua link
  ];

  sheet.appendRow(row);
  return { status: "success", id: newId };
}

// Update data berdasarkan rowIndex
function updateData(formData, files) {
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(
    (row) => row[0].toString() == formData.id.toString()
  );
  if (rowIndex === -1) throw new Error("ID tidak ditemukan");

  const actualRow = rowIndex + 1;

  let fileLinks = [];
  if (files && files.length > 0) {
    // Hapus file lama
    const oldLinks = sheet.getRange(actualRow, 10).getValue();
    const oldIds = (oldLinks || "")
      .split(",")
      .map((link) => {
        const match = link.match(/[-\w]{25,}/);
        return match ? match[0] : null;
      })
      .filter((id) => id);

    oldIds.forEach((id) => {
      try {
        DriveApp.getFileById(id).setTrashed(true);
      } catch (e) {
        Logger.log("Gagal hapus file: " + id + " → " + e.message);
      }
    });

    // Upload baru
    const id = formData.id;
    fileLinks = uploadFiles(id, files);
  } else {
    // Tetap pakai link lama
    const oldLink = sheet.getRange(actualRow, 10).getValue();
    fileLinks = [oldLink];
  }

  const row = [
    formData.id,
    formData.penerima,
    formData.noUrut,
    formData.asalSurat,
    formData.nomorSurat,
    formData.tglSurat,
    formData.tglTerimaSurat,
    formData.perihal,
    formData.disposisi,
    fileLinks.join(", "),
  ];

  sheet.getRange(actualRow, 1, 1, row.length).setValues([row]);
  return { status: "updated" };
}

// Hapus data dan file terkait berdasarkan rowIndex
function deleteData(rowIndex) {
  // Ambil URL file dari data yang akan dihapus
  const fileUrls = sheet
    .getRange(rowIndex + 1, 10)
    .getValue()
    .split(", "); // Kolom 17 untuk file URLs

  // Hapus file-file terkait di Google Drive
  fileUrls.forEach((url) => {
    const fileId = url.match(/[-\w]{25,}/); // Ambil ID file dari URL
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId[0]);
        file.setTrashed(true); // Pindahkan ke trash
      } catch (e) {
        Logger.log("Gagal hapus file: " + e.message);
      }
    }
  });

  // Hapus baris data di Google Sheets
  sheet.deleteRow(rowIndex + 1); // Menghapus data yang benar

  return { status: "deleted" };
}

// upload files
function uploadFiles(id, files) {
  if (!files || files.length === 0) return []; // ⬅️ Tambahan penting
  const folder = DriveApp.getFolderById(FOLDER_NAME);
  const links = [];

  files.forEach((file) => {
    try {
      if (!file.data) {
        throw new Error(`File data kosong untuk file: ${file.name}`);
      }

      const decodedData = Utilities.base64Decode(file.data);
      const blob = Utilities.newBlob(
        decodedData,
        file.type || "application/octet-stream",
        file.name
      );
      const uploadedFile = folder.createFile(blob);

      // Ganti nama file dengan ID agar lebih unik
      uploadedFile.setName(`${id}_${file.name}`);

      links.push(uploadedFile.getUrl());
    } catch (e) {
      Logger.log(`Gagal upload file ${file.name}: ${e.message}`);
    }
  });

  return links;
}

// Buat folder upload jika belum ada
function getOrCreateFolder(name) {
  const parent = DriveApp.getRootFolder();
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}
