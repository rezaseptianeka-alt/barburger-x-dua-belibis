var sheetName = "Sheet1";
var folderId = "MASUKKAN_ID_FOLDER_GOOGLE_DRIVE_ANDA_DISINI"; // Ganti dengan ID folder Drive Anda

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(sheetName);
    var data = e.parameter;
    
    var fotoUrl = "";
    
    // Cek apakah ada file foto yang dikirimkan dalam format Base64
    if (data.fotoData && data.fotoNama) {
      var folder = DriveApp.getFolderById(folderId);
      
      // Pisahkan header base64 (misal: "data:image/jpeg;base64,...")
      var splitBase64 = data.fotoData.split(",");
      var contentType = splitBase64[0].match(/:(.*?);/)[1];
      var decodedBytes = Utilities.base64Decode(splitBase64[1]);
      
      // Buat file blob di Google Drive
      var blob = Utilities.newBlob(decodedBytes, contentType, data.fotoNama);
      var file = folder.createFile(blob);
      
      // Set agar file bisa dilihat oleh siapa saja yang memiliki link (opsional)
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // Ambil tautan (URL) file untuk dimasukkan ke spreadsheet
      fotoUrl = file.getUrl();
    }
    
    // Menyusun data sesuai urutan kolom spreadsheet
    var nextRow = sheet.getLastRow() + 1;
    var rowData = [
      new Date(),
      data.nama || "",
      fotoUrl, // Masukkan URL file foto dari Drive
      data.telepon || "",
      data.instagram || "",
      data.id_driver || ""
    ];
    
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "row": nextRow }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}