function doPost(e) {
try {
var data = e.parameter.fileContent;
var filename = e.parameter.filename;
var email = e.parameter.email;
var nama = e.parameter.nama;
var result=uploadFileToGoogleDrive(data,filename,nama,email,e);
return ContentService // kembali ke hasil JSON behasil
.createTextOutput(
JSON.stringify({"Anda berhasil mendaftar":"Data telah terkirim",
"data": JSON.stringify(result) }))
.setMimeType(ContentService.MimeType.JSON);
} catch(error) { // kembali ke sini jika error
Logger.log(error);
return ContentService
.createTextOutput(JSON.stringify({"result":"error", "error": error}))
.setMimeType(ContentService.MimeType.JSON);
}
}
// new property service GLOBAL
var SCRIPT_PROP = PropertiesService.getScriptProperties();
// see: https://developers.google.com/apps-script/reference/properties/
/**
* pilih sheet
*/
function setup() {
var doc = SpreadsheetApp.getActiveSpreadsheet();
SCRIPT_PROP.setProperty("xxx", doc.getId());
}
/**
* record_data adalah insert data yang diterima dari submisi HTML form
* e adalah data yang diterima dari POST
*/
function record_data(e,fileUrl) {
try {
var doc = SpreadsheetApp.openById('xxx');
var sheet = doc.getSheetByName('Data'); // pilih sheet respon
var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
var nextRow = sheet.getLastRow()+1; // dapatkan baris selanjutnya
var row = [ new Date() ]; // element pertama pada baris harus selalu diawali dengan timestamp
// loop through the header columns
for (var i = 1; i < headers.length; i++) { // start pada 1 untuk menghindari kolom timestamp
if(headers[i].length > 0 && headers[i] == "file") {
row.push(fileUrl); // tambah data ke baris
}
else if(headers[i].length > 0) {
row.push(e.parameter[headers[i]]); // tambah data ke baris
}
}
// more efficient to set values as [][] array than individually
sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
}
catch(error) {
Logger.log(e);
}
finally {
return;
}
}
function uploadFileToGoogleDrive(data, file, nama, email,e) {
try {
var dropbox = "REGISTRASI";
var folder, folders = DriveApp.getFoldersByName(dropbox);
if (folders.hasNext()) {
folder = folders.next();
} else {
folder = DriveApp.createFolder(dropbox);
}
var contentType = data.substring(5,data.indexOf(';')),
bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)),
blob = Utilities.newBlob(bytes, contentType, file);
var file = folder.createFile(blob);
var fileUrl=file.getUrl();
record_data(e,fileUrl);

return file.getUrl();
} catch (f) {
return ContentService // kembali ke hasil JSON behasil.
.createTextOutput(
JSON.stringify({"Maaf!":"Upload data gagal!",
"data": JSON.stringify(f) }))
.setMimeType(ContentService.MimeType.JSON);
}
}













