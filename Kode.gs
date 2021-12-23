function doGet(){
  return HtmlService.createHtmlOutputFromFile('absen');
}

function ceknim(nim) {
var lembarabsen=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1')
var baristerakhir=lembarabsen.getLastRow()
var datasiswa=lembarabsen.getRange('A4:B'+baristerakhir).getValues()
var hasil="gagal"
for(x=0;x<baristerakhir-3;x++) {
  if(nim === datasiswa[x][0]){
   hasil={
     baris:x+4,
     nama:datasiswa[x][1]
   } 
  }
}
return hasil
}

function cekabsen() {
var lembarabsen=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1')
var datacentang=lembarabsen.getRange('C2:Z2').getValues()
var hasil="gagal"
for(x=0;x<24;x++) {
  if('v' === datacentang[0][x]){
   hasil=x +3
  }
}
return hasil
}

function isiabsen(baris,kolom){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1').getRange(baris, kolom).setValue('v')
}

function intiabsen(nim) {
  var hasil=''
  if(ceknim(nim) === 'gagal'){
    hasil ='Maaf, nomor pokok mahasiswa tidak ditemukan. Harap masukkan NPM dengan benar'  
  } else if(cekabsen() === 'gagal'){
    hasil ='Maaf, absen sedang ditutup. Silakan absensi dengan tepat waktu dan hubungi komisi advokasi'
  } else {
    var baris = ceknim(nim).baris
    var kolom = cekabsen()
    var nama = ceknim(nim).nama
     isiabsen(baris,kolom)
    hasil = 'Terima Kasih!' + (" ") + nama + (" ") + 'berhasil mengabsen'
  }
  return hasil
}
