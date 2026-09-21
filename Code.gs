function doGet(e) {
  // Cek apakah parameter action=getData dikirim
  if (!e || !e.parameter || e.parameter.action !== 'getData') {
    return HtmlService.createHtmlOutput('<h1 style="text-align:center;margin-top:50px;font-family:sans-serif;">404 Halaman Tidak Ditemukan</h1><p style="text-align:center;font-family:sans-serif;">Anda tidak memiliki akses langsung ke API ini.</p>');
  }

  // Mengembalikan data sebagai JSON API
  return ContentService.createTextOutput(getDashboardData())
    .setMimeType(ContentService.MimeType.JSON);
}

function getDashboardData() {
  var id = '1bOxsmyW3w_hQgtUmYHhXvQWLwb-VgZObu9mRluqcqbw';
  var ss = SpreadsheetApp.openById(id);
  var sheets = ss.getSheets();
  var dateSheets = [];
  var dateRegex = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
  
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    if (dateRegex.test(sName)) {
      dateSheets.push({ name: sName, sheet: sheets[i] });
    }
  }
  
  if (dateSheets.length === 0) {
    return { error: 'Tidak ada sheet dengan format tanggal (DD/MM/YYYY) yang ditemukan' };
  }
  
  // Sort by date (oldest first)
  dateSheets.sort(function(a, b) {
    var pA = a.name.split('/');
    var dA = new Date(pA[2], pA[1]-1, pA[0]);
    var pB = b.name.split('/');
    var dB = new Date(pB[2], pB[1]-1, pB[0]);
    return dA - dB;
  });
  
  var history = [];
  for (var s = 0; s < dateSheets.length; s++) {
    var sheetData = dateSheets[s].sheet.getDataRange().getValues();
    if (sheetData.length < 2) continue;
    
    var headers = sheetData[0];
    var rows = sheetData.slice(1);
    var result = [];
    
    for (var i = 0; i < rows.length; i++) {
      // Abaikan baris kosong
      if (!rows[i][0] && !rows[i][2]) continue;
      
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = rows[i][j];
      }
      result.push(obj);
    }
    
    history.push({
      date: dateSheets[s].name,
      data: result
    });
  }
  
  var updateDateStr = "";
  var updateSheet = ss.getSheetByName('update data');
  if (updateSheet) {
    var updateData = updateSheet.getDataRange().getValues();
    if (updateData.length > 0) {
      var upHeaders = updateData[0];
      var tglIndex = -1;
      for (var k = 0; k < upHeaders.length; k++) {
        if (upHeaders[k] && upHeaders[k].toString().toLowerCase() === 'tanggal') {
          tglIndex = k;
          break;
        }
      }
      if (tglIndex !== -1 && updateData.length > 1) {
        var dateValue = updateData[1][tglIndex];
        var parsedDate = null;
        
        if (dateValue instanceof Date) {
          parsedDate = dateValue;
        } else if (typeof dateValue === 'string') {
          var parts = dateValue.split(/[\/\-]/);
          if (parts.length === 3) {
            // Asumsi format DD/MM/YYYY
            parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
          } else {
            parsedDate = new Date(dateValue);
          }
        } else {
          parsedDate = new Date(dateValue);
        }

        if (parsedDate && !isNaN(parsedDate.getTime())) {
          var d = parsedDate.getDate();
          var dd = d < 10 ? '0' + d : d;
          var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
          updateDateStr = dd + " " + months[parsedDate.getMonth()] + " " + parsedDate.getFullYear();
        } else {
          updateDateStr = dateValue.toString();
        }
      }
    }
  }
  
  return JSON.stringify({ data: history[history.length - 1].data, history: history, updateDate: updateDateStr });
}

function doPost(e) {
  var action = e.parameter.action;
  
  if (action === 'login') {
    var username = e.parameter.username;
    var password = e.parameter.password;
    
    var id = '1bOxsmyW3w_hQgtUmYHhXvQWLwb-VgZObu9mRluqcqbw';
    var ss = SpreadsheetApp.openById(id);
    var sheet = ss.getSheetByName('akun');
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Sheet akun tidak ditemukan' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Data akun kosong' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var headers = data[0];
    var userIdx = -1;
    var passIdx = -1;
    var roleIdx = -1;
    
    for (var j = 0; j < headers.length; j++) {
      var h = headers[j] ? headers[j].toString().toLowerCase().trim() : '';
      if (h === 'username') userIdx = j;
      if (h === 'password') passIdx = j;
      if (h === 'role') roleIdx = j;
    }
    
    if (userIdx === -1 || passIdx === -1) {
       return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Kolom Username/Password tidak ditemukan di sheet akun' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][userIdx] && data[i][userIdx].toString() === username) {
        if (data[i][passIdx] && data[i][passIdx].toString() === password) {
          var role = roleIdx !== -1 && data[i][roleIdx] ? data[i][roleIdx].toString() : 'user';
          return ContentService.createTextOutput(JSON.stringify({ success: true, role: role })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Username atau Password salah' })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Invalid action' })).setMimeType(ContentService.MimeType.JSON);
}
