var SPREADSHEET_ID = '1oxNQ_oQwOVQN6F8ukZmjWt34GB42TwdRjI-x2SQGOYw';

function doGet(e) {
  // Cek apakah parameter action=getData dikirim
  if (!e || !e.parameter || e.parameter.action !== 'getData') {
    return HtmlService.createHtmlOutput(
      '<h1 style="text-align:center;margin-top:50px;font-family:sans-serif;">404 Halaman Tidak Ditemukan</h1>' +
      '<p style="text-align:center;font-family:sans-serif;">Anda tidak memiliki akses langsung ke API ini.</p>'
    );
  }

  // Mengembalikan data sebagai JSON API
  return ContentService.createTextOutput(getDashboardData())
    .setMimeType(ContentService.MimeType.JSON);
}

function findSheetByName(ss, targetName) {
  var target = targetName.toLowerCase().trim();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().trim() === target) {
      return sheets[i];
    }
  }
  return null;
}

function parseDateString(str) {
  if (!str) return null;
  var parts = str.toString().trim().split(/[\/\-]/);
  if (parts.length !== 3) return null;

  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var year = parseInt(parts[2], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  if (year < 100) year += 2000;

  var d = new Date(year, month - 1, day);
  if (isNaN(d.getTime())) return null;
  return d;
}

function formatDateShort(dateObj) {
  var d = dateObj.getDate();
  var m = dateObj.getMonth() + 1;
  var dd = d < 10 ? '0' + d : '' + d;
  var mm = m < 10 ? '0' + m : '' + m;
  return dd + '/' + mm + '/' + dateObj.getFullYear();
}

function formatDateIndo(dateObj) {
  var d = dateObj.getDate();
  var dd = d < 10 ? '0' + d : '' + d;
  var months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  return dd + ' ' + months[dateObj.getMonth()] + ' ' + dateObj.getFullYear();
}

function maskIdentifier(val, keepStart, keepEnd) {
  if (!val) return '';
  var s = val.toString().trim();
  if (!s) return '';
  if (s.length <= keepStart + keepEnd) {
    if (s.length <= 2) return '**';
    return s.substring(0, 1) + new Array(s.length - 1).join('*') + s.substring(s.length - 1);
  }
  var maskedLen = s.length - keepStart - keepEnd;
  return s.substring(0, keepStart) + new Array(maskedLen + 1).join('*') + s.substring(s.length - keepEnd);
}

function extractIssueCategories(masalahStr) {
  if (!masalahStr) return ['Tanpa Keterangan'];
  var raw = masalahStr.toString().trim();
  if (!raw) return ['Tanpa Keterangan'];

  var parts = raw.split('|');
  var seen = {};
  var categories = [];

  for (var i = 0; i < parts.length; i++) {
    var segment = parts[i].trim();
    if (!segment) continue;
    var colonIdx = segment.indexOf(':');
    var cat = colonIdx !== -1 ? segment.substring(0, colonIdx).trim() : segment;
    if (cat && !seen[cat]) {
      seen[cat] = true;
      categories.push(cat);
    }
  }

  return categories.length > 0 ? categories : ['Tanpa Keterangan'];
}

function getDashboardData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var dateSheets = [];
  // Mendukung format nama sheet tanggal DD/MM/YY maupun DD/MM/YYYY
  var dateRegex = /^\d{1,2}\/\d{1,2}\/(\d{2}|\d{4})$/;

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (dateRegex.test(sName)) {
      var parsed = parseDateString(sName);
      if (parsed) {
        dateSheets.push({
          name: sName,
          label: formatDateShort(parsed),
          dateObj: parsed,
          sheet: sheets[i]
        });
      }
    }
  }

  if (dateSheets.length === 0) {
    return JSON.stringify({
      error: 'Tidak ada sheet dengan format tanggal (DD/MM/YY atau DD/MM/YYYY) yang ditemukan'
    });
  }

  // Urutkan berdasarkan tanggal (terlama ke terbaru)
  dateSheets.sort(function(a, b) {
    return a.dateObj.getTime() - b.dateObj.getTime();
  });

  // Saring hanya sheet tanggal yang memiliki minimal 1 baris data selain header
  var validDateSheets = [];
  for (var v = 0; v < dateSheets.length; v++) {
    if (dateSheets[v].sheet.getLastRow() >= 2) {
      validDateSheets.push(dateSheets[v]);
    }
  }

  if (validDateSheets.length === 0) {
    return JSON.stringify({ error: 'Data pada sheet tanggal kosong.' });
  }

  var history = [];
  var latestRows = [];
  var latestSheetLabel = '';

  for (var s = 0; s < validDateSheets.length; s++) {
    // Gunakan getDisplayValues agar angka panjang (NIK) dan nol di depan (NISN) tetap utuh
    var sheetData = validDateSheets[s].sheet.getDataRange().getDisplayValues();
    if (sheetData.length < 2) continue;

    var rawHeaders = sheetData[0];
    var headers = [];
    for (var h = 0; h < rawHeaders.length; h++) {
      headers.push(rawHeaders[h] ? rawHeaders[h].toString().trim() : '');
    }

    var isLatest = (s === validDateSheets.length - 1);
    var rows = sheetData.slice(1);
    var schoolMap = {};
    var schoolList = [];
    var studentResult = [];

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var obj = {};
      var hasContent = false;

      for (var j = 0; j < headers.length; j++) {
        if (!headers[j]) continue;
        var val = r[j] ? r[j].toString().trim() : '';
        obj[headers[j]] = val;
        if (val !== '') hasContent = true;
      }

      // Abaikan baris kosong
      if (!hasContent) continue;

      var namaSekolah = obj['Nama Satuan Pendidikan'] || obj['Nama Sekolah'] || '';
      var npsn = obj['NPSN'] || '';
      var namaSiswa = obj['Nama Peserta Didik'] || '';
      if (!namaSekolah && !npsn && !namaSiswa) continue;

      var bentuk = obj['Bentuk Pendidikan'] || '';
      var status = obj['Status Sekolah'] || '';
      var kecamatan = obj['Kecamatan'] || '';
      var masalah = obj['Masalah Residu'] || '';

      if (isLatest) {
        // Masking NIK dan NISN sebelum dikirim lewat API (NPSN tetap tampil penuh)
        var outObj = {};
        for (var keyName in obj) {
          if (!obj.hasOwnProperty(keyName)) continue;
          var lowerKey = keyName.toLowerCase().trim();
          if (lowerKey === 'nik') {
            outObj[keyName] = maskIdentifier(obj[keyName], 4, 4);
          } else if (lowerKey === 'nisn') {
            outObj[keyName] = maskIdentifier(obj[keyName], 3, 2);
          } else {
            outObj[keyName] = obj[keyName];
          }
        }
        studentResult.push(outObj);
      }

      // Ringkasan per sekolah untuk tren historis agar payload tetap ringan
      var key = (npsn || namaSekolah) + '||' + kecamatan;
      if (!schoolMap[key]) {
        schoolMap[key] = {
          npsn: npsn,
          nama: namaSekolah,
          bentuk: bentuk,
          status: status,
          kecamatan: kecamatan,
          jumlahResidu: 0,
          issues: {}
        };
        schoolList.push(schoolMap[key]);
      }

      schoolMap[key].jumlahResidu += 1;

      var cats = extractIssueCategories(masalah);
      for (var c = 0; c < cats.length; c++) {
        var cat = cats[c];
        schoolMap[key].issues[cat] = (schoolMap[key].issues[cat] || 0) + 1;
      }
    }

    history.push({
      date: validDateSheets[s].label,
      sheetName: validDateSheets[s].name,
      schools: schoolList
    });

    if (isLatest) {
      latestRows = studentResult;
      latestSheetLabel = validDateSheets[s].label;
    }
  }

  if (history.length === 0) {
    return JSON.stringify({ error: 'Data pada sheet tanggal kosong.' });
  }

  var updateDateStr = '';
  var updateSheet = findSheetByName(ss, 'Update Data');
  if (updateSheet) {
    var updateValues = updateSheet.getDataRange().getValues();
    if (updateValues.length > 1) {
      var upHeaders = updateValues[0];
      var tglIndex = -1;
      for (var k = 0; k < upHeaders.length; k++) {
        if (upHeaders[k] && upHeaders[k].toString().toLowerCase().trim() === 'tanggal') {
          tglIndex = k;
          break;
        }
      }
      if (tglIndex !== -1) {
        var dateValue = updateValues[1][tglIndex];
        var parsedDate = null;

        if (dateValue instanceof Date) {
          parsedDate = dateValue;
        } else if (typeof dateValue === 'string') {
          parsedDate = parseDateString(dateValue) || new Date(dateValue);
        } else if (dateValue) {
          parsedDate = new Date(dateValue);
        }

        if (parsedDate && !isNaN(parsedDate.getTime())) {
          updateDateStr = formatDateIndo(parsedDate);
        } else if (dateValue) {
          updateDateStr = dateValue.toString().trim();
        }
      }
    }
  }

  // Fallback jika sheet Update Data kosong: pakai tanggal dari sheet terbaru
  if (!updateDateStr && validDateSheets.length > 0) {
    updateDateStr = formatDateIndo(validDateSheets[validDateSheets.length - 1].dateObj);
  }

  return JSON.stringify({
    data: latestRows,
    history: history,
    latestDate: latestSheetLabel,
    updateDate: updateDateStr
  });
}

function doPost(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: 'Request tidak valid' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var action = e.parameter.action;

  if (action === 'login') {
    var username = e.parameter.username ? e.parameter.username.toString().trim() : '';
    var password = e.parameter.password ? e.parameter.password.toString() : '';

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = findSheetByName(ss, 'Akun');

    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: 'Sheet Akun tidak ditemukan' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: 'Data akun kosong' })
      ).setMimeType(ContentService.MimeType.JSON);
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
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: 'Kolom Username/Password tidak ditemukan di sheet Akun' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var normInputUser = username.toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowUser = data[i][userIdx] ? data[i][userIdx].toString().trim() : '';
      var rowPass = data[i][passIdx] ? data[i][passIdx].toString() : '';

      if (rowUser && rowUser.toLowerCase() === normInputUser) {
        if (rowPass === password) {
          var role = roleIdx !== -1 && data[i][roleIdx] ? data[i][roleIdx].toString().trim() : 'User';
          return ContentService.createTextOutput(
            JSON.stringify({ success: true, role: role, username: rowUser })
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: 'Username atau Password salah' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ success: false, message: 'Invalid action' })
  ).setMimeType(ContentService.MimeType.JSON);
}
