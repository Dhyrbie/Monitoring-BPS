// ==========================================
// 1. HELPER & FORMATTER
// ==========================================
function parseDateValue(val) {
  if (!val || String(val).trim() === "") return null;
  if (val instanceof Date) return val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateISO(d) {
  if (!d) return "";
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + "-" + m + "-" + day;
}

function getBulanPendataan(tanggalStr) {
  if (!tanggalStr) return "Lainnya";
  var date = new Date(tanggalStr);
  if (isNaN(date.getTime())) return "Lainnya";
  var monthIndex = date.getMonth();
  var namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return namaBulan[monthIndex] || "Lainnya";
}

// ==========================================
// 2. SISTEM CACHE MEMORI SERVER (SUPER NGEBUT)
// ==========================================
// Memecah data besar (chunking) agar lolos batas memori Google 100KB/item
function setCache(key, data) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    var chunkSize = 90000; 
    var chunks = Math.ceil(json.length / chunkSize);
    
    for (var i = 0; i < chunks; i++) {
      cache.put(key + '_' + i, json.substring(i * chunkSize, (i + 1) * chunkSize), 1800); // Tahan di RAM 30 menit
    }
    cache.put(key + '_chunks', chunks.toString(), 1800);
  } catch (e) {
    // Abaikan jika memori cache gagal, sistem akan otomatis membaca sheets biasa
  }
}

function getCache(key) {
  try {
    var cache = CacheService.getScriptCache();
    var chunksStr = cache.get(key + '_chunks');
    if (!chunksStr) return null;
    
    var chunks = parseInt(chunksStr, 10);
    var json = '';
    for (var i = 0; i < chunks; i++) {
      var chunk = cache.get(key + '_' + i);
      if (!chunk) return null; // Jika ada part hilang, paksa baca sheets ulang
      json += chunk;
    }
    return JSON.parse(json);
  } catch (e) { return null; }
}

function clearAllCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('allPenugasan_chunks');
    cache.remove('allLogs_chunks');
  } catch (e) {}
}


// ==========================================
// 3. FUNGSI PENGAMBILAN DATA CORE (DENGAN CACHE)
// ==========================================
function getRawLogs() {
  var cached = getCache('allLogs');
  if (cached) return cached; // Baca dari RAM Instan (0.1 detik)
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Log_Progress');
  var data = logSheet ? logSheet.getDataRange().getValues() : [];
  
  setCache('allLogs', data);
  return data;
}

function getAllPenugasanData() {
  var cached = getCache('allPenugasan');
  if (cached) return cached; // Baca dari RAM Instan (0.1 detik)
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var combinedData = [];
  var dummyHeader = [];
  
  var sheetSusenas = ss.getSheetByName('Penugasan_Susenas');
  if (sheetSusenas) {
    var dataSusenas = sheetSusenas.getDataRange().getValues();
    if (dataSusenas.length > 0) {
      dummyHeader = dataSusenas[0];
      if (dataSusenas.length > 1) combinedData = combinedData.concat(dataSusenas.slice(1));
    }
  }
  
  var sheetSeruti = ss.getSheetByName('Penugasan_Seruti');
  if (sheetSeruti) {
    var dataSeruti = sheetSeruti.getDataRange().getValues();
    if (dataSeruti.length > 0) {
      if (dummyHeader.length === 0) dummyHeader = dataSeruti[0];
      if (dataSeruti.length > 1) combinedData = combinedData.concat(dataSeruti.slice(1));
    }
  }
  
  var result = combinedData.length === 0 ? [] : [dummyHeader].concat(combinedData);
  setCache('allPenugasan', result);
  return result;
}

// ==========================================
// 4. DROPDOWN & FILTER LIST
// ==========================================
function getEmployeeList(prefetchedData) {
  var data = prefetchedData || getAllPenugasanData();
  if (data.length <= 1) return [];
  var names = new Set();
  for (var i = 1; i < data.length; i++) {
    if (data[i][9]) names.add(String(data[i][9]).trim());
    if (data[i][10]) names.add(String(data[i][10]).trim());
  }
  return Array.from(names).filter(Boolean).sort(function(a, b) {
    return a.localeCompare(b, 'id', { sensitivity: 'base' });
  });
}

function getWaktuPendataanList(prefetchedData) {
  var data = prefetchedData || getAllPenugasanData();
  if (data.length <= 1) return [];
  var waktuSet = new Set();
  for (var i = 1; i < data.length; i++) {
    if (data[i][12]) waktuSet.add(String(data[i][12]).trim());
  }
  var standardOrder = [
    "Januari", "Februari", "Maret", "Triwulan I", "April", "Mei", "Juni", "Triwulan II", 
    "Juli", "Agustus", "September", "Triwulan III", "Oktober", "November", "Desember", "Triwulan IV"
  ];
  return Array.from(waktuSet).filter(Boolean).sort(function(a, b) {
    var indexA = standardOrder.indexOf(a);
    var indexB = standardOrder.indexOf(b);
    return (indexA !== -1 ? indexA : 99) - (indexB !== -1 ? indexB : 99);
  });
}

function getSensusList(prefetchedData) {
  var data = prefetchedData || getAllPenugasanData();
  if (data.length <= 1) return [];
  var sensusSet = new Set();
  for (var i = 1; i < data.length; i++) {
    if (data[i][11]) {
      var parts = String(data[i][11]).split(',');
      parts.forEach(function(s) {
        var clean = s.trim();
        if (clean) sensusSet.add(clean);
      });
    }
  }
  return Array.from(sensusSet);
}


// ==========================================
// 5. PENGOLAHAN DATA DASHBOARD
// ==========================================
function getInitialDashboardData() {
  var data = getAllPenugasanData(); 
  return {
    names: getEmployeeList(data),
    waktus: getWaktuPendataanList(data),
    sensus: getSensusList(data),
    summary: getDashboardSummary('Semua', 'Semua', ['Semua'], data)
  };
}

function getDashboardSummary(filterNama, filterWaktu, filterSensusArray, prefetchedData) {
  var data = prefetchedData || getAllPenugasanData();
  if (data.length <= 1) return null;
  
  var totalTarget = 0, totalPPL = 0, totalPML = 0;
  var detailTugas = [];
  var waktuBarMap = {}; 
  var taskWaktuMap = {}; 
  
  var tepatWaktu = 0, terlambatPPL = 0, terlambatPML = 0, dalamProses = 0;
  var today = new Date();
  today.setHours(0,0,0,0);
  
  if (!filterSensusArray || !Array.isArray(filterSensusArray) || filterSensusArray.length === 0) {
    filterSensusArray = ['Semua'];
  }
  
  for (var i = 1; i < data.length; i++) {
    var nks = data[i][7]; 
    if (!nks || String(nks).trim() === "") continue;
    
    var ppl = data[i][9] ? String(data[i][9]).trim() : "";
    var pml = data[i][10] ? String(data[i][10]).trim() : "";
    var waktu = data[i][12] ? String(data[i][12]).trim() : "Tanpa Keterangan"; 
    var rawSensus = data[i][11] ? String(data[i][11]).trim() : "";
    
    var rowSensusList = rawSensus.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    rowSensusList.forEach(function(s) {
      taskWaktuMap[String(nks).trim() + "_" + s] = waktu;
    });
    if (!taskWaktuMap[String(nks).trim()]) taskWaktuMap[String(nks).trim()] = waktu;
    
    var matchNama = (!filterNama || filterNama === 'Semua' || ppl.toLowerCase() === filterNama.toLowerCase() || pml.toLowerCase() === filterNama.toLowerCase());
    var matchWaktu = (!filterWaktu || filterWaktu === 'Semua' || waktu === filterWaktu);
    
    var matchSensus = false;
    if (filterSensusArray.indexOf('Semua') !== -1) {
      matchSensus = true;
    } else {
      matchSensus = rowSensusList.some(function(s) { return filterSensusArray.indexOf(s) !== -1; });
    }
    
    // PERBAIKAN: Seluruh perhitungan grafik Bar Chart, Donut, KPI, dan Tabel 
    // kini HANYA akan dieksekusi JIKA data tersebut lolos KETIGA filter sekaligus.
    if (matchNama && matchWaktu && matchSensus) {
      var target = Number(data[i][18]) || 10;
      var rPPL = Number(data[i][19]) || 0;
      var rPML = Number(data[i][20]) || 0;
      var statusSheet = data[i][21] ? String(data[i][21]).trim() : "";
      
      var deadPPL = parseDateValue(data[i][14]);
      var deadPML = parseDateValue(data[i][15]);
      var donePPL = parseDateValue(data[i][16]);
      var donePML = parseDateValue(data[i][17]);
      
      var deadPPLTime = deadPPL ? deadPPL.setHours(0,0,0,0) : null;
      var deadPMLTime = deadPML ? deadPML.setHours(0,0,0,0) : null;
      var donePPLTime = donePPL ? donePPL.setHours(0,0,0,0) : null;
      var donePMLTime = donePML ? donePML.setHours(0,0,0,0) : null;
      
      var isPPLTerlambat = false;
      if (donePPLTime && deadPPLTime && donePPLTime > deadPPLTime) isPPLTerlambat = true;
      else if (!donePPLTime && deadPPLTime && today.getTime() > deadPPLTime && rPPL < target) isPPLTerlambat = true;
      
      var isPMLTerlambat = false;
      if (donePMLTime && deadPMLTime && donePMLTime > deadPMLTime) isPMLTerlambat = true;
      else if (!donePMLTime && deadPMLTime && today.getTime() > deadPMLTime && rPML < target) isPMLTerlambat = true;
      
      // Mengisi Bar Chart (Evaluasi Waktu) di dalam kondisi filter
      if (waktu) {
        if (!waktuBarMap[waktu]) {
          waktuBarMap[waktu] = { pplTepat: 0, pplTerlambat: 0, pmlTepat: 0, pmlTerlambat: 0 };
        }
        if (isPPLTerlambat) waktuBarMap[waktu].pplTerlambat++;
        else if (rPPL >= target) waktuBarMap[waktu].pplTepat++;
        
        if (isPMLTerlambat) waktuBarMap[waktu].pmlTerlambat++;
        else if (rPML >= target) waktuBarMap[waktu].pmlTepat++;
      }
      
      // Akumulasi KPI
      totalTarget += target;
      totalPPL += rPPL;
      totalPML += rPML;
      
      // Akumulasi Donut Chart
      if (filterNama && filterNama !== 'Semua') {
        if (ppl.toLowerCase() === filterNama.toLowerCase()) {
          if (isPPLTerlambat) terlambatPPL++; else if (rPPL >= target) tepatWaktu++; else dalamProses++;
        } else if (pml.toLowerCase() === filterNama.toLowerCase()) {
          if (isPMLTerlambat) terlambatPML++; else if (rPML >= target) tepatWaktu++; else dalamProses++;
        }
      } else {
        if (isPPLTerlambat) terlambatPPL++;
        else if (isPMLTerlambat) terlambatPML++;
        else if (rPPL >= target && rPML >= target) tepatWaktu++;
        else dalamProses++;
      }
      
      // Pengisian Tabel Detail
      detailTugas.push({
        nks: String(nks), waktuPendataan: waktu, sensus: rawSensus,
        lokasi: "Kec. " + data[i][2] + ", Kel. " + data[i][4] + " (" + data[i][8] + ")",
        ppl: ppl, pml: pml, target: target, realisasiPPL: rPPL, realisasiPML: rPML,
        status: statusSheet, isPPLTerlambat: isPPLTerlambat, isPMLTerlambat: isPMLTerlambat
      });
    }
  }

  var allWaktus = Object.keys(waktuBarMap);
  var monthOrder = { 
    "Januari": 1, "Februari": 2, "Maret": 3, "Triwulan I": 3.5, 
    "April": 4, "Mei": 5, "Juni": 6, "Triwulan II": 6.5, 
    "Juli": 7, "Agustus": 8, "September": 9, "Triwulan III": 9.5, 
    "Oktober": 10, "November": 11, "Desember": 12, "Triwulan IV": 12.5,
    "Tanpa Keterangan": 99 
  };
  allWaktus.sort(function(a, b) { return (monthOrder[a] || 99) - (monthOrder[b] || 99); });

  var dailyMap = {};
  var logData = getRawLogs(); 
  
  if (logData && logData.length > 0) {
    for (var k = 1; k < logData.length; k++) {
      var ts = logData[k][0];
      var nksLog = logData[k][1] ? String(logData[k][1]).trim() : ""; 
      var petugasLog = logData[k][2] ? String(logData[k][2]).trim() : "";
      var surveiLog = logData[k][4] ? String(logData[k][4]).trim() : "";
      var realisasiLog = Number(logData[k][5]) || 0;
      
      if (!ts) continue;
      var dObj = (ts instanceof Date) ? ts : new Date(ts);
      if (isNaN(dObj.getTime())) continue;
      
      var matchPetugasLog = (!filterNama || filterNama === 'Semua' || petugasLog.toLowerCase() === filterNama.toLowerCase());
      var taskWaktuLog = taskWaktuMap[nksLog + "_" + surveiLog] || taskWaktuMap[nksLog] || getBulanPendataan(ts);
      var matchWaktuLog = (!filterWaktu || filterWaktu === 'Semua' || taskWaktuLog === filterWaktu);
      var matchSensusLog = (filterSensusArray.indexOf('Semua') !== -1 || filterSensusArray.indexOf(surveiLog) !== -1);
      
      if (matchPetugasLog && matchWaktuLog && matchSensusLog) {
        var dateKey = formatDateISO(dObj);
        if (!dailyMap[dateKey]) dailyMap[dateKey] = 0;
        dailyMap[dateKey] += realisasiLog;
      }
    }
  }

  var sortedDates = Object.keys(dailyMap).sort();
  var dailyLabels = sortedDates.map(function(dStr) {
    var parts = dStr.split('-');
    return parts[2] + '/' + parts[1];
  });
  var dailyValues = sortedDates.map(function(dStr) { return dailyMap[dStr]; });

  return {
    totalTarget: totalTarget, totalPPL: totalPPL, totalPML: totalPML,
    percentPPL: totalTarget > 0 ? Math.round((totalPPL / totalTarget) * 100) : 0,
    percentPML: totalPPL > 0 ? Math.round((totalPML / totalPPL) * 100) : 0,
    detailTugas: detailTugas,
    donutData: {
      tepatWaktu: tepatWaktu,
      terlambat: terlambatPPL + terlambatPML,
      dalamProses: dalamProses,
      persenTepat: detailTugas.length > 0 ? Math.round((tepatWaktu / detailTugas.length) * 100) : 0
    },
    waktuLabels: allWaktus, waktuBarData: waktuBarMap,
    dailyTrendLabels: dailyLabels, dailyTrendData: dailyValues
  };
}


// ==========================================
// 6. PENGOLAHAN DATA FORM PROGRESS PETUGAS
// ==========================================
function getEmployeeData(namaPetugas) {
  return {
    tasks: getEmployeeTasks(namaPetugas),
    logs: getEmployeeLogs(namaPetugas)
  };
}

function getEmployeeTasks(namaPetugas) {
  var data = getAllPenugasanData();
  if (data.length <= 1 || !namaPetugas) return [];
  var tasks = [];
  var namaClean = String(namaPetugas).trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    var nks = data[i][7];
    if (!nks || String(nks).trim() === "") continue;
    var ppl = data[i][9] ? String(data[i][9]).trim().toLowerCase() : "";
    var pml = data[i][10] ? String(data[i][10]).trim().toLowerCase() : "";
    var isPPL = (ppl === namaClean);
    var isPML = (pml === namaClean);
    
    if (isPPL || isPML) {
      var targetVal = Number(data[i][18]) || 10;
      var rPPL = Number(data[i][19]) || 0;
      var rPML = Number(data[i][20]) || 0;
      var tglSelesaiPPL = data[i][16];
      var tglSelesaiPML = data[i][17];
      if (isPPL && rPPL >= targetVal && tglSelesaiPPL) continue;
      if (isPML && rPML >= targetVal && tglSelesaiPML) continue;
      
      tasks.push({
        rowIndex: i + 1, no: data[i][0], nks: String(nks),
        sensus: data[i][11] ? String(data[i][11]).trim() : "Sensus Umum",
        waktuPendataan: data[i][12] ? String(data[i][12]).trim() : "",
        lokasi: "Kec. " + data[i][2] + ", Kel. " + data[i][4] + " (" + data[i][8] + ")",
        role: isPPL ? "PPL" : "PML", target: targetVal, realisasiPPL: rPPL, realisasiPML: rPML
      });
    }
  }
  return tasks;
}

function getEmployeeLogs(namaPetugas) {
  var data = getRawLogs(); // Baca dari Cache Instan
  var logs = [];
  var namaClean = String(namaPetugas).trim().toLowerCase();
  
  for (var i = data.length - 1; i >= 1; i--) {
    var petugasLog = data[i][2] ? String(data[i][2]).trim().toLowerCase() : "";
    if (petugasLog === namaClean) {
      
      // Formatting tanggal yang aman walau ditarik dari RAM/Cache (String Format)
      var rawTimestamp = data[i][0];
      var safeTimestamp = "";
      if (rawTimestamp instanceof Date) {
        safeTimestamp = Utilities.formatDate(rawTimestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      } else {
        var d = new Date(rawTimestamp);
        if (!isNaN(d.getTime())) safeTimestamp = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        else safeTimestamp = String(rawTimestamp);
      }
      
      logs.push({ 
        timestamp: safeTimestamp, nks: String(data[i][1]), 
        peran: String(data[i][3]), survei: String(data[i][4]), realisasi: Number(data[i][5]) || 0 
      });
      if (logs.length >= 10) break; 
    }
  }
  return logs;
}


// ==========================================
// 7. PENYIMPANAN DATA (WRITE) & RESET CACHE
// ==========================================
function recordProgressLog(nks, namaPetugas, role, realisasiVal, namaSensus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Log_Progress');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log_Progress');
    logSheet.appendRow(['Timestamp', 'NKS', 'Nama Petugas', 'Peran', 'Nama Survei', 'Realisasi KK']);
  }
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  logSheet.appendRow([timestamp, nks, namaPetugas, role, namaSensus, realisasiVal]);
}

function findSheetAndRow(rowIndexFromCombined) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSusenas = ss.getSheetByName('Penugasan_Susenas');
  var sheetSeruti = ss.getSheetByName('Penugasan_Seruti');
  var countSusenas = sheetSusenas ? sheetSusenas.getLastRow() - 1 : 0;
  if (rowIndexFromCombined <= countSusenas + 1) {
    return { sheet: sheetSusenas, row: rowIndexFromCombined };
  } else {
    return { sheet: sheetSeruti, row: rowIndexFromCombined - countSusenas };
  }
}

function updateTaskProgress(rowIndex, valPPL, valPML) {
  var targetInfo = findSheetAndRow(rowIndex);
  var sheet = targetInfo.sheet;
  var actualRow = targetInfo.row;
  if (!sheet) throw new Error("Sheet tab penugasan tidak ditemukan!");
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var targetVal = Number(sheet.getRange(actualRow, 19).getValue()) || 10;
  var nks = sheet.getRange(actualRow, 8).getValue();
  var pplName = sheet.getRange(actualRow, 10).getValue();
  var pmlName = sheet.getRange(actualRow, 11).getValue();
  var namaSensus = sheet.getRange(actualRow, 12).getValue();
  
  if (valPPL !== null && valPPL !== "") {
    var inputHarianPPL = Number(valPPL);
    var currentPPL = Number(sheet.getRange(actualRow, 20).getValue()) || 0;
    var newTotalPPL = currentPPL + inputHarianPPL;
    sheet.getRange(actualRow, 20).setValue(newTotalPPL);
    if (newTotalPPL >= targetVal && !sheet.getRange(actualRow, 17).getValue()) sheet.getRange(actualRow, 17).setValue(todayStr);
    recordProgressLog(nks, pplName, "PPL", inputHarianPPL, namaSensus);
  }
  
  if (valPML !== null && valPML !== "") {
    var inputHarianPML = Number(valPML);
    var currentPML = Number(sheet.getRange(actualRow, 21).getValue()) || 0;
    var newTotalPML = currentPML + inputHarianPML;
    sheet.getRange(actualRow, 21).setValue(newTotalPML);
    if (newTotalPML >= targetVal && !sheet.getRange(actualRow, 18).getValue()) sheet.getRange(actualRow, 18).setValue(todayStr);
    recordProgressLog(nks, pmlName, "PML", inputHarianPML, namaSensus);
  }
  
  // SANGAT PENTING: Bersihkan Cache setelah data diupdate agar user lain dapat data terbaru!
  clearAllCache(); 
  
  return "Progres berhasil diperbarui, diakumulasikan, dan dicatat ke riwayat!";
}


// ==========================================
// 8. MASTER DATA LAINNYA
// ==========================================
function getMasterKecamatan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master_Wilayah');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var namaKec = data[i][1] ? String(data[i][1]).trim() : ""; 
    if (namaKec !== "") list.push(namaKec);
  }
  return Array.from(new Set(list)).sort();
}

function getKodeKecamatan(namaKec) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master_Wilayah');
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var kode = String(data[i][0]).trim(); 
    var nama = String(data[i][1]).trim(); 
    if (nama.toLowerCase() === namaKec.toLowerCase()) return kode; 
  }
  return "";
}

function addNewTask(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sensusName = String(formData.namaSensus).toLowerCase();
  var sheetName = (sensusName.indexOf('seruti') !== -1) ? 'Penugasan_Seruti' : 'Penugasan_Susenas';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' tidak ditemukan!");
  
  var values = sheet.getRange("H1:H" + sheet.getMaxRows()).getValues();
  var lastRowWithData = 0;
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null) { lastRowWithData = i + 1; break; }
  }
  var targetRow = lastRowWithData + 1;
  var newNo = targetRow - 1;
  var namaKec = String(formData.namaKecamatan).trim();
  var targetVolume = Number(formData.targetVolume) || 10;
  var waktuPendataan = ""; 
  
  var newRow = [
    newNo, '', namaKec, formData.kodeDesa, formData.namaDesa, 
    formData.kodeSLS, formData.kodeSubSLS, formData.nks, formData.namaSLS, 
    formData.namaPPL, formData.namaPML, formData.namaSensus, waktuPendataan, 
    formData.tglMulai, formData.tglDeadlinePPL, formData.tglDeadlinePML, '', '', 
    targetVolume, 0, 0, ''
  ];
  sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
  
  // SANGAT PENTING: Bersihkan Cache setelah Admin nambah data baru
  clearAllCache();
  
  return "Penugasan NKS " + formData.nks + " berhasil ditambahkan ke tab " + sheetName + "!";
}

function forceRefreshDashboardData() {
  clearAllCache(); // Hapus memori server yang lama
  return getInitialDashboardData(); // Tarik ulang data murni dari Google Sheets
}