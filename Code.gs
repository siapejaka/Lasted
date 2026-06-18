const SPREADSHEET_ID = '17MDxSku4JHzv7mv-t3EObHLucYFuXMpwlcl0wFKJ6fY';
const TAB_CRM = 'CRM';
const TAB_PRODUK = 'Produk';
const TAB_PENGGUNA = 'Pengguna';
const TAB_WILAYAH = 'Wilayah';
const TAB_ONGKOS_KIRIM = 'OngkosKirim';
const TAB_TRANSAKSI = 'Transaksi';
const TAB_SETTINGS = 'Settings';

// ============================================
// DOGET - Serve the web app
// ============================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CRM Consumer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function serializeForClient(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(function(item) {
      return serializeForClient(item);
    });
  }
  if (typeof value === 'object') {
    var result = {};
    for (var key in value) {
      if (value.hasOwnProperty(key)) {
        result[key] = serializeForClient(value[key]);
      }
    }
    return result;
  }
  return value;
}

function sanitizeInput(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/<[^>]*>/g, '').trim();
}

function formatDateForClient(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  if (date instanceof Date) {
    return date.toISOString().split('T')[0];
  }
  return String(date);
}

function formatDateTimeForClient(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  if (date instanceof Date) {
    return date.toISOString().slice(0, 19).replace('T', ' ');
  }
  return String(date);
}

function hashPassword(password) {
  if (!password) return '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  var hash = '';
  for (var i = 0; i < digest.length; i++) {
    var byte = digest[i];
    if (byte < 0) byte += 256;
    var hex = ('0' + byte.toString(16)).slice(-2);
    hash += hex;
  }
  return hash;
}

function generateId(prefix, sheet) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var targetSheet = ss.getSheetByName(sheet);
  var lastRow = targetSheet ? targetSheet.getLastRow() : 1;
  var sequence = lastRow;
  return prefix + ('000' + sequence).slice(-3);
}

function getNextId(prefix, sheet, columnIndex) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var targetSheet = ss.getSheetByName(sheet);
  var lastRow = targetSheet.getLastRow();
  var maxNum = 0;
  
  for (var i = 2; i <= lastRow; i++) {
    var id = targetSheet.getRange(i, columnIndex).getValue();
    if (id && id.toString().indexOf(prefix) === 0) {
      var num = parseInt(id.replace(prefix, ''));
      if (num > maxNum) maxNum = num;
    }
  }
  
  return prefix + ('000' + (maxNum + 1)).slice(-3);
}

// ============================================
// SETUP DATABASE
// ============================================
function setupDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var existingSheets = ss.getSheets().map(function(s) { return s.getName(); });
  
  // CRM Sheet
  if (existingSheets.indexOf(TAB_CRM) === -1) {
    ss.insertSheet(TAB_CRM);
  }
  var crmSheet = ss.getSheetByName(TAB_CRM);
  crmSheet.clear();
  crmSheet.getRange(1, 1, 1, 15).setValues([[
    'id_crm', 'nama_lengkap', 'nomor_whatsapp', 'email', 'tag', 'produk_diminati', 
    'sumber_lead', 'notes', 'assigned_to', 'tanggal_follow_up', 'status', 
    'kategori', 'produk_dibeli', 'tanggal_pembelian', 'created_date'
  ]]);
  crmSheet.getRange(1, 1, 1, 15).setFontWeight('bold');
  crmSheet.setFrozenRows(1);
  
  // Produk Sheet
  if (existingSheets.indexOf(TAB_PRODUK) === -1) {
    ss.insertSheet(TAB_PRODUK);
  }
  var produkSheet = ss.getSheetByName(TAB_PRODUK);
  produkSheet.clear();
  produkSheet.getRange(1, 1, 1, 18).setValues([[
    'id_produk', 'nama_produk', 'deskripsi', 'kategori', 
    'tier_1', 'tier_2', 'tier_3', 'tier_4', 'tier_5', 'tier_6', 'tier_7', 'tier_8', 'tier_9',
    'tier_harga_modal', 'unit', 'masa_konsumsi', 'status', 'created_date'
  ]]);
  produkSheet.getRange(1, 1, 1, 18).setFontWeight('bold');
  produkSheet.setFrozenRows(1);
  
  // Pengguna Sheet
  if (existingSheets.indexOf(TAB_PENGGUNA) === -1) {
    ss.insertSheet(TAB_PENGGUNA);
  }
  var penggunaSheet = ss.getSheetByName(TAB_PENGGUNA);
  penggunaSheet.clear();
  penggunaSheet.getRange(1, 1, 1, 9).setValues([[
    'id_pengguna', 'nama_lengkap', 'username', 'password_hash', 'role', 'status', 'created_date', '', ''
  ]]);
  penggunaSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  penggunaSheet.setFrozenRows(1);
  
  // Add default admin user
  var adminPasswordHash = hashPassword('admin123');
  penggunaSheet.appendRow(['USR-001', 'Administrator', 'admin', adminPasswordHash, 'Admin', 'Aktif', new Date()]);
  
  // Wilayah Sheet
  if (existingSheets.indexOf(TAB_WILAYAH) === -1) {
    ss.insertSheet(TAB_WILAYAH);
  }
  var wilayahSheet = ss.getSheetByName(TAB_WILAYAH);
  wilayahSheet.clear();
  wilayahSheet.getRange(1, 1, 1, 9).setValues([[
    'id_wilayah', 'nama_wilayah', 'provinsi', 'kabupaten_kota', 'kecamatan', 
    'kelurahan', 'kode_pos', 'status', 'created_date'
  ]]);
  wilayahSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  wilayahSheet.setFrozenRows(1);
  
  // Ongkos Kirim Sheet
  if (existingSheets.indexOf(TAB_ONGKOS_KIRIM) === -1) {
    ss.insertSheet(TAB_ONGKOS_KIRIM);
  }
  var ongkirSheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
  ongkirSheet.clear();
  ongkirSheet.getRange(1, 1, 1, 8).setValues([[
    'id_ongkos', 'id_wilayah', 'nama_ekspedisi', 'harga', 'estimasi_waktu', 'status', 'created_date', ''
  ]]);
  ongkirSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  ongkirSheet.setFrozenRows(1);
  
  // Settings Sheet
  if (existingSheets.indexOf(TAB_SETTINGS) === -1) {
    ss.insertSheet(TAB_SETTINGS);
  }
  var settingsSheet = ss.getSheetByName(TAB_SETTINGS);
  settingsSheet.clear();
  settingsSheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'updated_at']]);
  settingsSheet.setFrozenRows(1);
  
  // Transaksi Sheet
  if (existingSheets.indexOf(TAB_TRANSAKSI) === -1) {
    ss.insertSheet(TAB_TRANSAKSI);
  }
  var transaksiSheet = ss.getSheetByName(TAB_TRANSAKSI);
  transaksiSheet.clear();
  transaksiSheet.getRange(1, 1, 1, 16).setValues([[
    'id_transaksi', 'id_crm', 'nama_customer', 'nomor_whatsapp', 'alamat_pengiriman',
    'tanggal_transaksi', 'produk', 'jumlah', 'harga_tier', 'harga_satuan',
    'subtotal', 'ongkir', 'total', 'ekspedisi', 'status_pembayaran', 'created_date'
  ]]);
  transaksiSheet.getRange(1, 1, 1, 16).setFontWeight('bold');
  transaksiSheet.setFrozenRows(1);
  
  return { ok: true, message: 'Database berhasil disetup' };
}

// ============================================
// CRM FUNCTIONS
// ============================================
function getCRMData(type) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_CRM);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: [], summary: getCRMSummary([]) };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = {
        id_crm: data[i][0],
        nama_lengkap: data[i][1],
        nomor_whatsapp: data[i][2],
        email: data[i][3] || '',
        tag: data[i][4] || 'Cold',
        produk_diminati: data[i][5] || '',
        sumber_lead: data[i][6] || '',
        notes: data[i][7] || '',
        assigned_to: data[i][8] || '',
        tanggal_follow_up: formatDateForClient(data[i][9]),
        status: data[i][10] || 'Baru',
        kategori: data[i][11] || 'Lead',
        produk_dibeli: data[i][12] || '',
        tanggal_pembelian: formatDateForClient(data[i][13]),
        created_date: formatDateTimeForClient(data[i][14])
      };
      
      if (type === 'customer' && row.kategori === 'Customer') {
        result.push(row);
      } else if (type === 'lead' && row.kategori === 'Lead') {
        result.push(row);
      }
    }
    
    return { ok: true, data: result, summary: getCRMSummary(result) };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data CRM: ' + error.message };
  }
}

function getCRMSummary(data) {
  var summary = {
    total: data.length,
    hot: 0,
    warm: 0,
    cold: 0,
    baru: 0,
    dalamProspek: 0,
    deal: 0,
    tidakAktif: 0
  };
  
  for (var i = 0; i < data.length; i++) {
    var tag = data[i].tag;
    var status = data[i].status;
    
    if (tag === 'Hot') summary.hot++;
    else if (tag === 'Warm') summary.warm++;
    else if (tag === 'Cold') summary.cold++;
    
    if (status === 'Baru') summary.baru++;
    else if (status === 'Dalam Prospek') summary.dalamProspek++;
    else if (status === 'Deal') summary.deal++;
    else if (status === 'Tidak Aktif') summary.tidakAktif++;
  }
  
  return summary;
}

function getAllCRMData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_CRM);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: [] };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      result.push({
        id_crm: data[i][0],
        nama_lengkap: data[i][1],
        nomor_whatsapp: data[i][2],
        email: data[i][3] || '',
        tag: data[i][4] || 'Cold',
        produk_diminati: data[i][5] || '',
        sumber_lead: data[i][6] || '',
        notes: data[i][7] || '',
        assigned_to: data[i][8] || '',
        tanggal_follow_up: formatDateForClient(data[i][9]),
        status: data[i][10] || 'Baru',
        kategori: data[i][11] || 'Lead',
        produk_dibeli: data[i][12] || '',
        tanggal_pembelian: formatDateForClient(data[i][13]),
        created_date: formatDateTimeForClient(data[i][14])
      });
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data CRM: ' + error.message };
  }
}

function saveCRMData(crmData) {
  try {
    // Validation
    if (!crmData.nama_lengkap) {
      return { ok: false, message: 'Nama Lengkap wajib diisi' };
    }
    if (!crmData.nomor_whatsapp) {
      return { ok: false, message: 'Nomor WhatsApp wajib diisi' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_CRM);
    
    var namaLengkap = sanitizeInput(crmData.nama_lengkap);
    var nomorWhatsApp = sanitizeInput(crmData.nomor_whatsapp).replace(/[\s\-\.\,]/g, '');
    var email = sanitizeInput(crmData.email || '');
    var tag = sanitizeInput(crmData.tag || 'Cold');
    var produkDiminati = sanitizeInput(crmData.produk_diminati || '');
    var sumberLead = sanitizeInput(crmData.sumber_lead || '');
    var notes = sanitizeInput(crmData.notes || '');
    var assignedTo = sanitizeInput(crmData.assigned_to || '');
    var tanggalFollowUp = crmData.tanggal_follow_up || null;
    var status = sanitizeInput(crmData.status || 'Baru');
    
    if (crmData.id_crm) {
      // Update existing
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() == crmData.id_crm) {
          sheet.getRange(i, 2).setValue(namaLengkap);
          sheet.getRange(i, 3).setValue(nomorWhatsApp);
          sheet.getRange(i, 4).setValue(email);
          sheet.getRange(i, 5).setValue(tag);
          sheet.getRange(i, 6).setValue(produkDiminati);
          sheet.getRange(i, 7).setValue(sumberLead);
          sheet.getRange(i, 8).setValue(notes);
          sheet.getRange(i, 9).setValue(assignedTo);
          sheet.getRange(i, 10).setValue(tanggalFollowUp ? new Date(tanggalFollowUp) : null);
          sheet.getRange(i, 11).setValue(status);
          return { ok: true, message: 'Data CRM berhasil diperbarui' };
        }
      }
      return { ok: false, message: 'Data CRM tidak ditemukan' };
    } else {
      // Check duplicate WhatsApp
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        var existingWa = sheet.getRange(i, 3).getValue().toString().replace(/[\s\-\.\,]/g, '');
        if (existingWa === nomorWhatsApp) {
          return { ok: false, message: 'Nomor WhatsApp sudah terdaftar' };
        }
      }
      
      // Insert new
      var newId = getNextId('CRM-', TAB_CRM, 1);
      var newRow = [
        newId,
        namaLengkap,
        nomorWhatsApp,
        email,
        tag,
        produkDiminati,
        sumberLead,
        notes,
        assignedTo,
        tanggalFollowUp ? new Date(tanggalFollowUp) : null,
        status,
        'Lead',
        '',
        null,
        new Date()
      ];
      sheet.appendRow(newRow);
      return { ok: true, message: 'Data CRM berhasil disimpan', id: newId };
    }
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan data CRM: ' + error.message };
  }
}

function deleteCRMData(idCRM) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_CRM);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idCRM) {
        sheet.deleteRow(i);
        return { ok: true, message: 'Data CRM berhasil dihapus' };
      }
    }
    
    return { ok: false, message: 'Data CRM tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menghapus data CRM: ' + error.message };
  }
}

function convertLeadToCustomer(idCRM, purchaseData) {
  try {
    if (!purchaseData.produk_dibeli) {
      return { ok: false, message: 'Produk yang dibeli wajib diisi' };
    }
    if (!purchaseData.tanggal_pembelian) {
      return { ok: false, message: 'Tanggal pembelian wajib diisi' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_CRM);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idCRM) {
        sheet.getRange(i, 11).setValue('Deal');
        sheet.getRange(i, 12).setValue('Customer');
        sheet.getRange(i, 13).setValue(sanitizeInput(purchaseData.produk_dibeli));
        sheet.getRange(i, 14).setValue(new Date(purchaseData.tanggal_pembelian));
        sheet.getRange(i, 10).setValue(new Date(purchaseData.tanggal_pembelian));
        return { ok: true, message: 'Lead berhasil dikonversi ke Customer' };
      }
    }
    
    return { ok: false, message: 'Data CRM tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal mengkonversi lead: ' + error.message };
  }
}

// ============================================
// PRODUK FUNCTIONS
// ============================================
function getProdukData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PRODUK);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: [] };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      result.push({
        id_produk: data[i][0],
        nama_produk: data[i][1],
        deskripsi: data[i][2] || '',
        kategori: data[i][3] || '',
        tier_1: data[i][4] || 0,
        tier_2: data[i][5] || 0,
        tier_3: data[i][6] || 0,
        tier_4: data[i][7] || 0,
        tier_5: data[i][8] || 0,
        tier_6: data[i][9] || 0,
        tier_7: data[i][10] || 0,
        tier_8: data[i][11] || 0,
        tier_9: data[i][12] || 0,
        tier_harga_modal: data[i][13] || 1,
        unit: data[i][14] || '',
        masa_konsumsi: data[i][15] || 0,
        status: data[i][16] || 'Aktif',
        created_date: formatDateTimeForClient(data[i][17])
      });
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data produk: ' + error.message };
  }
}

function saveProduk(produkData) {
  try {
    if (!produkData.nama_produk) {
      return { ok: false, message: 'Nama Produk wajib diisi' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PRODUK);
    
    var namaProduk = sanitizeInput(produkData.nama_produk);
    var deskripsi = sanitizeInput(produkData.deskripsi || '');
    var kategori = sanitizeInput(produkData.kategori || '');
    var tier1 = parseFloat(produkData.tier_1) || 0;
    var tier2 = parseFloat(produkData.tier_2) || 0;
    var tier3 = parseFloat(produkData.tier_3) || 0;
    var tier4 = parseFloat(produkData.tier_4) || 0;
    var tier5 = parseFloat(produkData.tier_5) || 0;
    var tier6 = parseFloat(produkData.tier_6) || 0;
    var tier7 = parseFloat(produkData.tier_7) || 0;
    var tier8 = parseFloat(produkData.tier_8) || 0;
    var tier9 = parseFloat(produkData.tier_9) || 0;
    var tierHargaModal = parseInt(produkData.tier_harga_modal) || 1;
    var unit = sanitizeInput(produkData.unit || '');
    var masaKonsumsi = parseInt(produkData.masa_konsumsi) || 0;
    var status = sanitizeInput(produkData.status || 'Aktif');
    
    if (produkData.id_produk) {
      // Update existing
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() == produkData.id_produk) {
          sheet.getRange(i, 2).setValue(namaProduk);
          sheet.getRange(i, 3).setValue(deskripsi);
          sheet.getRange(i, 4).setValue(kategori);
          sheet.getRange(i, 5).setValue(tier1);
          sheet.getRange(i, 6).setValue(tier2);
          sheet.getRange(i, 7).setValue(tier3);
          sheet.getRange(i, 8).setValue(tier4);
          sheet.getRange(i, 9).setValue(tier5);
          sheet.getRange(i, 10).setValue(tier6);
          sheet.getRange(i, 11).setValue(tier7);
          sheet.getRange(i, 12).setValue(tier8);
          sheet.getRange(i, 13).setValue(tier9);
          sheet.getRange(i, 14).setValue(tierHargaModal);
          sheet.getRange(i, 15).setValue(unit);
          sheet.getRange(i, 16).setValue(masaKonsumsi);
          sheet.getRange(i, 17).setValue(status);
          return { ok: true, message: 'Produk berhasil diperbarui' };
        }
      }
      return { ok: false, message: 'Produk tidak ditemukan' };
    } else {
      // Check duplicate name
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 2).getValue().toString().toLowerCase() === namaProduk.toLowerCase()) {
          return { ok: false, message: 'Nama produk sudah terdaftar' };
        }
      }
      
      // Insert new
      var newId = getNextId('PRD-', TAB_PRODUK, 1);
      var newRow = [
        newId,
        namaProduk,
        deskripsi,
        kategori,
        tier1, tier2, tier3, tier4, tier5, tier6, tier7, tier8, tier9,
        tierHargaModal,
        unit,
        masaKonsumsi,
        status,
        new Date()
      ];
      sheet.appendRow(newRow);
      return { ok: true, message: 'Produk berhasil disimpan', id: newId };
    }
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan produk: ' + error.message };
  }
}

function deleteProduk(idProduk) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PRODUK);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idProduk) {
        sheet.deleteRow(i);
        return { ok: true, message: 'Produk berhasil dihapus' };
      }
    }
    
    return { ok: false, message: 'Produk tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menghapus produk: ' + error.message };
  }
}

function getProdukOptions() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PRODUK);
    var result = { kategori: [], produk: [] };
    
    if (sheet && sheet.getLastRow() > 1) {
      var lastRow = sheet.getLastRow();
      var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
      var kategoriSet = {};
      
      for (var i = 0; i < data.length; i++) {
        if (data[i][16] === 'Aktif') {
          result.produk.push({
            id: data[i][0],
            nama: data[i][1],
            kategori: data[i][3]
          });
          if (data[i][3]) {
            kategoriSet[data[i][3]] = true;
          }
        }
      }
      
      result.kategori = Object.keys(kategoriSet);
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil opsi produk: ' + error.message };
  }
}

// ============================================
// PENGGUNA FUNCTIONS
// ============================================
function getPenggunaData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PENGGUNA);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: [] };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      result.push({
        id_pengguna: data[i][0],
        nama_lengkap: data[i][1],
        username: data[i][2],
        password_hash: data[i][3],
        role: data[i][4] || 'Sales',
        status: data[i][5] || 'Aktif',
        created_date: formatDateTimeForClient(data[i][6])
      });
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data pengguna: ' + error.message };
  }
}

function savePengguna(penggunaData) {
  try {
    if (!penggunaData.nama_lengkap) {
      return { ok: false, message: 'Nama Lengkap wajib diisi' };
    }
    if (!penggunaData.username) {
      return { ok: false, message: 'Username wajib diisi' };
    }
    if (!penggunaData.id_pengguna && !penggunaData.password) {
      return { ok: false, message: 'Password wajib diisi untuk pengguna baru' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PENGGUNA);
    
    var namaLengkap = sanitizeInput(penggunaData.nama_lengkap);
    var username = sanitizeInput(penggunaData.username).toLowerCase();
    var role = sanitizeInput(penggunaData.role || 'Sales');
    var status = sanitizeInput(penggunaData.status || 'Aktif');
    var passwordHash = penggunaData.id_pengguna ? null : hashPassword(penggunaData.password);
    
    if (penggunaData.id_pengguna) {
      // Update existing
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() == penggunaData.id_pengguna) {
          sheet.getRange(i, 2).setValue(namaLengkap);
          sheet.getRange(i, 3).setValue(username);
          sheet.getRange(i, 5).setValue(role);
          sheet.getRange(i, 6).setValue(status);
          if (penggunaData.password) {
            sheet.getRange(i, 4).setValue(hashPassword(penggunaData.password));
          }
          return { ok: true, message: 'Pengguna berhasil diperbarui' };
        }
      }
      return { ok: false, message: 'Pengguna tidak ditemukan' };
    } else {
      // Check duplicate username
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 3).getValue().toString().toLowerCase() === username) {
          return { ok: false, message: 'Username sudah digunakan' };
        }
      }
      
      // Insert new
      var newId = getNextId('USR-', TAB_PENGGUNA, 1);
      var newRow = [
        newId,
        namaLengkap,
        username,
        passwordHash,
        role,
        status,
        new Date()
      ];
      sheet.appendRow(newRow);
      return { ok: true, message: 'Pengguna berhasil disimpan', id: newId };
    }
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan pengguna: ' + error.message };
  }
}

function deletePengguna(idPengguna) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PENGGUNA);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idPengguna) {
        sheet.deleteRow(i);
        return { ok: true, message: 'Pengguna berhasil dihapus' };
      }
    }
    
    return { ok: false, message: 'Pengguna tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menghapus pengguna: ' + error.message };
  }
}

function getPenggunaOptions(role) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_PENGGUNA);
    var result = [];
    
    if (sheet && sheet.getLastRow() > 1) {
      var lastRow = sheet.getLastRow();
      var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      
      for (var i = 0; i < data.length; i++) {
        var userRole = data[i][4];
        var userStatus = data[i][5];
        
        if (userStatus === 'Aktif') {
          if (!role || (role === 'Sales' && userRole === 'Sales')) {
            result.push({
              id: data[i][0],
              nama: data[i][1]
            });
          }
        }
      }
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil opsi pengguna: ' + error.message };
  }
}

// ============================================
// WILAYAH FUNCTIONS
// ============================================
function getWilayahData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_WILAYAH);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: [] };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      result.push({
        id_wilayah: data[i][0],
        nama_wilayah: data[i][1],
        provinsi: data[i][2] || '',
        kabupaten_kota: data[i][3] || '',
        kecamatan: data[i][4] || '',
        kelurahan: data[i][5] || '',
        kode_pos: data[i][6] || '',
        status: data[i][7] || 'Aktif',
        created_date: formatDateTimeForClient(data[i][8])
      });
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data wilayah: ' + error.message };
  }
}

function saveWilayah(wilayahData) {
  try {
    if (!wilayahData.nama_wilayah) {
      return { ok: false, message: 'Nama Wilayah wajib diisi' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_WILAYAH);
    
    var namaWilayah = sanitizeInput(wilayahData.nama_wilayah);
    var provinsi = sanitizeInput(wilayahData.provinsi || '');
    var kabupatenKota = sanitizeInput(wilayahData.kabupaten_kota || '');
    var kecamatan = sanitizeInput(wilayahData.kecamatan || '');
    var kelurahan = sanitizeInput(wilayahData.kelurahan || '');
    var kodePos = sanitizeInput(wilayahData.kode_pos || '');
    var status = sanitizeInput(wilayahData.status || 'Aktif');
    
    if (wilayahData.id_wilayah) {
      // Update existing
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() == wilayahData.id_wilayah) {
          sheet.getRange(i, 2).setValue(namaWilayah);
          sheet.getRange(i, 3).setValue(provinsi);
          sheet.getRange(i, 4).setValue(kabupatenKota);
          sheet.getRange(i, 5).setValue(kecamatan);
          sheet.getRange(i, 6).setValue(kelurahan);
          sheet.getRange(i, 7).setValue(kodePos);
          sheet.getRange(i, 8).setValue(status);
          return { ok: true, message: 'Wilayah berhasil diperbarui' };
        }
      }
      return { ok: false, message: 'Wilayah tidak ditemukan' };
    } else {
      // Insert new
      var newId = getNextId('WLY-', TAB_WILAYAH, 1);
      var newRow = [
        newId,
        namaWilayah,
        provinsi,
        kabupatenKota,
        kecamatan,
        kelurahan,
        kodePos,
        status,
        new Date()
      ];
      sheet.appendRow(newRow);
      return { ok: true, message: 'Wilayah berhasil disimpan', id: newId };
    }
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan wilayah: ' + error.message };
  }
}

function deleteWilayah(idWilayah) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_WILAYAH);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idWilayah) {
        sheet.deleteRow(i);
        
        // Also delete related ongkir
        var ongkirSheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
        var ongkirLastRow = ongkirSheet.getLastRow();
        for (var j = ongkirLastRow; j >= 2; j--) {
          if (ongkirSheet.getRange(j, 2).getValue() == idWilayah) {
            ongkirSheet.deleteRow(j);
          }
        }
        
        return { ok: true, message: 'Wilayah berhasil dihapus' };
      }
    }
    
    return { ok: false, message: 'Wilayah tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menghapus wilayah: ' + error.message };
  }
}

// ============================================
// ONGKOS KIRIM FUNCTIONS
// ============================================
function getOngkosKirimData(wilayahId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
    var result = [];
    
    if (sheet && sheet.getLastRow() > 1) {
      var lastRow = sheet.getLastRow();
      var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      
      for (var i = 0; i < data.length; i++) {
        if (!wilayahId || data[i][1] == wilayahId) {
          result.push({
            id_ongkos: data[i][0],
            id_wilayah: data[i][1],
            nama_ekspedisi: data[i][2],
            harga: data[i][3] || 0,
            estimasi_waktu: data[i][4] || '',
            status: data[i][5] || 'Aktif',
            created_date: formatDateTimeForClient(data[i][6])
          });
        }
      }
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data ongkos kirim: ' + error.message };
  }
}

function saveOngkosKirim(ongkosData) {
  try {
    if (!ongkosData.id_wilayah) {
      return { ok: false, message: 'Wilayah wajib dipilih' };
    }
    if (!ongkosData.nama_ekspedisi) {
      return { ok: false, message: 'Nama Ekspedisi wajib diisi' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
    
    var idWilayah = sanitizeInput(ongkosData.id_wilayah);
    var namaEkspedisi = sanitizeInput(ongkosData.nama_ekspedisi);
    var harga = parseFloat(ongkosData.harga) || 0;
    var estimasiWaktu = sanitizeInput(ongkosData.estimasi_waktu || '');
    var status = sanitizeInput(ongkosData.status || 'Aktif');
    
    if (ongkosData.id_ongkos) {
      // Update existing
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() == ongkosData.id_ongkos) {
          sheet.getRange(i, 2).setValue(idWilayah);
          sheet.getRange(i, 3).setValue(namaEkspedisi);
          sheet.getRange(i, 4).setValue(harga);
          sheet.getRange(i, 5).setValue(estimasiWaktu);
          sheet.getRange(i, 6).setValue(status);
          return { ok: true, message: 'Ongkos Kirim berhasil diperbarui' };
        }
      }
      return { ok: false, message: 'Ongkos Kirim tidak ditemukan' };
    } else {
      // Insert new
      var newId = getNextId('ONG-', TAB_ONGKOS_KIRIM, 1);
      var newRow = [
        newId,
        idWilayah,
        namaEkspedisi,
        harga,
        estimasiWaktu,
        status,
        new Date()
      ];
      sheet.appendRow(newRow);
      return { ok: true, message: 'Ongkos Kirim berhasil disimpan', id: newId };
    }
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan ongkos kirim: ' + error.message };
  }
}

function deleteOngkosKirim(idOngkos) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idOngkos) {
        sheet.deleteRow(i);
        return { ok: true, message: 'Ongkos Kirim berhasil dihapus' };
      }
    }
    
    return { ok: false, message: 'Ongkos Kirim tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menghapus ongkos kirim: ' + error.message };
  }
}

function getWilayahOptions() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_WILAYAH);
    var result = [];
    
    if (sheet && sheet.getLastRow() > 1) {
      var lastRow = sheet.getLastRow();
      var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      
      for (var i = 0; i < data.length; i++) {
        if (data[i][7] === 'Aktif') {
          result.push({
            id: data[i][0],
            nama: data[i][1]
          });
        }
      }
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil opsi wilayah: ' + error.message };
  }
}

// ============================================
// DASHBOARD DATA
// ============================================
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var result = {
      totalCustomer: 0,
      totalLead: 0,
      totalProduk: 0,
      totalWilayah: 0,
      hotProspek: 0,
      dealBulanIni: 0
    };
    
    // CRM stats
    var crmSheet = ss.getSheetByName(TAB_CRM);
    if (crmSheet && crmSheet.getLastRow() > 1) {
      var lastRow = crmSheet.getLastRow();
      var data = crmSheet.getRange(2, 1, lastRow - 1, 15).getValues();
      var now = new Date();
      var currentMonth = now.getMonth();
      var currentYear = now.getFullYear();
      
      for (var i = 0; i < data.length; i++) {
        var kategori = data[i][11];
        var tag = data[i][4];
        var status = data[i][10];
        var tanggalPembelian = data[i][13];
        
        if (kategori === 'Customer') result.totalCustomer++;
        else if (kategori === 'Lead') result.totalLead++;
        
        if (tag === 'Hot') result.hotProspek++;
        
        if (tanggalPembelian) {
          var purchaseDate = new Date(tanggalPembelian);
          if (purchaseDate.getMonth() === currentMonth && purchaseDate.getFullYear() === currentYear && status === 'Deal') {
            result.dealBulanIni++;
          }
        }
      }
    }
    
    // Produk stats
    var produkSheet = ss.getSheetByName(TAB_PRODUK);
    if (produkSheet) {
      result.totalProduk = Math.max(0, produkSheet.getLastRow() - 1);
    }
    
    // Wilayah stats
    var wilayahSheet = ss.getSheetByName(TAB_WILAYAH);
    if (wilayahSheet) {
      result.totalWilayah = Math.max(0, wilayahSheet.getLastRow() - 1);
    }
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data dashboard: ' + error.message };
  }
}

// ============================================
// SUMBER LEAD OPTIONS
// ============================================
function getSumberLeadOptions() {
  return {
    ok: true,
    data: [
      'Media Sosial',
      'Teman',
      'Event',
      'Iklan',
      'Website',
      'Referral',
      'Walk-in',
      'Telemarketing',
      'Lainnya'
    ]
  };
}

// ============================================
// TRANSAKSI FUNCTIONS
// ============================================
function getTransaksiData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_TRANSAKSI);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: [], summary: getTransaksiSummary([]) };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      result.push({
        id_transaksi: data[i][0],
        id_crm: data[i][1] || '',
        nama_customer: data[i][2] || '',
        nomor_whatsapp: data[i][3] || '',
        alamat_pengiriman: data[i][4] || '',
        tanggal_transaksi: formatDateForClient(data[i][5]),
        produk: data[i][6] || '',
        jumlah: parseInt(data[i][7]) || 0,
        harga_tier: data[i][8] || 1,
        harga_satuan: parseFloat(data[i][9]) || 0,
        subtotal: parseFloat(data[i][10]) || 0,
        ongkir: parseFloat(data[i][11]) || 0,
        total: parseFloat(data[i][12]) || 0,
        ekspedisi: data[i][13] || '',
        status_pembayaran: data[i][14] || 'Belum Lunas',
        created_date: formatDateTimeForClient(data[i][15])
      });
    }
    
    return { ok: true, data: result, summary: getTransaksiSummary(result) };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil data transaksi: ' + error.message };
  }
}

function getTransaksiSummary(data) {
  var summary = {
    totalTransaksi: data.length,
    totalPendapatan: 0,
    lunas: 0,
    belumLunas: 0,
    bulanIni: 0
  };
  
  var now = new Date();
  var currentMonth = now.getMonth();
  var currentYear = now.getFullYear();
  
  for (var i = 0; i < data.length; i++) {
    summary.totalPendapatan += parseFloat(data[i].total) || 0;
    
    if (data[i].status_pembayaran === 'Lunas') {
      summary.lunas++;
    } else {
      summary.belumLunas++;
    }
    
    if (data[i].tanggal_transaksi) {
      var transDate = new Date(data[i].tanggal_transaksi);
      if (transDate.getMonth() === currentMonth && transDate.getFullYear() === currentYear) {
        summary.bulanIni++;
      }
    }
  }
  
  return summary;
}

function saveTransaksi(transaksiData) {
  try {
    if (!transaksiData.nama_customer) {
      return { ok: false, message: 'Nama Customer wajib diisi' };
    }
    if (!transaksiData.produk) {
      return { ok: false, message: 'Produk wajib diisi' };
    }
    if (!transaksiData.jumlah || transaksiData.jumlah <= 0) {
      return { ok: false, message: 'Jumlah wajib diisi' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_TRANSAKSI);
    
    var namaCustomer = sanitizeInput(transaksiData.nama_customer);
    var nomorWhatsApp = sanitizeInput(transaksiData.nomor_whatsapp || '').replace(/[\s\-\.\,]/g, '');
    var alamatPengiriman = sanitizeInput(transaksiData.alamat_pengiriman || '');
    var tanggalTransaksi = transaksiData.tanggal_transaksi || null;
    var produk = sanitizeInput(transaksiData.produk);
    var jumlah = parseInt(transaksiData.jumlah) || 1;
    var hargaTier = parseInt(transaksiData.harga_tier) || 1;
    var hargaSatuan = parseFloat(transaksiData.harga_satuan) || 0;
    var subtotal = parseFloat(transaksiData.subtotal) || (jumlah * hargaSatuan);
    var ongkir = parseFloat(transaksiData.ongkir) || 0;
    var total = parseFloat(transaksiData.total) || (subtotal + ongkir);
    var ekspedisi = sanitizeInput(transaksiData.ekspedisi || '');
    var statusPembayaran = sanitizeInput(transaksiData.status_pembayaran || 'Belum Lunas');
    
    if (transaksiData.id_transaksi) {
      // Update existing
      var lastRow = sheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() == transaksiData.id_transaksi) {
          sheet.getRange(i, 3).setValue(namaCustomer);
          sheet.getRange(i, 4).setValue(nomorWhatsApp);
          sheet.getRange(i, 5).setValue(alamatPengiriman);
          sheet.getRange(i, 6).setValue(tanggalTransaksi ? new Date(tanggalTransaksi) : null);
          sheet.getRange(i, 7).setValue(produk);
          sheet.getRange(i, 8).setValue(jumlah);
          sheet.getRange(i, 9).setValue(hargaTier);
          sheet.getRange(i, 10).setValue(hargaSatuan);
          sheet.getRange(i, 11).setValue(subtotal);
          sheet.getRange(i, 12).setValue(ongkir);
          sheet.getRange(i, 13).setValue(total);
          sheet.getRange(i, 14).setValue(ekspedisi);
          sheet.getRange(i, 15).setValue(statusPembayaran);
          return { ok: true, message: 'Transaksi berhasil diperbarui' };
        }
      }
      return { ok: false, message: 'Transaksi tidak ditemukan' };
    } else {
      // Insert new
      var newId = getNextId('TRX-', TAB_TRANSAKSI, 1);
      var newRow = [
        newId,
        transaksiData.id_crm || '',
        namaCustomer,
        nomorWhatsApp,
        alamatPengiriman,
        tanggalTransaksi ? new Date(tanggalTransaksi) : null,
        produk,
        jumlah,
        hargaTier,
        hargaSatuan,
        subtotal,
        ongkir,
        total,
        ekspedisi,
        statusPembayaran,
        new Date()
      ];
      sheet.appendRow(newRow);
      return { ok: true, message: 'Transaksi berhasil disimpan', id: newId };
    }
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan transaksi: ' + error.message };
  }
}

function deleteTransaksi(idTransaksi) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_TRANSAKSI);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == idTransaksi) {
        sheet.deleteRow(i);
        return { ok: true, message: 'Transaksi berhasil dihapus' };
      }
    }
    
    return { ok: false, message: 'Transaksi tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menghapus transaksi: ' + error.message };
  }
}

function getTransaksiOptions() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var crmSheet = ss.getSheetByName(TAB_CRM);
    var produkSheet = ss.getSheetByName(TAB_PRODUK);
    var wilayahSheet = ss.getSheetByName(TAB_WILAYAH);
    var ongkirSheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
    
    var result = { customers: [], produk: [], ekspedisi: [] };
    
    // Get customers (only Customer category)
    if (crmSheet && crmSheet.getLastRow() > 1) {
      var crmData = crmSheet.getRange(2, 1, crmSheet.getLastRow() - 1, 12).getValues();
      for (var i = 0; i < crmData.length; i++) {
        if (crmData[i][11] === 'Customer') {
          result.customers.push({
            id: crmData[i][0],
            nama: crmData[i][1],
            whatsapp: crmData[i][2] || '',
            alamat: ''
          });
        }
      }
    }
    
    // Get produk
    if (produkSheet && produkSheet.getLastRow() > 1) {
      var produkData = produkSheet.getRange(2, 1, produkSheet.getLastRow() - 1, 14).getValues();
      for (var i = 0; i < produkData.length; i++) {
        if (produkData[i][16] === 'Aktif') {
          result.produk.push({
            id: produkData[i][0],
            nama: produkData[i][1],
            tier_1: produkData[i][4] || 0,
            tier_2: produkData[i][5] || 0,
            tier_3: produkData[i][6] || 0,
            tier_4: produkData[i][7] || 0,
            tier_5: produkData[i][8] || 0,
            tier_6: produkData[i][9] || 0,
            tier_7: produkData[i][10] || 0,
            tier_8: produkData[i][11] || 0,
            tier_9: produkData[i][12] || 0
          });
        }
      }
    }
    
    // Get unique ekspedisi options
    var ekspedisiSet = {};
    if (ongkirSheet && ongkirSheet.getLastRow() > 1) {
      var ongkirData = ongkirSheet.getRange(2, 1, ongkirSheet.getLastRow() - 1, 6).getValues();
      for (var i = 0; i < ongkirData.length; i++) {
        if (ongkirData[i][2]) {
          ekspedisiSet[ongkirData[i][2]] = true;
        }
      }
    }
    result.ekspedisi = Object.keys(ekspedisiSet);
    
    return { ok: true, data: result };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil opsi transaksi: ' + error.message };
  }
}

function getOngkirByWilayah(wilayahId, ekspedisi) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_ONGKOS_KIRIM);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, data: 0 };
    }
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][1] == wilayahId && data[i][2] === ekspedisi) {
        return { ok: true, data: parseFloat(data[i][3]) || 0 };
      }
    }
    
    return { ok: true, data: 0 };
  } catch (error) {
    return { ok: false, message: 'Gagal mengambil ongkir: ' + error.message };
  }
}

function saveBuktiTransfer(transaksiId, base64Data, fileName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_TRANSAKSI);
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() == transaksiId) {
        // Save bukti transfer URL/note to column 18 (after created_date at column 16)
        // Column 17: produk_list
        // Column 18: bukti_transfer
        // Column 19: biaya_cod
        
        // For now, store the file name as reference
        sheet.getRange(i, 18).setValue(fileName);
        
        return { ok: true, message: 'Bukti transfer berhasil disimpan' };
      }
    }
    
    return { ok: false, message: 'Transaksi tidak ditemukan' };
  } catch (error) {
    return { ok: false, message: 'Gagal menyimpan bukti transfer: ' + error.message };
  }
}
