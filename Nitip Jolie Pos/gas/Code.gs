/**
 * POS Web App – backend Google Apps Script
 * Database: Google Sheets (Produk, Pesanan, Pengaturan)
 * Langkah setup: lihat Panduan_POS_WhatsApp.txt
 */

/* ---------- API JSON (untuk frontend yang dihosting terpisah) ---------- */
// Hanya fungsi di daftar ini yang bisa dipanggil dari luar. Fungsi admin tetap dijaga PIN di dalamnya.
var API_FUNCS_ = {
  getData: getData, saveOrder: saveOrder, orderStatus: orderStatus, quoteDelivery: quoteDelivery,
  listOrders: listOrders, getAdminProducts: getAdminProducts, setStatus: setStatus,
  saveSettings: saveSettings, uploadImage: uploadImage, saveProduct: saveProduct,
  setProductActive: setProductActive, setBrandActive: setBrandActive, deleteProduct: deleteProduct,
  adminLogin: adminLogin, checkCart: checkCart, orderReceipt: orderReceipt,
  getReport: getReport, listCustomers: listCustomers
};
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function api_(fn, args) {
  if (typeof fn !== 'string' || !Object.prototype.hasOwnProperty.call(API_FUNCS_, fn)) return { ok: false, error: 'Fungsi tidak dikenal' };
  if (!Array.isArray(args) || args.length > 6) return { ok: false, error: 'Argumen tidak valid' };
  try { return { ok: true, data: API_FUNCS_[fn].apply(null, args) }; }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}
// Frontend memakai POST dengan Content-Type text/plain agar tidak memicu preflight CORS.
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Permintaan tidak valid' }); }
  return json_(api_(body && body.fn, body && body.args));
}

function doGet(e) {
  // Tanpa parameter: tetap menyajikan halaman HTML bawaan (kompatibel dengan deployment lama).
  var out;
  try { out = HtmlService.createHtmlOutputFromFile('Index'); }
  catch (err) { out = HtmlService.createHtmlOutputFromFile('index'); }
  return out
    .setTitle('POS • Pesan Online')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- helpers ---------- */
var requestSpreadsheet_;
function ss_() {
  if(requestSpreadsheet_)return requestSpreadsheet_;
  var id = PropertiesService.getScriptProperties().getProperty('POS_SPREADSHEET_ID');
  if (!id) throw new Error('Setup belum dijalankan. Buka Google Sheets → Extensions → Apps Script, jalankan setupPOS.');
  return requestSpreadsheet_ = SpreadsheetApp.openById(id);
}
function setupPOS() {
  // Bound-editor execution only: no active spreadsheet in a web app request.
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Jalankan setupPOS dari editor Apps Script yang dibuka melalui Google Sheets.');
  PropertiesService.getScriptProperties().setProperty('POS_SPREADSHEET_ID', active.getId());
  buildSheets_(false);
  var sh = active.getSheetByName('Pengaturan');
  var keys = sh.getRange(1,1,sh.getLastRow(),1).getValues().map(function(r){return r[0];});
  if (keys.indexOf('Ongkir') < 0) sh.appendRow(['Ongkir', 0]);
  if (keys.indexOf('BatasBayarMenit') < 0) sh.appendRow(['BatasBayarMenit', 15]);
  if (keys.indexOf('QRISPayload') < 0) sh.appendRow(['QRISPayload', '']);
  // Tarif ongkir ala ojek online (bisa diubah dari Pengaturan tanpa menyentuh kode).
  if (keys.indexOf('TarifDasar') < 0) sh.appendRow(['TarifDasar', 8000]);
  if (keys.indexOf('JarakDasar') < 0) sh.appendRow(['JarakDasar', 2]);
  if (keys.indexOf('TarifPerKm') < 0) sh.appendRow(['TarifPerKm', 2000]);
  if (keys.indexOf('MaksJarakKm') < 0) sh.appendRow(['MaksJarakKm', 20]);
  if (keys.indexOf('TokoLat') < 0) sh.appendRow(['TokoLat', '']);
  if (keys.indexOf('TokoLng') < 0) sh.appendRow(['TokoLng', '']);
  active.getSheetByName('Pesanan').getRange('D:D').setNumberFormat('@');
  sh.getRange('B:B').setNumberFormat('@');
  ensureSheets_();
  ensureReceiptColumns_();
  // Column R (NomorWhatsApp) must stay text so the number is never turned into a numeric value.
  active.getSheetByName('Pesanan').getRange('R:R').setNumberFormat('@');
  ensureBrandColumn_();
  ensureStockColumn_();
  ensureCostColumn_();
  ensureCustomerSheet_();
  invalidateCatalog_();
  try { installExpiryTrigger(); } catch (e) { console.warn('Trigger pembatalan otomatis belum terpasang: ' + e.message + '. Jalankan installExpiryTrigger secara manual.'); }
  console.log('Setup selesai. Isi PIN admin di Pengaturan (minimal 8 karakter). PIN kosong menonaktifkan admin web.');
}
function safeText_(value) {
  var s = String(value == null ? '' : value);
  return /^[=+@\-]/.test(s) ? "'" + s : s;
}
function money_(value, label) {
  var n = Number(value || 0);
  if (!isFinite(n) || n < 0 || n > 1000000000 || Math.floor(n) !== n) throw new Error(label + ' harus berupa rupiah bulat, 0–1 miliar');
  return n;
}

function rows_(name) {
  var sh=ss_().getSheetByName(name);if(!sh||sh.getLastRow()<2)return [];
  var width=name==='Pengaturan'?2:name==='Pesanan'?19:sh.getLastColumn();
  var values=sh.getRange(1,1,sh.getLastRow(),width).getValues(),headers=values.shift();
  return objectRows_(headers,values);
}
function objectRows_(headers,values){
  return values.filter(function(r){return r[0]!=='';}).map(function(r){var o={};headers.forEach(function(k,i){o[k]=r[i] instanceof Date?r[i].toISOString():r[i];});return o;});
}
function invalidateCatalog_(){PropertiesService.getScriptProperties().setProperty('POS_CATALOG_REV',Utilities.getUuid());}
function onEdit(e){if(e&&e.range&&['Produk','Pengaturan'].indexOf(e.range.getSheet().getName())>=0)invalidateCatalog_();}
function pageQuery_(q){q=q||{};return {offset:Math.max(0,Math.floor(Number(q.offset)||0)),limit:Math.min(50,Math.max(1,Math.floor(Number(q.limit)||24))),brand:typeof q.brand==='string'?q.brand:null,category:typeof q.category==='string'?q.category:null};}
function productCompare_(a,b){return (a.brand||'').localeCompare(b.brand||'','id')||a.nama.localeCompare(b.nama,'id')||a.id.localeCompare(b.id);}
function productPage_(all,q){
  q=pageQuery_(q);all.sort(productCompare_);
  var brands=Array.from(new Set(all.map(function(p){return p.brand;}))).sort(function(a,b){return a.localeCompare(b,'id');});
  var categories=Array.from(new Set(all.map(function(p){return p.kategori;}).filter(Boolean))).sort();
  var filtered=all.filter(function(p){return (q.brand===null||p.brand===q.brand)&&(q.category===null||p.kategori===q.category);});
  return {produk:filtered.slice(q.offset,q.offset+q.limit),brands:brands,categories:categories,total:filtered.length,nextOffset:q.offset+q.limit<filtered.length?q.offset+q.limit:null};
}

// Dibaca dari sheet sekali per eksekusi; pemanggil mendapat salinan (getData menghapus PIN dari hasilnya).
var SETTINGS_MEMO_ = null;
function settings_() {
  if (!SETTINGS_MEMO_) {
    var s = {};
    rows_('Pengaturan').forEach(function (r) { s[r.Kunci] = r.Nilai; });
    SETTINGS_MEMO_ = s;
  }
  var copy = {};
  for (var k in SETTINGS_MEMO_) copy[k] = SETTINGS_MEMO_[k];
  return copy;
}

// Admin: jika PIN diisi di Pengaturan, semua fungsi admin wajib PIN yang benar ATAU token sesi dari adminLogin.
// Token disimpan di CacheService (6 jam) dan terikat pada hash PIN: mengganti PIN otomatis mencabut semua sesi.
var ADMIN_TOKEN_TTL_ = 21600;
function pinHash_(pin) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'pos-admin:' + pin)).slice(0, 22);
}
// Tebakan salah diperlambat makin lama makin besar (penghitung bersama, maks 15 dtk). PIN benar tidak pernah diperlambat,
// jadi pemilik toko tidak bisa dikunci dari luar, tetapi tebakan paralel menjadi tidak praktis.
function failDelay_() {
  var c = CacheService.getScriptCache(), n = Number(c.get('POS_PINFAIL') || 0) + 1;
  try { c.put('POS_PINFAIL', String(n), 600); } catch (ignore) {}
  Utilities.sleep(Math.min(15000, 1000 + n * 500));
}
function auth_(secret) {
  var p = String(settings_().PIN || '').trim();
  if (p.length < 8) throw new Error('Admin belum diaktifkan. Isi PIN minimal 8 karakter pada sheet Pengaturan.');
  var s = String(secret || '').trim();
  if (/^tk_[0-9a-f]{64}$/.test(s)) {
    var hit = CacheService.getScriptCache().get('POS_ADM_' + s);
    if (hit && hit === pinHash_(p)) return;
    throw new Error('Sesi admin berakhir. Masuk lagi dengan PIN.');
  }
  if (s !== p) { failDelay_(); throw new Error('PIN salah'); }
}

/* ---------- API publik ---------- */
function getData(query) {
  var q=pageQuery_(query),rev=PropertiesService.getScriptProperties().getProperty('POS_CATALOG_REV')||'0';
  var key='POS_V5_'+rev+'_'+productVersion_(q),cache=CacheService.getScriptCache(),hit=cache.get(key);
  if(hit){try{return JSON.parse(hit);}catch(ignore){}}
  var settings=settings_(),hasPin=String(settings.PIN||'').trim().length>=8;delete settings.PIN;
  var result=productPage_(catalog_(false),q);result.produk=result.produk.map(publicProduct_);result.settings=settings;result.hasPin=hasPin;
  var json=JSON.stringify(result);
  // Cache is optional. Oversized payloads and evictions fall back to Sheets.
  if(Utilities.newBlob(json).getBytes().length<90000){try{cache.put(key,json,60);}catch(ignore){}}
  return result;
}
// Harga dasar (modal) hanya untuk admin: tidak pernah ikut ke katalog publik maupun struk pelanggan.
function publicProduct_(p){var c={};Object.keys(p).forEach(function(k){if(k!=='hargaDasar')c[k]=p[k];});return c;}
function productFromRow_(p){
  var product={id:String(p.ID),brand:String(p.Brand||'').trim(),kategori:String(p.Kategori||''),nama:String(p.Nama||''),deskripsi:String(p.Deskripsi||''),harga:money_(p.Harga,'Harga produk'),hargaCoret:Number(p.HargaCoret)||0,suka:Number(p.Suka)||0,aktif:String(p.Aktif).toLowerCase()!=='tidak',gambar:String(p.Gambar||''),grup:[]};
  // Stok kosong = tidak dibatasi. Nilai yang bukan angka >= 0 dianggap tidak dibatasi.
  var stok=(p.Stok===''||p.Stok==null)?null:Math.floor(Number(p.Stok));
  if(stok!==null&&(!isFinite(stok)||stok<0))stok=null;
  product.stok=stok;product.habis=stok!==null&&stok<=0;
  var hd=(p.HargaDasar===''||p.HargaDasar==null)?null:Number(p.HargaDasar);
  if(hd!==null&&(!isFinite(hd)||hd<0))hd=null;
  product.hargaDasar=hd;
  product.version=productVersion_(product);return product;
}
function productById_(id){
  var sh=ss_().getSheetByName('Produk');if(sh.getLastRow()<2)return null;
  var match=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();if(!match)return null;
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],v=sh.getRange(match.getRow(),1,1,9).getValues()[0],p={};
  headers.slice(0,9).forEach(function(k,i){p[k]=v[i];});var col=headers.indexOf('Brand');p.Brand=col>=0?sh.getRange(match.getRow(),col+1).getValue():'';
  var scol=headers.indexOf('Stok');p.Stok=scol>=0?sh.getRange(match.getRow(),scol+1).getValue():'';
  var ccol=headers.indexOf('HargaDasar');p.HargaDasar=ccol>=0?sh.getRange(match.getRow(),ccol+1).getValue():'';
  return productFromRow_(p);
}
function catalog_(includeInactive){
  var sh=ss_().getSheetByName('Produk');if(!sh)throw new Error('Sheet Produk belum ada. Jalankan setupPOS.');
  if(sh.getLastRow()<2)return [];
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],brandCol=headers.indexOf('Brand');
  // Only the nine product columns plus Brand; unrelated columns are not read.
  var values=sh.getRange(2,1,sh.getLastRow()-1,9).getValues();
  var brands=brandCol>=9?sh.getRange(2,brandCol+1,values.length,1).getValues():null;
  var stokCol=headers.indexOf('Stok'),stoks=stokCol>=9?sh.getRange(2,stokCol+1,values.length,1).getValues():null;
  var costCol=headers.indexOf('HargaDasar'),costs=costCol>=9?sh.getRange(2,costCol+1,values.length,1).getValues():null;
  return values.map(function(r,i){
    if(r[0]==='')return null;var p={};headers.slice(0,9).forEach(function(k,j){p[k]=r[j];});p.Brand=brands?brands[i][0]:'';p.Stok=stoks?stoks[i][0]:'';p.HargaDasar=costs?costs[i][0]:'';
    // One malformed row (e.g. price typed as "Rp 25.000") must not take the whole store down.
    try{return productFromRow_(p);}catch(err){console.warn('Baris produk '+(i+2)+' dilewati: '+err.message);return null;}
  }).filter(function(p){return p&&(includeInactive||p.aktif);}).sort(productCompare_);
}

function saveOrder(o) {
  // Panggilan jaringan (OSRM, sampai ~20 dtk) dilakukan SEBELUM mengambil lock agar tidak memblokir pesanan lain.
  var preQuote = null;
  if (o && typeof o === 'object' && o.tipe === 'Dikirim') {
    var preSettings = settings_();
    if (deliveryConfig_(preSettings).enabled) preQuote = quoteFor_(o.lat, o.lng, preSettings);
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    if (!o || typeof o !== 'object') throw new Error('Data pesanan tidak valid');
    if (!String(o.nama || '').trim() || String(o.nama).length > 100) throw new Error('Nama wajib diisi, maksimal 100 karakter');
    if (!/^62\d{8,13}$/.test(String(o.telepon || ''))) throw new Error('Nomor telepon tidak valid');
    if (['Diambil','Dikirim'].indexOf(o.tipe) < 0) throw new Error('Tipe pengantaran tidak valid');
    if (o.tipe === 'Dikirim' && !String(o.alamat || '').trim()) throw new Error('Alamat wajib diisi');
    if (String(o.alamat || '').length > 1000) throw new Error('Alamat terlalu panjang');
    var catatanPesanan = String(o.catatan || '').trim();
    if (catatanPesanan.length > 500) throw new Error('Catatan maksimal 500 karakter');
    if (!/^[a-zA-Z0-9-]{16,80}$/.test(String(o.requestId || ''))) throw new Error('ID permintaan tidak valid. Muat ulang aplikasi.');
    ensureReceiptColumns_();
    var sh = ss_().getSheetByName('Pesanan');
    var now = new Date();
    var id = 'INV-' + o.requestId;
    // Deterministic ID + lock: a retry never writes a duplicate order.
    var found = sh.getRange(1,1,Math.max(sh.getLastRow(),1),1).createTextFinder(id).matchEntireCell(true).matchCase(true).findNext();
    if (found) {
      return receiptFromRow_(sh.getRange(found.getRow(),1,1,19).getValues()[0], settings_());
    }

    var jadwal = validateSchedule_(o.jadwal);

    // Total dihitung ulang di server (tidak mempercayai angka dari browser)
    var data = {settings:settings_(),produk:catalog_(false)};
    var biayaLayanan = 0; // biaya layanan dihapus; kolom BiayaLayanan di sheet dipertahankan (selalu 0)
    var subtotal = 0;
    if (!Array.isArray(o.items) || !o.items.length || o.items.length > 50) throw new Error('Keranjang kosong');
    var wantQty = {};
    var items = o.items.map(function (it) {
      var p = data.produk.filter(function (x) { return x.id === String(it.id); })[0];
      if (!p) throw new Error('Produk tidak ditemukan: ' + it.id);
      var tambahan=0,detail=[];
      if(it.pilihan && (!Array.isArray(it.pilihan)||it.pilihan.length))throw new Error('Varian tidak lagi digunakan. Muat ulang katalog dan tambahkan ulang produk.');
      var qty = Number(it.qty);
      if (!Number.isInteger(qty) || qty < 1 || qty > 100) throw new Error('Jumlah harus 1–100');
      if (String(it.catatan || '').length > 500) throw new Error('Catatan maksimal 500 karakter');
      wantQty[p.id] = (wantQty[p.id] || 0) + qty;
      if (p.stok !== null && p.stok !== undefined) {
        if (p.habis) throw new Error(p.nama + ' sudah habis. Hapus dari keranjang.');
        if (wantQty[p.id] > p.stok) throw new Error('Stok ' + p.nama + ' tinggal ' + p.stok + '. Kurangi jumlahnya.');
      }
      var total = (p.harga + tambahan) * qty;
      subtotal += total;
      return { id: p.id, nama: p.nama, qty: qty, harga: p.harga, hpp: p.hargaDasar, hargaCoret: p.hargaCoret, detail: detail,
               catatan: String(it.catatan || ''), bawa: !!it.bawa, total: total };
    });
    // Ongkir selalu dihitung ulang di server dari titik peta; angka dari browser tidak dipercaya.
    var ongkir = 0, alamatFinal = String(o.alamat || '');
    if (o.tipe === 'Dikirim') {
      if (deliveryConfig_(data.settings).enabled) {
        var quote = preQuote || quoteFor_(o.lat, o.lng, data.settings);
        ongkir = quote.ongkir;
        alamatFinal += '\nTitik peta: https://www.google.com/maps?q=' + quote.lat + ',' + quote.lng + '\nJarak: ' + quote.km + ' km';
      } else {
        ongkir = money_(data.settings.Ongkir, 'Ongkir');
      }
    }
    var total = subtotal + ongkir;

    var metode = o.metode;
    if (['Tunai','QRIS'].indexOf(metode) < 0) throw new Error('Metode tidak valid');
    if (metode === 'QRIS' && !data.settings.QRISGambar && !data.settings.QRISPayload) throw new Error('QRIS belum dikonfigurasi toko');
    var dibayar = 0, kembalian = 0, status = 'Menunggu Pembayaran';
    var tipe = o.tipe === 'Dikirim' ? 'Dikirim' : 'Diambil';

    var row = [id, now, String(o.nama).trim(), String(o.telepon), tipe, alamatFinal,
      metode, status, subtotal, biayaLayanan, ongkir, total, dibayar, kembalian,
      catatanPesanan || items.map(function (i) { return i.catatan; }).filter(String).join('; '), JSON.stringify(items)];
    row[17] = data.settings.WhatsApp;
    row[18] = jadwal;
    var result = receiptFromRow_(row, data.settings);
    if (row[15].length > 45000 || result.waText.length > 45000) throw new Error('Pesanan terlalu panjang. Kurangi item atau catatan.');
    row[16] = result.waText; row[17] = result.waNumber;
    sh.appendRow(row.map(function(v){return typeof v === 'string' ? safeText_(v) : v;}));
    adjustStock_(items, -1);
    // Data pelanggan disimpan/diperbarui per nomor telepon; kegagalan di sini tidak boleh menggagalkan pesanan.
    try {
      upsertCustomer_({ telepon: String(o.telepon), nama: String(o.nama).trim(), alamat: o.tipe === 'Dikirim' ? String(o.alamat || '').trim() : '',
        lat: quote ? quote.lat : o.lat, lng: quote ? quote.lng : o.lng }, now);
    } catch (e) { console.warn('Pelanggan tidak tersimpan: ' + e.message); }
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

// Publik: validasi keranjang yang dipulihkan dari perangkat pembeli (harga, ketersediaan, stok terbaru).
function checkCart(ids) {
  if (!Array.isArray(ids) || ids.length > 50) throw new Error('Data keranjang tidak valid');
  var want = {}, out = {};
  ids.forEach(function (i) { want[String(i)] = 1; });
  catalog_(true).forEach(function (p) {
    if (want[p.id]) out[p.id] = { nama: p.nama, harga: p.harga, hargaCoret: p.hargaCoret, aktif: p.aktif, stok: p.stok, habis: p.habis };
  });
  return out;
}

/* ---------- Ongkir ala ojek online ---------- */
// Tarif = TarifDasar untuk JarakDasar km pertama + TarifPerKm untuk setiap km berikutnya.
// Jarak selalu dibulatkan ke atas (2,1 km dihitung 3 km). Semua angka diatur di sheet Pengaturan.
function deliveryConfig_(settings) {
  var s = settings || settings_();
  function has(v) { return String(v == null ? '' : v).trim() !== '' && isFinite(Number(v)); }
  return {
    enabled: has(s.TokoLat) && has(s.TokoLng),
    lat: Number(s.TokoLat), lng: Number(s.TokoLng),
    tarifDasar: Number(s.TarifDasar) || 8000,
    jarakDasar: Number(s.JarakDasar) || 2,
    tarifPerKm: Number(s.TarifPerKm) || 2000,
    maksKm: Number(s.MaksJarakKm) || 20
  };
}
function calculateDeliveryCost(distanceKm, settings) {
  var c = deliveryConfig_(settings);
  var km = Math.max(1, Math.ceil(Number(distanceKm) - 1e-6));
  var extra = Math.max(0, km - c.jarakDasar);
  return { km: km, ongkir: c.tarifDasar + extra * c.tarifPerKm };
}
function haversineKm_(lat1, lng1, lat2, lng2) {
  var r = Math.PI / 180, dLat = (lat2 - lat1) * r, dLng = (lng2 - lng1) * r;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * r) * Math.cos(lat2 * r) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 6371 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
// Jarak lewat jalan (OSRM, gratis). Jika gagal: garis lurus x 1,3 sebagai perkiraan.
function routeKm_(lat1, lng1, lat2, lng2) {
  try {
    var url = 'https://router.project-osrm.org/route/v1/driving/' + lng1 + ',' + lat1 + ';' + lng2 + ',' + lat2 + '?overview=false';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var j = JSON.parse(res.getContentText());
      if (j.code === 'Ok' && j.routes && j.routes[0] && isFinite(j.routes[0].distance)) return { km: j.routes[0].distance / 1000, sumber: 'jalan' };
    }
  } catch (e) { console.warn('OSRM gagal: ' + e.message); }
  return { km: haversineKm_(lat1, lng1, lat2, lng2) * 1.3, sumber: 'perkiraan' };
}
function parseCoord_(v, min, max) {
  var n = Number(v);
  if (v === '' || v == null || !isFinite(n) || n < min || n > max) throw new Error('Titik lokasi pengantaran tidak valid. Geser pin di peta.');
  return Math.round(n * 100000) / 100000;
}
function quoteFor_(lat, lng, settings) {
  var s = settings || settings_(), c = deliveryConfig_(s);
  if (!c.enabled) throw new Error('Lokasi toko belum diatur di sheet Pengaturan (TokoLat, TokoLng).');
  lat = parseCoord_(lat, -90, 90); lng = parseCoord_(lng, -180, 180);
  var cache = CacheService.getScriptCache(), key = 'POS_RT_' + lat + '_' + lng + '_' + c.lat + '_' + c.lng, route = null, hit = cache.get(key);
  if (hit) { try { route = JSON.parse(hit); } catch (ignore) {} }
  if (!route) { route = routeKm_(c.lat, c.lng, lat, lng); try { cache.put(key, JSON.stringify(route), 21600); } catch (ignore) {} }
  var cost = calculateDeliveryCost(route.km, s);
  if (c.maksKm > 0 && cost.km > c.maksKm) throw new Error('Lokasi di luar jangkauan pengantaran (maksimal ' + c.maksKm + ' km, jarak sekitar ' + cost.km + ' km).');
  return { ok: true, lat: lat, lng: lng, km: cost.km, ongkir: cost.ongkir, sumber: route.sumber, maksKm: c.maksKm };
}
// Publik: pratinjau ongkir untuk pembeli. Nilai final tetap dihitung ulang di saveOrder.
function quoteDelivery(lat, lng) { return quoteFor_(lat, lng); }
// Jalankan SEKALI dari editor agar Apps Script meminta izin akses URL eksternal (OSRM).
function authorizeExternal() { UrlFetchApp.fetch('https://router.project-osrm.org/route/v1/driving/110.9161,-7.0869;110.93,-7.09?overview=false'); }

/* ---------- API admin (PIN) ---------- */
function adminLogin(pin) {
  var s = String(pin || '').trim();
  if (/^tk_/.test(s)) { failDelay_(); throw new Error('PIN salah'); }
  auth_(s);
  var token = 'tk_' + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('POS_ADM_' + token, pinHash_(String(settings_().PIN || '').trim()), ADMIN_TOKEN_TTL_);
  try { CacheService.getScriptCache().remove('POS_PINFAIL'); } catch (ignore) {}
  return { ok: true, token: token, ttl: ADMIN_TOKEN_TTL_ };
}

function listOrders(pin, limit, beforeRow) {
  auth_(pin);expireOrdersLocked_();return orderPage_(limit,beforeRow);
}

/* ---------- Batas waktu pembayaran & pembatalan otomatis ---------- */
// Hanya QRIS yang kedaluwarsa (saveOrder hanya menerima Tunai dan QRIS). Tunai dibayar saat pesanan diambil/diterima,
// jadi tidak dibatalkan otomatis.
var EXPIRING_METHODS_ = ['QRIS'];
function payTimeoutMin_(settings) {
  var n = Math.floor(Number((settings || settings_()).BatasBayarMenit));
  return isFinite(n) && n >= 5 && n <= 120 ? n : 15;
}
function expireOrders_() {
  var sh = ss_().getSheetByName('Pesanan');
  if (!sh || sh.getLastRow() < 2) return 0;
  var last = sh.getLastRow(), start = Math.max(2, last - 499), n = last - start + 1;
  // Kolom B..H: Waktu(0) Nama Telepon Tipe Alamat Metode(5) Status(6). Hanya 500 baris terakhir yang diperiksa.
  var v = sh.getRange(start, 2, n, 7).getValues();
  var limit = payTimeoutMin_() * 60000, now = Date.now(), cancelled = 0;
  var out = v.map(function (r) { return [r[6]]; });
  var cancelledRows = [];
  for (var i = 0; i < v.length; i++) {
    if (v[i][6] === 'Menunggu Pembayaran' && EXPIRING_METHODS_.indexOf(v[i][5]) >= 0 &&
        v[i][0] instanceof Date && now - v[i][0].getTime() > limit) {
      out[i][0] = 'Dibatalkan';
      cancelled++;
      cancelledRows.push(start + i);
    }
  }
  if (cancelled) {
    // Satu tulisan batch untuk kolom Status; rumus yang ada dipertahankan.
    var rng = sh.getRange(start, 8, n, 1), fm = rng.getFormulas();
    rng.setValues(out.map(function (r, i) { return [out[i][0] === 'Dibatalkan' ? 'Dibatalkan' : (fm[i][0] || r[0])]; }));
    // Pesanan kedaluwarsa mengembalikan stok yang tadi dipotong.
    cancelledRows.forEach(function (rowNo) { adjustStock_(itemsOfRow_(sh, rowNo), +1); });
    SpreadsheetApp.flush();
  }
  return cancelled;
}
function expireOrdersLocked_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return expireOrders_(); } finally { lock.releaseLock(); }
}
// Dipanggil oleh trigger waktu (setiap 1 menit).
function cancelExpiredOrders() { return expireOrdersLocked_(); }
function installExpiryTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cancelExpiredOrders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cancelExpiredOrders').timeBased().everyMinutes(1).create();
}
function orderLocate_(id) {
  id = String(id || '');
  if (!/^INV-[a-zA-Z0-9-]{16,80}$/.test(id)) throw new Error('ID pesanan tidak valid');
  var sh = ss_().getSheetByName('Pesanan');
  var match = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).createTextFinder(id).matchEntireCell(true).matchCase(true).findNext();
  if (!match) throw new Error('Pesanan tidak ditemukan');
  return { sh: sh, row: match.getRow() };
}
// Kunci hanya diambil bila pesanan masih menunggu pembayaran QRIS (satu-satunya yang bisa kedaluwarsa).
// Polling status tiap beberapa detik dari banyak pelanggan tidak lagi saling antre.
function expireIfPending_(loc) {
  var cur = loc.sh.getRange(loc.row, 7, 1, 2).getValues()[0];
  if (cur[1] === 'Menunggu Pembayaran' && EXPIRING_METHODS_.indexOf(cur[0]) >= 0) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { expireOrders_(); } finally { lock.releaseLock(); }
  }
}
// Publik: status terbaru satu pesanan + sisa waktu menurut jam server (tidak bergantung jam HP pembeli).
function orderStatus(id) {
  var loc = orderLocate_(id);
  expireIfPending_(loc);
  var row = loc.sh.getRange(loc.row, 1, 1, 8).getValues()[0];
  var out = { status: row[7], tipe: row[4], sisaDetik: 0 };
  if (row[7] === 'Menunggu Pembayaran' && EXPIRING_METHODS_.indexOf(row[6]) >= 0) {
    out.sisaDetik = Math.max(0, Math.floor((new Date(row[1]).getTime() + payTimeoutMin_() * 60000 - Date.now()) / 1000));
  }
  return out;
}
// Publik: struk lengkap untuk "Lacak pesanan" setelah halaman ditutup. ID pesanan acak (UUID) bertindak sebagai kunci.
function orderReceipt(id) {
  var loc = orderLocate_(id);
  expireIfPending_(loc);
  return receiptFromRow_(loc.sh.getRange(loc.row, 1, 1, 19).getValues()[0], settings_());
}
function itemsOfRow_(sh, rowNo) {
  try { var items = JSON.parse(sh.getRange(rowNo, 16).getValue() || '[]'); return Array.isArray(items) ? items : []; }
  catch (e) { return []; }
}
// Potong (sign = -1) atau kembalikan (+1) stok. Produk tanpa angka stok dilewati. Dipanggil saat memegang lock.
function adjustStock_(items, sign) {
  var sh = ss_().getSheetByName('Produk');
  if (!sh || sh.getLastRow() < 2) return;
  var sc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('Stok');
  if (sc < 0) return;
  var changed = false;
  (items || []).forEach(function (it) {
    var qty = Math.floor(Number(it && it.qty));
    if (!it || !isFinite(qty) || qty < 1) return;
    var m = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(it.id)).matchEntireCell(true).matchCase(true).findNext();
    if (!m) return;
    var cell = sh.getRange(m.getRow(), sc + 1), cur = cell.getValue();
    if (cur === '' || cur == null) return;
    var n = Math.floor(Number(cur));
    if (!isFinite(n)) return;
    cell.setValue(Math.max(0, n + sign * qty));
    changed = true;
  });
  if (changed) { SpreadsheetApp.flush(); invalidateCatalog_(); }
}
function orderPage_(limit,beforeRow){
  var sh=ss_().getSheetByName('Pesanan'),size=Math.min(50,Math.max(1,Number(limit)||30));
  var end=Math.min(sh.getLastRow(),beforeRow?Math.floor(Number(beforeRow))-1:sh.getLastRow());
  if(end<2)return {orders:[],nextRow:null};
  var start=Math.max(2,end-size+1),headers=sh.getRange(1,1,1,19).getValues()[0];
  return {orders:objectRows_(headers,sh.getRange(start,1,end-start+1,19).getValues()).reverse(),nextRow:start>2?start:null};
}

function setStatus(pin, id, status) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return setStatus_(pin, id, status); } finally { lock.releaseLock(); }
}
var STATUSES_ = ['Menunggu Pembayaran', 'Lunas', 'Diproses', 'Siap', 'Selesai', 'Dibatalkan'];
var PAID_STATUSES_ = ['Lunas', 'Diproses', 'Siap', 'Selesai'];
function setStatus_(pin,id,status){
  auth_(pin);if(STATUSES_.indexOf(status)<0)throw new Error('Status tidak valid');expireOrders_();
  var sh=ss_().getSheetByName('Pesanan'),match=sh.getRange(1,1,sh.getLastRow(),1).createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
  if(!match)throw new Error('Data tidak ditemukan');
  var row=match.getRow(),range=sh.getRange(row,8,1,6),v=range.getValues()[0],formulas=range.getFormulas()[0];
  var oldStatus=v[0];
  if(['Diproses','Siap','Selesai'].indexOf(status)>=0&&PAID_STATUSES_.indexOf(oldStatus)<0)throw new Error('Konfirmasi pembayaran terlebih dahulu');
  var total=v[4];v=v.map(function(x,i){return formulas[i]||x;});v[0]=status;v[5]=(status==='Menunggu Pembayaran'||status==='Dibatalkan')?0:total;range.setValues([v]);
  // Pembatalan mengembalikan stok; membuka kembali pesanan yang dibatalkan memotongnya lagi.
  if(status==='Dibatalkan'&&oldStatus!=='Dibatalkan')adjustStock_(itemsOfRow_(sh,row),+1);
  else if(oldStatus==='Dibatalkan'&&status!=='Dibatalkan')adjustStock_(itemsOfRow_(sh,row),-1);
  return {ok:true};
}

var KEYS_ = ['NamaToko', 'LogoTeks', 'LogoURL', 'StrukLogoURL', 'BannerURL', 'WarnaUtama', 'Alamat', 'JamBuka', 'WhatsApp',
  'OutletLain', 'OutletLainInfo', 'Ongkir', 'Bank', 'NoRekening', 'AtasNama', 'QRISGambar', 'QRISPayload', 'BatasBayarMenit', 'TarifDasar', 'JarakDasar', 'TarifPerKm', 'MaksJarakKm', 'TokoLat', 'TokoLng', 'PIN'];
// Koordinat bisa negatif; jangan diberi awalan ' oleh safeText_.
var COORD_KEYS_ = ['TokoLat', 'TokoLng'];

// CRC16-CCITT-FALSE (poly 0x1021, init 0xFFFF) sesuai EMVCo/QRIS.
function emvCrc_(s) {
  var crc = 0xFFFF;
  for (var i = 0; i < s.length; i++) {
    crc ^= s.charCodeAt(i) << 8;
    for (var j = 0; j < 8; j++) crc = (crc & 0x8000) ? ((crc << 1) ^ 0x1021) & 0xFFFF : (crc << 1) & 0xFFFF;
  }
  return ('0000' + crc.toString(16).toUpperCase()).slice(-4);
}
function qrisPayloadOk_(s) {
  return /^000201[\x20-\x7E]{20,700}$/.test(s) && /6304[0-9A-F]{4}$/.test(s) &&
    emvCrc_(s.slice(0, -4)) === s.slice(-4);
}
var TEXT_KEYS_ = ['WhatsApp', 'NoRekening', 'PIN'];

function saveSettings(pin, obj) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return saveSettings_(pin, obj); } finally { lock.releaseLock(); }
}
function saveSettings_(pin,obj){
  auth_(pin);if(!obj||typeof obj!=='object')throw new Error('Pengaturan tidak valid');
  if(obj.PIN&&String(obj.PIN).trim().length<8)throw new Error('PIN minimal 8 karakter');
  if(obj.PIN&&/^[=+@\-]/.test(String(obj.PIN).trim()))throw new Error('PIN tidak boleh diawali karakter = + @ atau -');
  if('WhatsApp' in obj)obj.WhatsApp=normalizeWA_(obj.WhatsApp,true);
  if('QRISPayload' in obj){obj.QRISPayload=String(obj.QRISPayload==null?'':obj.QRISPayload).trim();
    if(obj.QRISPayload&&!qrisPayloadOk_(obj.QRISPayload))throw new Error('Kode QRIS tidak valid (format EMVCo atau CRC salah). Salin ulang teks hasil scan QRIS statis toko.');}
  if('BatasBayarMenit' in obj){
    if(String(obj.BatasBayarMenit==null?'':obj.BatasBayarMenit).trim()==='')delete obj.BatasBayarMenit;
    else{var bm=Number(obj.BatasBayarMenit);if(!Number.isInteger(bm)||bm<5||bm>120)throw new Error('Batas waktu bayar harus 5–120 menit');obj.BatasBayarMenit=bm;}
  }
  var sh=ss_().getSheetByName('Pengaturan'),range=sh.getRange(1,1,Math.max(1,sh.getLastRow()),2),values=range.getValues(),formulas=range.getFormulas(),index={};
  values.forEach(function(r,i){index[r[0]]=i;});
  var output=values.map(function(r,i){return r.map(function(v,j){return formulas[i][j]||(typeof v==='string'&&!(j===1&&COORD_KEYS_.indexOf(r[0])>=0)?safeText_(v):v);});});
  KEYS_.forEach(function(k){if(!(k in obj)||(k==='PIN'&&(obj[k]===''||obj[k]==null)))return;
    var val;
    if(['JarakDasar','MaksJarakKm','TarifDasar','TarifPerKm'].indexOf(k)>=0&&String(obj[k]==null?'':obj[k]).trim()==='')return;
    if(['Ongkir','TarifDasar','TarifPerKm'].indexOf(k)>=0)val=money_(obj[k],k);
    else if(k==='JarakDasar'||k==='MaksJarakKm'){var km=Number(obj[k]);if(String(obj[k]).trim()===''||!Number.isInteger(km)||km<1||km>200)throw new Error(k+' harus bilangan bulat 1–200 km');val=km;}
    else if(COORD_KEYS_.indexOf(k)>=0){var raw=String(obj[k]==null?'':obj[k]).trim();
      if(raw==='')val='';else{var cn=Number(raw.replace(',','.')),lim=k==='TokoLat'?90:180;if(!isFinite(cn)||Math.abs(cn)>lim)throw new Error(k+' tidak valid (contoh: -7.0869)');val=String(cn);}}
    else val=obj[k];
    if(index[k]===undefined){index[k]=output.length;output.push([k,'']);}
    output[index[k]][1]=k==='PIN'?String(val).trim():COORD_KEYS_.indexOf(k)>=0?val:safeText_(val);
  });
  var lt=index.TokoLat!==undefined?String(output[index.TokoLat][1]).trim():'',lg=index.TokoLng!==undefined?String(output[index.TokoLng][1]).trim():'';
  if((lt===''||lg==='')&&(lt!==''||lg!==''))throw new Error('TokoLat dan TokoLng harus diisi keduanya, atau dikosongkan keduanya.');
  if(sh.getMaxRows()<output.length)sh.insertRowsAfter(sh.getMaxRows(),output.length-sh.getMaxRows());
  sh.getRange(1,2,output.length,1).setNumberFormat('@');sh.getRange(1,1,output.length,2).setValues(output);
  SETTINGS_MEMO_=null;SpreadsheetApp.flush();invalidateCatalog_();return {ok:true};
}

// Unggah gambar (banner/logo/QRIS/produk) ke Google Drive, kembalikan URL publik
function uploadImage(pin, filename, dataUrl) { auth_(pin);return uploadImage_(filename,dataUrl); }
function uploadImage_(filename,dataUrl) {
  var m = /^data:(image\/(?:png|jpeg|webp|gif));base64,(.+)$/.exec(String(dataUrl || ''));
  if (String(dataUrl || '').length > 7000000) throw new Error('Gambar maksimal 5 MB');
  if (!m) throw new Error('Format gambar tidak valid');
  var ext = m[1] === 'image/png' ? '.png' : m[1] === 'image/jpeg' ? '.jpg' : '.' + m[1].split('/')[1];
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1],
    String(filename || 'gambar').replace(/[^\w\-]+/g, '_') + '_' + Date.now() + ext);
  var it = DriveApp.getFoldersByName('POS Gambar');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('POS Gambar');
  var f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + f.getId();
}

/* ---------- setup ---------- */

// Jalankan SEKALI dari editor setelah menempel kode: memberi izin Sheets + Drive.
function authorizeImages_() {
  DriveApp.getRootFolder();
  ss_().getName();
}

// Membuat sheet yang belum ada (tidak menimpa data yang sudah ada)
function ensureSheets_() {
  var ss = ss_();
  var schemas = {
    Produk: ['ID','Kategori','Nama','Deskripsi','Harga','HargaCoret','Suka','Gambar','Aktif'],
    Pesanan: ['ID','Waktu','Nama','Telepon','Tipe','Alamat','Metode','Status','Subtotal','BiayaLayanan','Ongkir','Total','Dibayar','Kembalian','Catatan','ItemJSON'],
    Pengaturan: ['Kunci','Nilai']
  };
  Object.keys(schemas).forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) throw new Error('Sheet '+name+' belum ada. Jalankan setupPOS di editor.');
    var expected = schemas[name];
    var actual = sh.getRange(1,1,1,expected.length).getValues()[0];
    if (expected.some(function(k,i){return actual[i] !== k;})) throw new Error('Header sheet '+name+' tidak sesuai. Lihat PANDUAN; data tidak diubah.');
  });
}

// Alias kompatibilitas; setup tidak menghapus data.
function setupSheets() { setupPOS(); }

function buildSheets_(force) {
  var ss = ss_();
  function make(name, header, rows) {
    var sh = ss.getSheetByName(name);
    if (sh) return;
    sh = sh || ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#fde8e8');
    if (rows && rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.setFrozenRows(1);
    return sh;
  }
  make('Produk',
    ['ID', 'Kategori', 'Nama', 'Deskripsi', 'Harga', 'HargaCoret', 'Suka', 'Gambar', 'Aktif'],
    [
      ['P1', 'Signature Donat', 'Signature Donat Isi 6', '6 pcs Signature Donat Obun', 65000, 130000, 138, '', 'Ya'],
      ['P2', 'Signature Donat', 'Signature Donat Isi 2', '2 pcs Signature Donat Obun', 24000, 48000, 33, '', 'Ya'],
      ['P3', 'Signature Donat', 'Signature Donat Isi 1', '1 pc Signature Donat Obun', 13000, 26000, 4, '', 'Ya'],
      ['P4', 'Classic Donat', 'Classic Donat Isi 6', '6 pcs Classic Donat', 45000, 90000, 21, '', 'Ya'],
      ['P5', 'Classic Donat', 'Classic Donat Isi 3', '3 pcs Classic Donat', 24000, 48000, 9, '', 'Ya']
    ]);
  var pes = make('Pesanan',
    ['ID', 'Waktu', 'Nama', 'Telepon', 'Tipe', 'Alamat', 'Metode', 'Status', 'Subtotal',
     'BiayaLayanan', 'Ongkir', 'Total', 'Dibayar', 'Kembalian', 'Catatan', 'ItemJSON'], []);
  if (pes) pes.getRange('D:D').setNumberFormat('@');
  make('Pengaturan', ['Kunci', 'Nilai'], [
    ['NamaToko', 'Obun Bakeshop – Kalipancur'],
    ['LogoTeks', 'OBUN'],
    ['LogoURL', ''],
    ['StrukLogoURL', ''],
    ['BannerURL', ''],
    ['WarnaUtama', '#d32f2f'],
    ['Alamat', 'Jl. Abdul Rahman Saleh No.810, Kalipancur, Kec. Ngaliyan, Kota Semarang, Jawa Tengah 50183, Indonesia'],
    ['JamBuka', '09.00 - 21.00'],
    ['WhatsApp', ''],
    ['OutletLain', '1 Outlet Lain'],
    ['OutletLainInfo', 'Isi alamat outlet lain di sheet Pengaturan (kunci OutletLainInfo).'],
    ['Ongkir', 0],
    ['Bank', 'BCA'],
    ['NoRekening', ''],
    ['AtasNama', ''],
    ['QRISGambar', ''],
    ['PIN', '']
  ]);

}

/* ---------- Edit katalog dari aplikasi ---------- */
// Stok berubah setiap ada pesanan, jadi tidak ikut versi (supaya edit produk tidak sering bentrok).
function productVersion_(p) {
  var c = {};
  Object.keys(p).forEach(function (k) { if (k !== 'version' && k !== 'stok' && k !== 'habis') c[k] = p[k]; });
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(c)));
}
function getAdminProducts(pin,query){auth_(pin);return productPage_(catalog_(true),query);}
function setProductActive(pin,id,active,version){
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{auth_(pin);var p=productById_(id);
    if(!p)throw new Error('Produk tidak ditemukan');if(version!==p.version)throw new Error('Produk berubah. Muat ulang produk.');
    if(typeof active!=='boolean')throw new Error('Status tidak valid');
    var sh=ss_().getSheetByName('Produk'),match=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
    if(!match)throw new Error('Produk tidak ditemukan');sh.getRange(match.getRow(),9).setValue(active?'Ya':'Tidak');
    SpreadsheetApp.flush();invalidateCatalog_();p.aktif=active;delete p.version;p.version=productVersion_(p);return {ok:true,product:p};
  }finally{lock.releaseLock();}
}

// Aktifkan / nonaktifkan semua produk dalam satu brand ('' = produk tanpa brand).
function setBrandActive(pin,brand,active){
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{auth_(pin);
    if(typeof active!=='boolean')throw new Error('Status tidak valid');
    brand=String(brand==null?'':brand).trim();
    var sh=ss_().getSheetByName('Produk');if(!sh||sh.getLastRow()<2)throw new Error('Produk tidak ditemukan');
    var n=sh.getLastRow()-1,headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],bc=headers.indexOf('Brand');
    var brands=bc>=9?sh.getRange(2,bc+1,n,1).getValues():null;
    var ids=sh.getRange(2,1,n,1).getValues(),vals=sh.getRange(2,9,n,1).getValues(),changed=0,want=active?'Ya':'Tidak';
    for(var i=0;i<n;i++){
      if(ids[i][0]==='')continue;
      if(String(brands?brands[i][0]:'').trim()!==brand)continue;
      if(vals[i][0]!==want){vals[i][0]=want;changed++;}
    }
    if(changed){sh.getRange(2,9,n,1).setValues(vals);SpreadsheetApp.flush();invalidateCatalog_();}
    return {ok:true,changed:changed};
  }finally{lock.releaseLock();}
}

// Hapus permanen satu produk (baris Produk). Memakai versi untuk mencegah menghapus data yang sudah berubah.
function deleteProduct(pin,id,version,listQuery){
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{auth_(pin);var p=productById_(id);
    if(!p)throw new Error('Produk tidak ditemukan');if(version!==p.version)throw new Error('Produk berubah. Muat ulang produk.');
    var sh=ss_().getSheetByName('Produk'),match=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
    if(!match)throw new Error('Produk tidak ditemukan');
    sh.deleteRow(match.getRow());SpreadsheetApp.flush();invalidateCatalog_();
    return {ok:true,page:productPage_(catalog_(true),listQuery)};
  }finally{lock.releaseLock();}
}

function textField_(v, label, max, required) {
  var t = String(v == null ? '' : v).trim();
  if ((required && !t) || t.length > max) throw new Error(label + ' wajib sesuai batas ' + max + ' karakter');
  return t;
}
function saveProduct(pin,input){
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    auth_(pin);if(!input||typeof input!=='object')throw new Error('Produk tidak valid');
    var id=textField_(input.id,'ID',80,true);if(!/^[a-zA-Z0-9_-]+$/.test(id))throw new Error('ID tidak valid');
    var existing=productById_(id);
    if(existing&&existing.version!==input.version)throw new Error('Produk sudah berubah. Muat ulang produk lalu edit kembali.');
    if(!existing&&input.version)throw new Error('Produk sudah dihapus. Muat ulang produk.');
    var nama=textField_(input.nama,'Nama produk',150,true),brand=textField_(input.brand,'Brand',100,false),kategori=textField_(input.kategori,'Kategori',100,true),deskripsi=textField_(input.deskripsi,'Deskripsi',2000,false);
    var harga=money_(input.harga,'Harga'),coret=money_(input.hargaCoret,'Harga coret');
    var gambar=textField_(input.gambar,'URL gambar',2000,false);
    if(gambar&&!/^https:\/\/[^\s]+$/i.test(gambar))throw new Error('URL gambar harus HTTPS');
    if(input.imageData)gambar=uploadImage_('produk_'+id,input.imageData);
    var sh=ss_().getSheetByName('Produk'),brandCol=ensureBrandColumn_(),stokCol=ensureStockColumn_(),costCol=ensureCostColumn_();
    var match=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).matchCase(true).findNext():null;
    var row=match?match.getRow():sh.getLastRow()+1,width=Math.max(9,brandCol,stokCol,costCol);
    if(sh.getMaxRows()<row)sh.insertRowsAfter(sh.getMaxRows(),row-sh.getMaxRows());
    var range=sh.getRange(row,1,1,width),values=match?range.getValues()[0]:Array(width).fill(''),formulas=match?range.getFormulas()[0]:Array(width).fill('');
    values=values.map(function(v,i){return formulas[i]||(typeof v==='string'?safeText_(v):v);});
    var active=existing?existing.aktif:true;
    [id,kategori,nama,deskripsi,harga,coret,existing?existing.suka:0,gambar,active?'Ya':'Tidak'].forEach(function(v,i){values[i]=typeof v==='string'?safeText_(v):v;});values[brandCol-1]=safeText_(brand);
    // Stok hanya ditulis bila admin mengubahnya (stokDirty); jika tidak, nilai di sheet dipertahankan karena bisa berubah oleh pesanan baru.
    if(input.stokDirty||!match){
      var rawStok=input.stok==null?'':String(input.stok).trim();
      if(rawStok==='')values[stokCol-1]='';
      else{var sn=Number(rawStok);if(!Number.isInteger(sn)||sn<0||sn>1000000)throw new Error('Stok harus bilangan bulat 0–1.000.000, atau kosong untuk tidak dibatasi');values[stokCol-1]=sn;}
    }
    // Harga dasar hanya ditulis bila dikirim (editor selalu mengirimnya); panggilan lain tidak menghapusnya.
    if('hargaDasar' in input){var rawCost=input.hargaDasar==null?'':String(input.hargaDasar).trim();
      values[costCol-1]=rawCost===''?'':money_(Number(rawCost),'Harga dasar');}
    range.setValues([values]);SpreadsheetApp.flush();invalidateCatalog_();
    // Re-read the saved row so the returned version matches what is actually stored (safeText_ may alter text).
    return {ok:true,product:productById_(id),page:productPage_(catalog_(true),input.listQuery)};
  }finally{lock.releaseLock();}
}

/* ---------- Struk WhatsApp: snapshot pesanan yang tersimpan ---------- */
function normalizeWA_(v, optional) {
  var n = String(v == null ? '' : v).replace(/[^0-9]/g,'');
  if (!String(v == null ? '' : v).trim() && optional) return '';
  if (n.charAt(0) === '0') n = '62' + n.slice(1);
  else if (n.charAt(0) === '8') n = '62' + n;
  if (!/^[1-9][0-9]{7,14}$/.test(n)) throw new Error('Nomor WhatsApp tidak valid. Gunakan 08… atau kode negara 62…');
  return n;
}
function ensureReceiptColumns_(){
  var sh=ss_().getSheetByName('Pesanan');if(sh.getMaxColumns()<19)sh.insertColumnsAfter(sh.getMaxColumns(),19-sh.getMaxColumns());
  var expected=['StrukWhatsApp','NomorWhatsApp','JadwalPengantaran'],actual=sh.getRange(1,17,1,3).getValues()[0];
  expected.forEach(function(k,i){if(actual[i]&&actual[i]!==k)throw new Error('Kolom '+(17+i)+' Pesanan sudah dipakai. Pindahkan kolom tambahan tersebut.');});
  if(expected.some(function(k,i){return actual[i]!==k;}))sh.getRange(1,17,1,3).setValues([expected]).setFontWeight('bold');
}
function receiptFromRow_(row, settings) {
  var r = {ok:true,id:row[0],waktu:new Date(row[1]).toISOString(),nama:row[2],telepon:String(row[3]),tipe:row[4],alamat:row[5],metode:row[6],status:row[7],
    subtotal:Number(row[8]),biayaLayanan:Number(row[9]),ongkir:Number(row[10]),total:Number(row[11]),dibayar:Number(row[12]),kembalian:Number(row[13]),items:JSON.parse(row[15]).map(function(it){var c={};Object.keys(it).forEach(function(k){if(k!=='hpp')c[k]=it[k];});return c;})};
  r.jadwal = row[18] ? String(row[18]) : '';
  r.catatan = String(row[14] || '');
  // Batas bayar hanya untuk Transfer/QRIS yang masih menunggu; sisaDetik dihitung dengan jam server.
  if (r.status === 'Menunggu Pembayaran' && EXPIRING_METHODS_.indexOf(r.metode) >= 0) {
    var deadline = new Date(new Date(row[1]).getTime() + payTimeoutMin_(settings) * 60000);
    r.batasBayar = deadline.toISOString();
    r.batasLabel = Utilities.formatDate(deadline, 'Asia/Jakarta', 'HH:mm');
    r.sisaDetik = Math.max(0, Math.floor((deadline.getTime() - Date.now()) / 1000));
  }
  // Existing orders retain the destination and receipt captured at checkout.
  try { r.waNumber = normalizeWA_(row.length >= 18 ? row[17] : settings.WhatsApp, true); } catch(e) { r.waNumber = ''; }
  r.waText = row[16] || receiptText_(r, settings);
  return r;
}
function receiptText_(r, settings) {
  // toLocaleString ignores the locale on the Apps Script server, so format manually (Rp65.000).
  function rp(n){return 'Rp'+String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g,'.');}
  var lines = ['STRUK PESANAN — '+String(settings.NamaToko || 'Toko'),r.id,
    Utilities.formatDate(new Date(r.waktu),'Asia/Jakarta','dd/MM/yyyy HH:mm')+' WIB',
    'Nama: '+r.nama,'WhatsApp pembeli: +'+r.telepon,'Pengantaran: '+r.tipe];
  lines.push('Jadwal: '+(r.jadwal ? r.jadwal.replace('T',' ')+' WIB' : 'Secepatnya'));
  if (r.tipe === 'Dikirim') lines.push('Alamat: '+r.alamat);
  if (r.catatan && !r.items.some(function(it){return it.catatan;})) lines.push('Catatan: '+r.catatan);
  lines.push('Pembayaran: '+r.metode,'Status saat dibuat: '+r.status);
  if (r.batasLabel) lines.push('Bayar sebelum: '+r.batasLabel+' WIB (pesanan dibatalkan otomatis jika belum dibayar)');
  lines.push('','RINCIAN PESANAN');
  r.items.forEach(function(it,i){
    lines.push((i+1)+'. '+it.qty+' × '+it.nama+' — '+rp(it.total), '   Harga dasar / unit: '+rp(it.harga));
    (it.detail || []).forEach(function(d){lines.push('   '+d.grup+': '+d.nama.join(', ')+' (+'+rp(d.tambahan)+' / unit)');});
    if(it.catatan)lines.push('   Catatan: '+it.catatan);
    if(it.bawa)lines.push('   Dibawa pulang: Ya');
  });
  lines.push('', 'Subtotal: '+rp(r.subtotal),'Ongkir: '+rp(r.ongkir),'TOTAL: '+rp(r.total),'',
    'Mohon konfirmasi pesanan ini.','Struk pesanan digital, bukan bukti pembayaran terverifikasi.');
  return lines.join('\n');
}

// Append Brand after existing columns; do not replace user-defined columns.
function ensureBrandColumn_() {
  var sh = ss_().getSheetByName('Produk');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf('Brand') + 1;
  if (col) return col;
  col = Math.max(9, headers.length) + 1;
  if (sh.getMaxColumns() < col) sh.insertColumnsAfter(sh.getMaxColumns(), col-sh.getMaxColumns());
  sh.getRange(1,col).setValue('Brand').setFontWeight('bold');
  return col;
}

// Kolom Stok ditambahkan setelah kolom yang sudah ada. Kosong = tidak dibatasi, angka = jumlah tersisa.
function ensureStockColumn_() {
  var sh = ss_().getSheetByName('Produk');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf('Stok') + 1;
  if (col) return col;
  col = Math.max(9, headers.length) + 1;
  if (sh.getMaxColumns() < col) sh.insertColumnsAfter(sh.getMaxColumns(), col-sh.getMaxColumns());
  sh.getRange(1,col).setValue('Stok').setFontWeight('bold');
  return col;
}

// Kolom HargaDasar (modal per unit) ditambahkan setelah kolom yang ada. Kosong = belum diisi.
function ensureCostColumn_() {
  var sh = ss_().getSheetByName('Produk');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf('HargaDasar') + 1;
  if (col) return col;
  col = Math.max(9, headers.length) + 1;
  if (sh.getMaxColumns() < col) sh.insertColumnsAfter(sh.getMaxColumns(), col-sh.getMaxColumns());
  sh.getRange(1,col).setValue('HargaDasar').setFontWeight('bold');
  return col;
}

/* ---------- Pelanggan (nama, telepon, alamat) ---------- */
var CUSTOMER_HEADERS_ = ['Telepon','Nama','Alamat','Lat','Lng','Dibuat','Terakhir'];
function ensureCustomerSheet_() {
  var ss = ss_(), sh = ss.getSheetByName('Pelanggan');
  if (!sh) {
    sh = ss.insertSheet('Pelanggan');
    sh.getRange(1,1,1,CUSTOMER_HEADERS_.length).setValues([CUSTOMER_HEADERS_]).setFontWeight('bold').setBackground('#fde8e8');
    sh.setFrozenRows(1);
    sh.getRange(1,1,sh.getMaxRows(),1).setNumberFormat('@');
  }
  return sh;
}
function numOrBlank_(v) { return (v === '' || v == null || !isFinite(Number(v))) ? '' : Number(v); }
function upsertCustomer_(c, now) {
  var phone = String(c.telepon || '');
  if (!/^62\d{8,13}$/.test(phone)) return;
  var sh = ensureCustomerSheet_();
  var m = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(phone).matchEntireCell(true).matchCase(true).findNext() : null;
  var lat = numOrBlank_(c.lat), lng = numOrBlank_(c.lng);
  if (m) {
    var row = m.getRow(), cur = sh.getRange(row,1,1,7).getValues()[0];
    cur[1] = safeText_(c.nama);
    if (c.alamat) { cur[2] = safeText_(c.alamat); if (lat !== '' && lng !== '') { cur[3] = lat; cur[4] = lng; } }
    cur[6] = now;
    sh.getRange(row,1,1,7).setValues([cur]);
  } else {
    sh.appendRow([phone, safeText_(c.nama), safeText_(c.alamat || ''), lat, lng, now, now]);
  }
}
// Mengisi sheet Pelanggan dari pesanan lama (otomatis saat tab Pelanggan dibuka pertama kali; bisa juga dijalankan dari editor).
function backfillCustomers() {
  var sh = ensureCustomerSheet_(), ps = ss_().getSheetByName('Pesanan');
  if (!ps || ps.getLastRow() < 2 || sh.getLastRow() > 1) return 0;
  var seen = {};
  ps.getRange(2,1,ps.getLastRow()-1,6).getValues().forEach(function (r) {
    var ph = String(r[3] || '');
    if (!/^62\d{8,13}$/.test(ph)) return;
    var addr = String(r[5] || ''), mm = /maps\?q=(-?[\d.]+),(-?[\d.]+)/.exec(addr), old = seen[ph] || {};
    addr = addr.split('\nTitik peta:')[0];
    seen[ph] = { nama: r[2], alamat: (r[4] === 'Dikirim' && addr) ? addr : (old.alamat || ''),
      lat: (r[4] === 'Dikirim' && mm) ? mm[1] : (old.lat || ''), lng: (r[4] === 'Dikirim' && mm) ? mm[2] : (old.lng || ''),
      dibuat: old.dibuat || r[1], terakhir: r[1] };
  });
  var rows = Object.keys(seen).map(function (ph) {
    var c = seen[ph];
    return [ph, safeText_(c.nama), safeText_(c.alamat), numOrBlank_(c.lat), numOrBlank_(c.lng), c.dibuat, c.terakhir];
  });
  if (rows.length) sh.getRange(2,1,rows.length,7).setValues(rows);
  return rows.length;
}
// Admin: daftar pelanggan + statistik belanja (dihitung dari pesanan yang sudah dibayar).
function listCustomers(pin) {
  auth_(pin);
  var sh = ensureCustomerSheet_();
  if (sh.getLastRow() < 2) backfillCustomers();
  var cust = sh.getLastRow() < 2 ? [] : sh.getRange(2,1,sh.getLastRow()-1,7).getValues();
  var stats = {}, ps = ss_().getSheetByName('Pesanan');
  if (ps && ps.getLastRow() > 1) {
    ps.getRange(2,1,ps.getLastRow()-1,12).getValues().forEach(function (r) {
      var ph = String(r[3] || ''), st = stats[ph] || (stats[ph] = { pesanan: 0, dibayar: 0, belanja: 0, terakhir: 0 });
      st.pesanan++;
      var t = r[1] instanceof Date ? r[1].getTime() : 0;
      if (t > st.terakhir) st.terakhir = t;
      if (PAID_STATUSES_.indexOf(r[7]) >= 0) { st.dibayar++; st.belanja += Number(r[11]) || 0; }
    });
  }
  var out = cust.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var ph = String(r[0]), st = stats[ph] || { pesanan: 0, dibayar: 0, belanja: 0, terakhir: 0 };
    var last = st.terakhir || (r[6] instanceof Date ? r[6].getTime() : 0);
    return { telepon: ph, nama: String(r[1] || ''), alamat: String(r[2] || ''), lat: numOrBlank_(r[3]), lng: numOrBlank_(r[4]),
      pesanan: st.pesanan, dibayar: st.dibayar, belanja: st.belanja, terakhir: last ? new Date(last).toISOString() : '' };
  });
  out.sort(function (a, b) { return (b.terakhir || '').localeCompare(a.terakhir || ''); });
  return { customers: out.slice(0, 1000), total: out.length };
}

/* ---------- Laporan transaksi (jam / hari / bulan / tahun) ---------- */
// Pendapatan = Total pesanan (termasuk ongkir). Keuntungan = jumlah (harga jual - harga dasar) x qty; ongkir tidak dihitung laba.
// Hanya pesanan yang sudah dibayar (Lunas, Diproses, Siap, Selesai). Harga dasar diambil dari snapshot saat pesanan dibuat;
// pesanan lama tanpa snapshot memakai harga dasar produk sekarang. Item tanpa harga dasar dihitung laba penuh dan dilaporkan di 'unknown'.
function getReport(pin, level, opts) {
  auth_(pin);
  if (['jam','hari','bulan','tahun'].indexOf(level) < 0) throw new Error('Jenis laporan tidak valid');
  opts = opts || {};
  var nowY = Number(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy'));
  var year = Math.floor(Number(opts.year)) || nowY, month = Math.floor(Number(opts.month)) || 1, date = String(opts.date || '');
  if (year < 2000 || year > 2100) throw new Error('Tahun tidak valid');
  if (month < 1 || month > 12) throw new Error('Bulan tidak valid');
  if (level === 'jam' && !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid');
  var cost = {};
  catalog_(true).forEach(function (p) { if (p.hargaDasar !== null && p.hargaDasar !== undefined) cost[p.id] = p.hargaDasar; });
  var sh = ss_().getSheetByName('Pesanan'), list = [], unknown = 0;
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2,1,sh.getLastRow()-1,16).getValues().forEach(function (r) {
      if (PAID_STATUSES_.indexOf(r[7]) < 0 || !(r[1] instanceof Date)) return;
      var ts = Utilities.formatDate(r[1], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      var y = Number(ts.slice(0,4)), m = Number(ts.slice(5,7)), d = ts.slice(0,10), h = Number(ts.slice(11,13));
      if (level === 'bulan' && y !== year) return;
      if (level === 'hari' && !(y === year && m === month)) return;
      if (level === 'jam' && d !== date) return;
      var items = []; try { items = JSON.parse(r[15] || '[]'); } catch (e) {}
      var untung = 0;
      items.forEach(function (it) {
        var qty = Number(it.qty) || 0, c = (it.hpp != null && isFinite(Number(it.hpp))) ? Number(it.hpp) : (cost[it.id] !== undefined ? cost[it.id] : null);
        if (c === null) { unknown++; c = 0; }
        untung += (Number(it.total) || (Number(it.harga) || 0) * qty) - c * qty;
      });
      list.push({ id: String(r[0]), nama: String(r[2] || ''), metode: String(r[6] || ''), tipe: String(r[4] || ''), ts: ts, y: y, m: m, d: d, h: h,
        pendapatan: Number(r[11]) || 0, untung: untung });
    });
  }
  var rows = [], idx = {};
  function bucket(key) { var b = { key: key, count: 0, pendapatan: 0, untung: 0 }; idx[key] = b; rows.push(b); }
  if (level === 'tahun') {
    var lo = nowY, hi = nowY;
    list.forEach(function (o) { if (o.y < lo) lo = o.y; if (o.y > hi) hi = o.y; });
    for (var yy = lo; yy <= hi; yy++) bucket(yy);
  } else if (level === 'bulan') { for (var mm = 1; mm <= 12; mm++) bucket(mm); }
  else if (level === 'hari') {
    var days = new Date(Date.UTC(year, month, 0)).getUTCDate(), pre = year + '-' + ('0' + month).slice(-2) + '-';
    for (var dd = 1; dd <= days; dd++) bucket(pre + ('0' + dd).slice(-2));
  } else { for (var hh = 0; hh < 24; hh++) bucket(hh); }
  var totals = { count: 0, pendapatan: 0, untung: 0 }, trx = [];
  list.forEach(function (o) {
    var b = idx[level === 'tahun' ? o.y : level === 'bulan' ? o.m : level === 'hari' ? o.d : o.h];
    if (b) { b.count++; b.pendapatan += o.pendapatan; b.untung += o.untung; }
    totals.count++; totals.pendapatan += o.pendapatan; totals.untung += o.untung;
    if (level === 'jam') trx.push({ id: o.id, nama: o.nama, jam: o.ts.slice(11), metode: o.metode, tipe: o.tipe, pendapatan: o.pendapatan, untung: o.untung });
  });
  trx.sort(function (a, b) { return b.jam.localeCompare(a.jam); });
  return { level: level, year: year, month: month, date: date, rows: rows, trx: trx, totals: totals, unknown: unknown };
}

/* ---------- Jam buka (Pengaturan → JamBuka, mis. "09.00 - 21.00", WIB) ---------- */
function parseHours_(str) {
  var m = /(\d{1,2})[.:](\d{2})\s*[-–—]\s*(\d{1,2})[.:](\d{2})/.exec(String(str || ''));
  if (!m) return null; // format tak dikenal / "24 jam": tidak dibatasi
  var o = Number(m[1]) * 60 + Number(m[2]), c = Number(m[3]) * 60 + Number(m[4]);
  if (Number(m[2]) > 59 || Number(m[4]) > 59 || o > 1439 || c > 1440 || o === c) return null;
  return { open: o, close: c };
}
function minutesJakarta_(date) {
  var p = Utilities.formatDate(date, 'Asia/Jakarta', 'HH:mm').split(':');
  return Number(p[0]) * 60 + Number(p[1]);
}
function isOpenMin_(h, t) { return h.close > h.open ? (t >= h.open && t < h.close) : (t >= h.open || t < h.close); }
function hoursLabel_(h) {
  function f(x) { x = x % 1440; return ('0' + Math.floor(x / 60)).slice(-2) + '.' + ('0' + (x % 60)).slice(-2); }
  return f(h.open) + ' - ' + f(h.close);
}
// Opsional: aktif hanya bila Pengaturan berisi baris CekJamBuka = Ya.
function assertOpen_(date, scheduled) {
  var st = settings_();
  if (String(st.CekJamBuka || '').trim().toLowerCase() !== 'ya') return;
  var h = parseHours_(st.JamBuka);
  if (!h || isOpenMin_(h, minutesJakarta_(date))) return;
  throw new Error(scheduled
    ? 'Toko tutup pada jam tersebut. Jam buka ' + hoursLabel_(h) + ' WIB.'
    : 'Toko sedang tutup. Jam buka ' + hoursLabel_(h) + ' WIB. Pilih jadwal di jam buka.');
}

/* Scheduled fulfillment is stored as YYYY-MM-DDTHH:mm, interpreted in WIB. */
function validateSchedule_(value) {
  if(value == null || value === ''){assertOpen_(new Date(), false);return '';}
  if(typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(value))throw new Error('Tanggal / jam pengiriman tidak valid');
  var date=new Date(value+':00+07:00');
  if(!isFinite(date.getTime()) || Utilities.formatDate(date,'Asia/Jakarta',"yyyy-MM-dd'T'HH:mm")!==value)throw new Error('Tanggal / jam pengiriman tidak valid');
  if(date.getTime()<=Date.now())throw new Error('Waktu pengiriman sudah lewat. Pilih ulang waktu (WIB).');
  assertOpen_(date, true);
  return value;
}
