const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { DatabaseSync } = require('node:sqlite');
const XLSX = require('xlsx-js-style');

const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'admin123';
const DATA_DIR = path.join(__dirname, 'data');
const DB_PATH = path.join(DATA_DIR, 'sumbangan.db');
const UPLOAD_DIR = path.join(__dirname, 'public', 'uploads');

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// ---------- Log: redam duplikat & batasi ukuran ----------
const LOG_MAX_BYTES = 2 * 1024 * 1024;
const LOG_DEDUPE_WINDOW = 60 * 1000;
const logDedupe = new Map();

function batasiLogFile(file) {
  try {
    if (!fs.existsSync(file)) return;
    const st = fs.statSync(file);
    if (st.size > LOG_MAX_BYTES) fs.truncateSync(file, 0);
  } catch (e) {}
}

function pasangFilterLog() {
  const asliLog = console.log.bind(console);
  const asliErr = console.error.bind(console);
  const filter = (asli, pesan) => (...args) => {
    try {
      const teks = args
        .map(a => (a instanceof Error ? a.message : typeof a === 'string' ? a : JSON.stringify(a)))
        .join(' ');
      const now = Date.now();
      const kunci = pesan + '|' + teks.slice(0, 120);
      const last = logDedupe.get(kunci);
      if (last && now - last < LOG_DEDUPE_WINDOW) {
        logDedupe.set(kunci, now);
        return;
      }
      logDedupe.set(kunci, now);
      if (logDedupe.size > 500) {
        for (const [k, t] of logDedupe) if (now - t > LOG_DEDUPE_WINDOW) logDedupe.delete(k);
      }
      asli(pesan, ...args);
    } catch (e) {}
  };
  console.log = filter(asliLog, '[app]');
  console.error = filter(asliErr, '[app]');
  for (const f of ['server.log', 'server-error.log']) batasiLogFile(path.join(__dirname, f));
  batasiLogFile(path.join(DATA_DIR, 'server.log'));
  batasiLogFile(path.join(DATA_DIR, 'server-err.log'));
}
pasangFilterLog();

process.on('unhandledRejection', (alasan) => {
  const pesan = alasan && alasan.message ? alasan.message : String(alasan);
  console.error('[app] unhandledRejection:', pesan);
});
process.on('uncaughtException', (err) => {
  console.error('[app] uncaughtException:', err && err.message ? err.message : String(err));
});

function initDb() {
  const d = new DatabaseSync(DB_PATH);
  d.exec('PRAGMA journal_mode = WAL');
  d.exec('PRAGMA synchronous = NORMAL');
  d.exec('PRAGMA busy_timeout = 5000');
  d.exec('PRAGMA foreign_keys = ON');
  d.exec(`
    CREATE TABLE IF NOT EXISTS setting (
      key TEXT PRIMARY KEY,
      value TEXT
    );
    CREATE TABLE IF NOT EXISTS jenis_sumbangan (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nama TEXT NOT NULL,
      target INTEGER DEFAULT 0,
      aktif INTEGER DEFAULT 1,
      kelas TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS siswa (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nis TEXT,
      nama TEXT NOT NULL,
      kelas TEXT,
      jenis_kelamin TEXT,
      no_hp TEXT
    );
    CREATE TABLE IF NOT EXISTS pembayaran (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      siswa_id INTEGER,
      jenis_id INTEGER NOT NULL,
      nominal INTEGER NOT NULL,
      tanggal TEXT NOT NULL,
      periode TEXT,
      keterangan TEXT,
      nama_pihak TEXT,
      created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS penggunaan_dana (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tanggal TEXT NOT NULL,
      keterangan TEXT NOT NULL,
      nominal INTEGER NOT NULL,
      dokumen TEXT,
      created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS pengingat (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      siswa_id INTEGER,
      bulan TEXT NOT NULL,
      periode TEXT,
      pesan TEXT,
      status TEXT NOT NULL,
      hasil TEXT,
      created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_pembayaran_tanggal ON pembayaran(tanggal);
    CREATE INDEX IF NOT EXISTS idx_pembayaran_siswa ON pembayaran(siswa_id);
    CREATE INDEX IF NOT EXISTS idx_pembayaran_jenis ON pembayaran(jenis_id);
    CREATE INDEX IF NOT EXISTS idx_pembayaran_tanggal_jenis ON pembayaran(tanggal, jenis_id);
    CREATE INDEX IF NOT EXISTS idx_siswa_kelas ON siswa(kelas);
    CREATE INDEX IF NOT EXISTS idx_penggunaan_tanggal ON penggunaan_dana(tanggal);
    CREATE INDEX IF NOT EXISTS idx_pengingat_bulan ON pengingat(bulan);
  `);
  try { d.exec('ALTER TABLE pembayaran ADD COLUMN periode TEXT'); } catch (e) {}
  try { d.exec('ALTER TABLE siswa ADD COLUMN jenis_kelamin TEXT'); } catch (e) {}
  try { d.exec('ALTER TABLE siswa ADD COLUMN no_hp TEXT'); } catch (e) {}
  try { d.exec('ALTER TABLE penggunaan_dana ADD COLUMN dokumen TEXT'); } catch (e) {}
  try { d.exec('ALTER TABLE jenis_sumbangan ADD COLUMN kelas TEXT DEFAULT \'\''); } catch (e) {}
  try { d.exec('ALTER TABLE pembayaran ADD COLUMN nama_pihak TEXT'); } catch (e) {}
  migrasiPembayaranDll(d);
  return d;
}

function migrasiPembayaranDll(d) {
  const cols = d.prepare("PRAGMA table_info(pembayaran)").all().map(c => c.name);
  if (cols.includes('nama_pihak') && !cols.includes('siswa_id')) return;
  const hasNamaPihak = cols.includes('nama_pihak');
  const siswaNullable = d.prepare("PRAGMA table_info(pembayaran)").all().find(c => c.name === 'siswa_id' && !c.notnull);
  if (hasNamaPihak && siswaNullable) return;
  d.exec('BEGIN');
  try {
    d.exec(`
      CREATE TABLE pembayaran_baru (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        siswa_id INTEGER,
        jenis_id INTEGER NOT NULL,
        nominal INTEGER NOT NULL,
        tanggal TEXT NOT NULL,
        periode TEXT,
        keterangan TEXT,
        nama_pihak TEXT,
        created_at TEXT NOT NULL
      );
    `);
    d.exec(`
      INSERT INTO pembayaran_baru (id, siswa_id, jenis_id, nominal, tanggal, periode, keterangan, nama_pihak, created_at)
      SELECT id, siswa_id, jenis_id, nominal, tanggal, periode, keterangan, NULL, created_at FROM pembayaran
    `);
    d.exec('DROP TABLE pembayaran');
    d.exec('ALTER TABLE pembayaran_baru RENAME TO pembayaran');
    d.exec('COMMIT');
  } catch (e) {
    d.exec('ROLLBACK');
  }
}

let db = initDb();

function seed() {
  const st = db.prepare('SELECT COUNT(*) AS c FROM setting');
  const { c } = st.get();
  if (c === 0) {
    const ins = db.prepare('INSERT INTO setting (key, value) VALUES (?, ?)');
    ins.run('nama_sekolah', 'SMP Negeri 6 Kebumen');
    ins.run('tahun_ajaran', '2026/2027');
    ins.run('target_tahunan', '10000000');

    const jIns = db.prepare('INSERT INTO jenis_sumbangan (nama, target) VALUES (?, ?)');
    jIns.run('Sumbangan Sukarela', 0);
  }
}
seed();

const settingGet = (key) => {
  const row = db.prepare('SELECT value FROM setting WHERE key = ?').get(key);
  return row ? row.value : null;
};
const settingSet = (key, value) => {
  db.prepare('INSERT INTO setting (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value').run(key, value);
};

function valKelasJenis(v) {
  const s = String(v == null ? '' : v).trim();
  return ['7', '8', '9'].includes(s) ? s : '';
}

if (!settingGet('admin_nama')) settingSet('admin_nama', 'Admin');

// ---------- WhatsApp notification ----------
const WA_LIB = (() => {
  try { return require('@whiskeysockets/baileys'); } catch (e) { return null; }
})();
const QR_LIB = (() => {
  try { return require('qrcode'); } catch (e) { return null; }
})();

const WA_SESSION_DIR = path.join(DATA_DIR, 'wa-session');
let waClient = null;
let waState = { status: 'off', qr: null, ready: false, nomor: null, error: null };
let waStarting = false;

function sanitizeWaNumber(input) {
  let n = String(input || '').replace(/[^\d]/g, '');
  if (!n) return null;
  if (n.startsWith('0')) n = '62' + n.slice(1);
  else if (n.startsWith('8')) n = '62' + n;
  if (!n.startsWith('62')) n = '62' + n;
  return n + '@s.whatsapp.net';
}

function waParseTarget(input) {
  const s = String(input || '').trim();
  if (!s) return null;
  if (s.includes('@')) {
    const jid = s.replace(/[\s]/g, '');
    if (jid.endsWith('@g.us') || jid.endsWith('@s.whatsapp.net')) return jid;
    const angka = jid.split('@')[0].replace(/[^\d]/g, '');
    if (angka && (jid.endsWith('@c.us') || jid.endsWith('@lid'))) return angka + '@s.whatsapp.net';
    return null;
  }
  return sanitizeWaNumber(s);
}

function waAdminList() {
  const list = new Set();
  const nomorBot = String(waState.nomor || '').replace(/[^\d]/g, '');
  if (nomorBot) list.add(nomorBot);
  const raw = String(settingGet('wa_admin') || '');
  for (const item of raw.split(/[,\s;]+/)) {
    const n = item.replace(/[^\d]/g, '');
    if (n.length >= 8) list.add(n);
  }
  return list;
}

function waIsAdmin(jid) {
  const n = String(jid || '').split('@')[0].replace(/[^\d]/g, '');
  if (!n) return false;
  return waAdminList().has(n);
}

function waTextPesan(msg) {
  const m = (msg && msg.message) || {};
  const sumber = [m.conversation, m.extendedTextMessage, m.imageMessage, m.videoMessage, m.documentMessage, m.buttonsResponseMessage];
  for (const s of sumber) {
    const t = s && (s.text || s.caption || (s.selectedButtonId || ''));
    if (typeof t === 'string' && t.trim()) return t.trim();
  }
  return '';
}

const WA_BANTUAN = [
  'Perintah: /forward <nomor> <pesan>',
  'Contoh: /forward 081234567890 Assalamu\'alaikum, info sumbani bulan ini.',
  'Tujuan bisa nomor HP atau ID grup. Hanya nomor admin yang bisa memakai.'
].join('\n');

async function waStart() {
  if (!WA_LIB || waClient || waStarting) return;
  waStarting = true;
  try {
    const {
      default: makeWASocket,
      useMultiFileAuthState,
      DisconnectReason,
      fetchLatestBaileysVersion
    } = WA_LIB;
    const { state, saveCreds } = await useMultiFileAuthState(WA_SESSION_DIR);
    let logger;
    try {
      const pino = require('pino');
      logger = pino({ level: 'silent' });
    } catch (e) { logger = undefined; }
    let versi;
    try {
      const v = await fetchLatestBaileysVersion();
      versi = v.version;
    } catch (e) {
      versi = [2, 3000, 1017435734];
    }
    const sock = makeWASocket({
      version: versi,
      auth: state,
      logger,
      printQRInTerminal: false,
      browser: ['GotongRoyong', 'Chrome', '1.0.0'],
      markOnlineOnConnect: false,
      syncFullHistory: false
    });
    waClient = sock;
    waState = { status: 'connecting', qr: null, ready: false, nomor: null, error: null };
    sock.ev.on('creds.update', saveCreds);
    sock.ev.on('messages.upsert', async ({ messages }) => {
      for (const msg of messages || []) {
        try {
          if (!msg || msg.key && msg.key.fromMe) continue;
          const teks = waTextPesan(msg);
          if (!teks) continue;
          const perintah = teks.match(/^\/?(forward|forwarding)\b/i);
          if (!perintah) continue;
          if (!waIsAdmin(msg.key && msg.key.remoteJid)) {
            console.log('WhatsApp: perintah forward dari non-admin (' + ((msg.key && msg.key.remoteJid) || '?') + ') diabaikan.');
            continue;
          }
          const sisa = teks.slice(perintah[0].length).trim();
          const pisah = sisa.search(/\s/);
          const target = pisah === -1 ? sisa : sisa.slice(0, pisah);
          const isi = pisah === -1 ? '' : sisa.slice(pisah + 1).trim();
          if (!target || !isi) {
            await sock.sendMessage(msg.key.remoteJid, { text: WA_BANTUAN });
            continue;
          }
          const jid = waParseTarget(target);
          if (!jid) {
            await sock.sendMessage(msg.key.remoteJid, { text: 'Nomor tujuan tidak valid. Contoh: /forward 081234567890 halo' });
            continue;
          }
          try {
            await sock.sendMessage(jid, { text: isi });
            console.log('WhatsApp: pesan diteruskan dari ' + msg.key.remoteJid + ' ke ' + jid);
            await sock.sendMessage(msg.key.remoteJid, { text: 'âœ… Terkirim ke ' + target + ':\n' + isi });
          } catch (err) {
            await sock.sendMessage(msg.key.remoteJid, { text: 'âŒ Gagal meneruskan: ' + (err && err.message ? err.message : 'tidak diketahui') });
          }
        } catch (e) {
          console.error('WhatsApp: gagal proses perintah:', e.message);
        }
      }
    });
    sock.ev.on('connection.update', (update) => {
      const { connection, lastDisconnect, qr } = update;
      if (qr) {
        waState.qr = qr;
        waState.status = 'qr';
        waState.ready = false;
        waState.error = null;
      }
      if (connection === 'open') {
        const id = (sock.user && sock.user.id) ? String(sock.user.id).split(':')[0] : '';
        waState.status = 'ready';
        waState.ready = true;
        waState.qr = null;
        waState.nomor = id;
        waState.error = null;
        console.log('WhatsApp terhubung sebagai ' + (waState.nomor || '?'));
      } else if (connection === 'close') {
        const kode = (lastDisconnect && lastDisconnect.error && lastDisconnect.error.output && lastDisconnect.error.output.statusCode) || 0;
        const perluRescan = kode === DisconnectReason.loggedOut || kode === DisconnectReason.badSession;
        waState.ready = false;
        if (waClient === sock) waClient = null;
        try { sock.end(); } catch (e) {}
        waStarting = false;
        if (perluRescan) {
          waState.status = 'disconnected';
          waState.error = 'Sesi WhatsApp berakhir. Pindai ulang QR.';
        } else if (settingGet('wa_enabled') === '1') {
          waState.status = 'connecting';
          waState.error = null;
          setTimeout(() => {
            if (!waClient && !waStarting && settingGet('wa_enabled') === '1') {
              console.log('WhatsApp: menyambungkan ulang...');
              waStart();
            }
          }, 5000);
        }
      }
    });
  } catch (e) {
    console.error('WhatsApp init gagal:', e.message);
    waState.status = 'error';
    waState.error = e.message;
    waClient = null;
    waStarting = false;
  }
}

function waStop() {
  if (waClient) { try { waClient.end(); } catch (e) {} waClient = null; }
  waStarting = false;
  waState = { status: 'off', qr: null, ready: false, nomor: null, error: null };
}

async function kirimWa(nomor, pesan) {
  if (!WA_LIB) return { ok: false, error: 'Library WhatsApp tidak tersedia.' };
  const to = sanitizeWaNumber(nomor);
  if (!to) return { ok: false, error: 'Nomor WhatsApp tidak valid.' };
  if (!waClient || waState.status !== 'ready') return { ok: false, error: 'WhatsApp belum terhubung. Hubungkan lewat menu Pengaturan.' };
  try {
    await waClient.sendMessage(to, { text: pesan });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function buatPesanPembayaran({ nama, kelas, jenis, nominal, tanggal }) {
  const sekolah = settingGet('nama_sekolah') || 'Sekolah';
  const adminNama = settingGet('admin_nama') || 'Admin';
  const total = db.prepare('SELECT IFNULL(SUM(nominal), 0) AS t FROM pembayaran').get().t;
  return [
    `Assalamu'alaikum Bapak/Ibu wali dari *${nama}* (${kelas || '-'}),`,
    '',
    `Pembayaran sumbangan di *${sekolah}* telah kami terima:`,
    `ðŸ“‹ Jenis   : *${jenis}*`,
    `ðŸ’° Nominal : *Rp ${Number(nominal).toLocaleString('id-ID')}*`,
    `ðŸ“… Tanggal : ${tanggal}`,
    '',
    `Total sumbangan terkumpul: Rp ${Number(total).toLocaleString('id-ID')}`,
    '',
    'Terima kasih atas partisipasi Bapak/Ibu.',
    `â€” ${adminNama} | ${sekolah}`
  ].join('\n');
}

if (settingGet('wa_enabled') === '1') waStart();

const fmtDate = (d) => d.toISOString().slice(0, 10);

// ---------- Admin auth ----------
const hashPass = (s) => crypto.createHash('sha256').update(String(s)).digest('hex');
function passwordOk(token) {
  const stored = settingGet('admin_password_hash');
  if (stored) return hashPass(token) === stored;
  return hashPass(token) === hashPass(ADMIN_PASSWORD);
}

// ---------- SSE realtime ----------
let sseClients = new Set();

function broadcast(data) {
  const payload = `data: ${JSON.stringify(data)}\n\n`;
  for (const res of sseClients) {
    try { res.write(payload); } catch (e) { sseClients.delete(res); }
  }
}
function notifyChange() {
  broadcast({ type: 'update', time: Date.now() });
}

// ---------- Express ----------
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.use((req, res, next) => {
  const isWrite = (req.method === 'POST' || req.method === 'DELETE') && req.path !== '/api/admin/login';
  const isLaporan = req.path.startsWith('/api/laporan/');
  const isWa = req.path.startsWith('/api/wa/');
  const isKwitansi = req.path.startsWith('/api/kwitansi');
  const isSensitiveGet = req.method === 'GET' &&
    (['/api/siswa', '/api/settings', '/api/jenis', '/api/backup/download'].includes(req.path) || isWa || isKwitansi);
  if (isWrite || isLaporan || isSensitiveGet) {
    const auth = req.headers.authorization || '';
    const token = auth.replace(/^Bearer\s+/i, '');
    if (!passwordOk(token)) {
      return res.status(401).json({ error: 'Tidak diizinkan. Periksa password admin.' });
    }
  }
  next();
});

// --- Events (SSE) ---
app.get('/api/events', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();
  res.write('retry: 3000\n\n');
  sseClients.add(res);
  req.on('close', () => sseClients.delete(res));
});

// --- Public stats (real-time view) ---
app.get('/api/stats', (req, res) => {
  const namaSekolah = settingGet('nama_sekolah');
  const tahunAjaran = settingGet('tahun_ajaran');
  const targetTahunan = Number(settingGet('target_tahunan')) || 0;

  const jenis = db.prepare(`
    SELECT j.id, j.nama, j.target,
           IFNULL(SUM(p.nominal), 0) AS terkumpul,
           COUNT(p.id) AS jumlah_transaksi
    FROM jenis_sumbangan j
    LEFT JOIN pembayaran p ON p.jenis_id = j.id
    WHERE j.aktif = 1
    GROUP BY j.id
    ORDER BY j.id
  `).all();

  const total = jenis.reduce((s, j) => s + j.terkumpul, 0);
  const siswaPembayar = db.prepare('SELECT COUNT(DISTINCT siswa_id) AS c FROM pembayaran').get().c;
  const totalSiswa = db.prepare('SELECT COUNT(*) AS c FROM siswa').get().c;
  const totalTransaksi = db.prepare('SELECT COUNT(*) AS c FROM pembayaran').get().c;
  const terakhirUpdate = db.prepare('SELECT MAX(created_at) AS t FROM pembayaran').get().t;
  const totalPenggunaan = db.prepare('SELECT IFNULL(SUM(nominal), 0) AS t FROM penggunaan_dana').get().t;

  res.json({
    namaSekolah,
    tahunAjaran,
    targetTahunan,
    logo: settingGet('logo'),
    total,
    totalPenggunaan,
    saldo: total - totalPenggunaan,
    sisaTarget: targetTahunan - total,
    progress: targetTahunan > 0 ? Math.min(100, Math.round((total / targetTahunan) * 100)) : 0,
    jenis,
    siswaPembayar,
    totalSiswa,
    totalTransaksi,
    terakhirUpdate,
    time: Date.now()
  });
});

// --- QRIS pembayaran ---
async function qrisPayload() {
  const aktif = settingGet('qris_enabled') === '1';
  const noVa = settingGet('qris_no_va') || '';
  const data = {
    aktif: aktif,
    nama: settingGet('qris_nama') || '',
    atas_nama: settingGet('qris_atas_nama') || '',
    no_va: noVa
  };
  if (aktif && noVa && QR_LIB) {
    try {
      data.qr = await QR_LIB.toDataURL(noVa, { margin: 1, width: 600 });
    } catch (_) {}
  }
  return data;
}

app.get('/api/qris', async (req, res) => {
  res.json(await qrisPayload());
});

app.get('/api/history', (req, res) => {
  const limit = Math.min(Math.max(Number(req.query.limit) || 50, 1), 500);
  const offset = Math.max(Number(req.query.offset) || 0, 0);
  const { where, args } = buildFilters(req);
  const sqlWhere = where || '';
  const total = db.prepare(`
    SELECT COUNT(*) AS c
    FROM pembayaran p
    LEFT JOIN siswa s ON s.id = p.siswa_id
    LEFT JOIN jenis_sumbangan j ON j.id = p.jenis_id
    ${sqlWhere}
  `).get(...args).c;
  const rows = db.prepare(`
    SELECT p.id, p.siswa_id, p.jenis_id, p.nominal, p.tanggal, p.periode, p.keterangan,
           p.nama_pihak, p.created_at,
           s.nis, COALESCE(s.nama, p.nama_pihak, '(tanpa nama)') AS nama_siswa, s.kelas,
           j.nama AS jenis
    FROM pembayaran p
    LEFT JOIN siswa s ON s.id = p.siswa_id
    LEFT JOIN jenis_sumbangan j ON j.id = p.jenis_id
    ${sqlWhere}
    ORDER BY p.tanggal DESC, p.id DESC
    LIMIT ? OFFSET ?
  `).all(...args, limit, offset);
  res.json({ rows, total, limit, offset, hasMore: offset + rows.length < total });
});

// --- Admin: auth check ---
app.get('/api/admin/check', (req, res) => {
  res.json({ ok: true, namaSekolah: settingGet('nama_sekolah'), namaAdmin: settingGet('admin_nama') });
});

app.post('/api/admin/login', (req, res) => {
  const { password } = req.body || {};
  if (!passwordOk(String(password))) return res.status(401).json({ error: 'Password salah' });
  res.json({ ok: true, namaAdmin: settingGet('admin_nama') });
});

app.post('/api/admin/password', (req, res) => {
  const { password_lama, password_baru } = req.body || {};
  if (!passwordOk(String(password_lama || ''))) {
    return res.status(403).json({ error: 'Password lama salah' });
  }
  const baru = String(password_baru || '');
  if (baru.length < 6) return res.status(400).json({ error: 'Password baru minimal 6 karakter' });
  settingSet('admin_password_hash', hashPass(baru));
  res.json({ ok: true });
});

// --- Admin: settings ---
app.get('/api/settings', (req, res) => {
  const rows = db.prepare('SELECT key, value FROM setting').all();
  const aman = rows.filter(r => !r.key.startsWith('admin_pass'));
  res.json(Object.fromEntries(aman.map(r => [r.key, r.value])));
});

app.post('/api/settings', (req, res) => {
  const { nama_sekolah, tahun_ajaran, target_tahunan, admin_nama } = req.body || {};
  if (nama_sekolah !== undefined) settingSet('nama_sekolah', String(nama_sekolah).trim());
  if (tahun_ajaran !== undefined) settingSet('tahun_ajaran', String(tahun_ajaran).trim());
  if (target_tahunan !== undefined) settingSet('target_tahunan', String(Math.max(0, Number(target_tahunan) || 0)));
  if (admin_nama !== undefined) settingSet('admin_nama', String(admin_nama).trim() || 'Admin');
  notifyChange();
  res.json({ ok: true });
});

// --- Admin: QRIS ---
app.post('/api/qris', (req, res) => {
  const { enabled, nama, atas_nama, no_va } = req.body || {};
  if (enabled !== undefined) settingSet('qris_enabled', enabled ? '1' : '0');
  if (nama !== undefined) settingSet('qris_nama', String(nama).trim());
  if (atas_nama !== undefined) settingSet('qris_atas_nama', String(atas_nama).trim());
  if (no_va !== undefined) settingSet('qris_no_va', String(no_va).trim());
  notifyChange();
  res.json({ ok: true });
});

// --- Admin: jenis sumbangan ---
app.get('/api/jenis', (req, res) => {
  const rows = db.prepare(`
    SELECT j.id, j.nama, j.target, j.aktif, IFNULL(j.kelas, '') AS kelas,
           IFNULL(SUM(p.nominal), 0) AS terkumpul
    FROM jenis_sumbangan j
    LEFT JOIN pembayaran p ON p.jenis_id = j.id
    GROUP BY j.id
    ORDER BY j.id
  `).all();
  res.json(rows);
});

app.post('/api/jenis', (req, res) => {
  const { nama, target, kelas } = req.body || {};
  if (!nama || !String(nama).trim()) return res.status(400).json({ error: 'Nama wajib diisi' });
  const kelasV = valKelasJenis(kelas);
  db.prepare('INSERT INTO jenis_sumbangan (nama, target, kelas) VALUES (?, ?, ?)').run(String(nama).trim(), Math.max(0, Number(target) || 0), kelasV);
  notifyChange();
  res.json({ ok: true });
});

app.post('/api/jenis/:id', (req, res) => {
  const id = Number(req.params.id);
  const cur = db.prepare('SELECT nama, target, aktif, kelas FROM jenis_sumbangan WHERE id = ?').get(id);
  if (!cur) return res.status(404).json({ error: 'Jenis tidak ditemukan' });
  const { nama, target, aktif, kelas } = req.body || {};
  if (nama !== undefined && !String(nama).trim()) return res.status(400).json({ error: 'Nama wajib diisi' });
  const namaV = nama !== undefined ? String(nama).trim() : cur.nama;
  const targetV = target !== undefined ? Math.max(0, Number(target) || 0) : cur.target;
  const aktifV = aktif !== undefined ? (aktif ? 1 : 0) : cur.aktif;
  const kelasV = kelas !== undefined ? valKelasJenis(kelas) : (cur.kelas || '');
  db.prepare('UPDATE jenis_sumbangan SET nama = ?, target = ?, aktif = ?, kelas = ? WHERE id = ?').run(namaV, targetV, aktifV, kelasV, id);
  notifyChange();
  res.json({ ok: true });
});

app.delete('/api/jenis/:id', (req, res) => {
  const id = Number(req.params.id);
  const used = db.prepare('SELECT COUNT(*) AS c FROM pembayaran WHERE jenis_id = ?').get(id).c;
  if (used > 0) {
    db.prepare('UPDATE jenis_sumbangan SET aktif = 0 WHERE id = ?').run(id);
  } else {
    db.prepare('DELETE FROM jenis_sumbangan WHERE id = ?').run(id);
  }
  notifyChange();
  res.json({ ok: true });
});

// --- Admin: siswa ---
app.get('/api/siswa', (req, res) => {
  const rows = db.prepare('SELECT * FROM siswa ORDER BY nis, nama, jenis_kelamin, kelas').all();
  res.json(rows);
});

app.post('/api/siswa', (req, res) => {
  const { nis, nama, kelas, jenis_kelamin, no_hp } = req.body || {};
  if (!nama || !String(nama).trim()) return res.status(400).json({ error: 'Nama siswa wajib diisi' });
  const result = db.prepare('INSERT INTO siswa (nis, nama, kelas, jenis_kelamin, no_hp) VALUES (?, ?, ?, ?, ?)')
    .run(String(nis || '').trim(), String(nama).trim(), String(kelas || '').trim(), String(jenis_kelamin || '').trim(), String(no_hp || '').trim());
  notifyChange();
  res.json({ ok: true, id: result.lastInsertRowid });
});

app.delete('/api/siswa/:id', (req, res) => {
  const id = Number(req.params.id);
  const used = db.prepare('SELECT COUNT(*) AS c FROM pembayaran WHERE siswa_id = ?').get(id).c;
  if (used > 0) return res.status(400).json({ error: 'Siswa punya riwayat pembayaran, hapus pembayarannya dulu.' });
  db.prepare('DELETE FROM siswa WHERE id = ?').run(id);
  notifyChange();
  res.json({ ok: true });
});

// --- Admin: import & template Excel siswa ---
const XLSX_ADDR = (r, c) => XLSX.utils.encode_cell({ r, c });

function xlsxTemplateSiswa() {
  const headers = ['NIS', 'Nama Lengkap', 'Jenis Kelamin', 'Kelas', 'No. HP Orang Tua/Wali'];
  const nCols = headers.length;
  const borderThin = { style: 'thin', color: { rgb: '93A5CF' } };
  const borderAll = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };
  const borderMed = { style: 'medium', color: { rgb: '2563EB' } };
  const borderAllMed = { top: borderMed, bottom: borderMed, left: borderMed, right: borderMed };
  const borderBottomAccent = { style: 'thick', color: { rgb: '1E40AF' } };
  const fillHeader = { fgColor: { rgb: '2563EB' } };
  const fillTitle = { fgColor: { rgb: '312E81' } };
  const fillSubtitle = { fgColor: { rgb: 'E0E7FF' } };
  const fillContoh = { fgColor: { rgb: 'EEF2FF' } };
  const fillWarning = { fgColor: { rgb: 'FEF3C7' } };
  const fillZebra = { fgColor: { rgb: 'F8FAFC' } };
  const fillFooter = { fgColor: { rgb: 'F1F5F9' } };

  const ws = {};
  ws['!cols'] = [
    { wch: 18 },
    { wch: 34 },
    { wch: 16 },
    { wch: 12 },
    { wch: 26 }
  ];
  const sekolah = settingGet('nama_sekolah') || '';
  const tahun = settingGet('tahun_ajaran') || '';
  const putih = { fgColor: { rgb: 'FFFFFF' } };
  const t = (r, c, v, s) => { ws[XLSX_ADDR(r, c)] = { t: 's', v: String(v), s, z: '@' }; };

  // Judul (sel tunggal, TANPA merge agar tidak menghalangi aksi paste)
  t(0, 0, 'TEMPLATE IMPORT DATA SISWA', {
    font: { bold: true, sz: 16, color: { rgb: '1E293B' } }
  });
  t(1, 0, `${sekolah}${tahun ? ' â€” Tahun Ajaran ' + tahun : ''}`, {
    font: { italic: true, sz: 11, color: { rgb: '64748B' } }
  });
  t(2, 0, 'Isi data mulai baris ke-5. Hapus BARIS CONTOH sebelum impor. Kolom NIS & No. HP tulis sebagai teks, Jenis Kelamin diisi L atau P.', {
    font: { bold: true, sz: 10, color: { rgb: '92400E' } },
    fill: { fgColor: { rgb: 'FEF3C7' } },
    alignment: { vertical: 'center', horizontal: 'left', wrapText: true }
  });

  // Baris judul kolom (indeks 3)
  headers.forEach((h, i) => {
    t(3, i, h, {
      font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
      fill: fillHeader,
      alignment: { vertical: 'center', horizontal: 'center', wrapText: true },
      border: borderAll
    });
  });

  // Baris contoh (indeks 4)
  const contoh = ['2026001 (contoh)', 'Contoh Nama Siswa', 'L', '7A', '081234567890'];
  const contohStyle = {
    font: { color: { rgb: '6D28D9' }, italic: true, sz: 10 },
    fill: fillContoh,
    alignment: { vertical: 'center', horizontal: 'left' },
    border: borderAll
  };
  contoh.forEach((v, i) => t(4, i, v, contohStyle));

  // Area isian 100 baris (indeks 5-104), tanpa merge sama sekali
  const dataStyleBase = { alignment: { vertical: 'center', horizontal: 'left' }, border: borderAll };
  for (let r = 5; r <= 104; r++) {
    const zebra = r % 2 === 0 ? fillZebra : putih;
    for (let c = 0; c < nCols; c++) {
      t(r, c, '', { ...dataStyleBase, fill: zebra });
    }
  }

  ws['!rows'] = [];
  for (let r = 0; r <= 104; r++) {
    ws['!rows'].push({ hpt: r === 3 ? 28 : 20 });
  }
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 104, c: nCols - 1 } });

  // ---------- Lembar Petunjuk ----------
  const ws2 = {};
  ws2['!cols'] = [{ wch: 24 }, { wch: 68 }, { wch: 22 }];
  const t2 = (r, c, v, s) => { ws2[XLSX_ADDR(r, c)] = { t: 's', v: String(v), s }; };
  const merges2 = [];
  const line = (r, txt, s) => {
    t2(r, 0, txt, s);
    merges2.push({ s: { r, c: 0 }, e: { r, c: 2 } });
  };
  const styleJudul = { font: { bold: true, sz: 15, color: { rgb: 'FFFFFF' } }, fill: fillTitle, alignment: { vertical: 'center', horizontal: 'left' } };
  const styleSection = { font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: fillHeader, alignment: { vertical: 'center', horizontal: 'left' } };
  const styleTeks = { font: { sz: 11, color: { rgb: '334155' } }, alignment: { vertical: 'center', horizontal: 'left' } };
  const styleHeaderTabel = { font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } }, fill: fillHeader, alignment: { vertical: 'center', horizontal: 'center' }, border: borderAll };
  const styleTabel = { font: { sz: 11, color: { rgb: '334155' } }, alignment: { vertical: 'center', horizontal: 'left' }, border: borderAll };

  line(0, 'PETUNJUK PENGISIAN TEMPLATE DATA SISWA', styleJudul);
  line(1, 'GotongRoyong â€” Sistem Informasi Sumbangan Wali Murid', { font: { sz: 10, italic: true, color: { rgb: '64748B' } }, alignment: { vertical: 'center' } });
  line(2, '', styleTeks);

  const rA = 3;
  line(rA, 'A. PERSIAPAN', styleSection);
  const persiapan = [
    '1. Unduh file template ini, lalu buka dengan Microsoft Excel, WPS Office, atau Google Sheets.',
    '2. Hapus baris CONTOH (baris ke-5 pada lembar "Data Siswa") sebelum mulai mengisi.',
    '3. Jangan mengubah, menghapus, atau menambah kolom judul. Sistem hanya membaca 5 kolom.',
    '4. Simpan salinan file dengan nama sesuai kebutuhan, misal: data_siswa_2027.xlsx.'
  ];
  persiapan.forEach((x, i) => line(rA + 1 + i, x, styleTeks));

  const rB = rA + persiapan.length + 2;
  line(rB, 'B. KETERANGAN KOLOM', styleSection);
  const kolom = ['Kolom', 'Keterangan', 'Contoh'];
  kolom.forEach((h, c) => t2(rB + 1, c, h, styleHeaderTabel));
  const barisKolom = [
    ['NIS', 'Nomor Induk Siswa. Kosongkan jika siswa tidak punya.', '2026001'],
    ['Nama Lengkap', 'Nama lengkap siswa. WAJIB DIISI.', 'Budi Setiyadi'],
    ['Jenis Kelamin', 'Isi huruf L (laki-laki) atau P (perempuan).', 'L'],
    ['Kelas', 'Kelas siswa, contoh: 7A, 8B, 9C.', '7A'],
    ['No. HP Orang Tua/Wali', 'Nomor WhatsApp wali. Dipakai untuk notifikasi pembayaran otomatis.', '081234567890']
  ];
  barisKolom.forEach((row, i) => row.forEach((v, c) => t2(rB + 2 + i, c, v, { ...styleTabel, fill: i % 2 === 0 ? { fgColor: { rgb: 'F8FAFC' } } : { fgColor: { rgb: 'FFFFFF' } } })));

  const rC = rB + 2 + barisKolom.length + 2;
  line(rC, 'C. KETENTUAN IMPOR', styleSection);
  const ketentuan = [
    '1. Kolom yang wajib diisi hanya "Nama Lengkap". Kolom lain boleh dikosongkan.',
    '2. Siswa dengan NIS atau nama yang sama dengan data yang sudah ada akan dilewati',
    '   otomatis agar tidak terjadi data ganda.',
    '3. Simpan file dalam format .xlsx atau .xls, lalu gunakan tombol "Import Excel"',
    '   pada menu Data Siswa di aplikasi GotongRoyong.',
    '4. Setelah impor, aplikasi menampilkan jumlah siswa yang ditambahkan dan dilewati.',
    '5. Pastikan No. HP valid agar notifikasi WhatsApp terkirim saat pembayaran dicatat.'
  ];
  ketentuan.forEach((x, i) => line(rC + 1 + i, x, styleTeks));

  const rFooter = rC + ketentuan.length + 2;
  line(rFooter, 'Butuh bantuan? Hubungi admin GotongRoyong.', { font: { bold: true, sz: 11, color: { rgb: '4F46E5' } }, alignment: { vertical: 'center' } });
  line(rFooter + 1, 'Â© GotongRoyong â€” Sistem Informasi Sumbangan Wali Murid', { font: { sz: 10, italic: true, color: { rgb: '94A3B8' } }, alignment: { vertical: 'center' } });

  ws2['!merges'] = merges2;
  ws2['!rows'] = [{ hpt: 26 }, { hpt: 16 }, { hpt: 10 }];
  ws2['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rFooter + 1, c: 2 } });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data Siswa');
  XLSX.utils.book_append_sheet(wb, ws2, 'Petunjuk');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.get('/api/siswa/template', (req, res) => {
  const buf = xlsxTemplateSiswa();
  res.setHeader('Content-Disposition', 'attachment; filename="template_import_siswa_gotongroyong.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

app.post('/api/siswa/import', (req, res) => {
  const { data } = req.body || {};
  if (!data) return res.status(400).json({ error: 'File tidak ditemukan' });
  let wb;
  try {
    wb = XLSX.read(Buffer.from(String(data), 'base64'), { type: 'buffer' });
  } catch (e) {
    return res.status(400).json({ error: 'File tidak dapat dibaca. Pastikan format .xls/.xlsx dari template.' });
  }
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  let startRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    if (/nis/i.test(String(r[0] ?? '')) && /nama/i.test(String(r[1] ?? ''))) {
      startRow = i + 1;
      break;
    }
  }
  if (startRow < 0) {
    return res.status(400).json({ error: 'Format template tidak dikenali. Unduh template terbaru dari tombol "Unduh Template".' });
  }
  let ditambahkan = 0;
  let dilewati = 0;
  const ins = db.prepare('INSERT INTO siswa (nis, nama, kelas, jenis_kelamin, no_hp) VALUES (?, ?, ?, ?, ?)');
  const seenNis = new Set(db.prepare("SELECT nis FROM siswa WHERE nis != ''").all().map(r => r.nis.toLowerCase()));
  const seenNama = new Set(db.prepare('SELECT nama FROM siswa').all().map(r => r.nama.toLowerCase()));
  for (let i = startRow; i < rows.length; i++) {
    const r = rows[i] || [];
    const nis = String(r[0] ?? '').trim();
    const nama = String(r[1] ?? '').trim();
    const jk = String(r[2] ?? '').trim().toUpperCase();
    const jenisKelamin = (jk === 'L' || jk === 'P') ? jk : '';
    const kelas = String(r[3] ?? '').trim();
    const noHp = String(r[4] ?? '').trim().replace(/[^\d+]/g, '');
    if (!nama) continue;
    if (/contoh/i.test(nis) || /contoh/i.test(nama)) continue;
    if (nis && seenNis.has(nis.toLowerCase())) {
      dilewati++;
      continue;
    }
    if (seenNama.has(nama.toLowerCase())) {
      dilewati++;
      continue;
    }
    seenNis.add(nis.toLowerCase());
    seenNama.add(nama.toLowerCase());
    ins.run(nis, nama, kelas, jenisKelamin, noHp);
    ditambahkan++;
  }
  notifyChange();
  res.json({ ok: true, ditambahkan, dilewati });
});

app.post('/api/siswa/bulk-delete', (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) return res.status(400).json({ error: 'Tidak ada siswa yang dipilih' });
  const bersih = [...new Set(ids.map(Number).filter(Number.isFinite))].filter(n => n > 0);
  if (!bersih.length) return res.status(400).json({ error: 'Data siswa tidak valid' });
  const tempat = bersih.map(() => '?').join(',');
  const punyaRiwayat = db.prepare(`SELECT DISTINCT siswa_id FROM pembayaran WHERE siswa_id IN (${tempat})`).all(...bersih);
  if (punyaRiwayat.length) {
    return res.status(400).json({ error: 'Ada siswa terpilih yang punya riwayat pembayaran. Hapus pembayarannya dulu.' });
  }
  db.prepare(`DELETE FROM siswa WHERE id IN (${tempat})`).run(...bersih);
  notifyChange();
  res.json({ ok: true, dihapus: bersih.length });
});

app.post('/api/siswa/:id', (req, res) => {
  const id = Number(req.params.id);
  const cur = db.prepare('SELECT * FROM siswa WHERE id = ?').get(id);
  if (!cur) return res.status(404).json({ error: 'Siswa tidak ditemukan' });
  const { nis, nama, kelas, jenis_kelamin, no_hp } = req.body || {};
  if (nama !== undefined && !String(nama).trim()) return res.status(400).json({ error: 'Nama siswa wajib diisi' });
  const namaV = nama !== undefined ? String(nama).trim() : cur.nama;
  const nisV = nis !== undefined ? String(nis).trim() : (cur.nis || '');
  const kelasV = kelas !== undefined ? String(kelas).trim() : (cur.kelas || '');
  const jkV = jenis_kelamin !== undefined ? String(jenis_kelamin).trim() : (cur.jenis_kelamin || '');
  const hpV = no_hp !== undefined ? String(no_hp).trim() : (cur.no_hp || '');
  db.prepare('UPDATE siswa SET nis = ?, nama = ?, kelas = ?, jenis_kelamin = ?, no_hp = ? WHERE id = ?').run(nisV, namaV, kelasV, jkV, hpV, id);
  notifyChange();
  res.json({ ok: true });
});

// --- Admin: pembayaran ---
app.post('/api/pembayaran', (req, res) => {
  const { siswa_id, jenis_id, nominal, tanggal, keterangan, nama_pihak } = req.body || {};
  if (!jenis_id) { if (!siswa_id && !String(nama_pihak || '').trim()) return res.status(400).json({ error: 'Jenis sumbangan wajib diisi' }); }
  if (!jenis_id) return res.status(400).json({ error: 'Siswa dan jenis sumbangan wajib diisi' });
  if (!nominal || Number(nominal) <= 0) return res.status(400).json({ error: 'Nominal harus lebih dari 0' });
  const j = db.prepare('SELECT id, nama FROM jenis_sumbangan WHERE id = ?').get(Number(jenis_id));
  if (!j) return res.status(400).json({ error: 'Jenis sumbangan tidak ditemukan' });
  const nominalV = Number(nominal);
  const tanggalV = tanggal || fmtDate(new Date());
  const createdAt = new Date().toISOString();
  const namaPihakV = String(nama_pihak || '').trim();

  let s = null;
  let siswaIdV = null;
  let waNotif = null;
  if (siswa_id) {
    s = db.prepare('SELECT id, nama, nis, kelas, no_hp FROM siswa WHERE id = ?').get(Number(siswa_id));
    if (!s) return res.status(400).json({ error: 'Siswa tidak ditemukan' });
    siswaIdV = Number(siswa_id);
  } else if (!namaPihakV) {
    return res.status(400).json({ error: 'Siswa atau nama pihak wajib diisi' });
  }

  const info = db.prepare('INSERT INTO pembayaran (siswa_id, jenis_id, nominal, tanggal, keterangan, nama_pihak, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
    .run(siswaIdV, Number(jenis_id), nominalV, tanggalV, String(keterangan || '').trim(), namaPihakV, createdAt);
  if (s && s.no_hp && settingGet('wa_enabled') === '1') {
    kirimWa(s.no_hp, buatPesanPembayaran({
      nama: s.nama,
      kelas: s.kelas,
      jenis: j.nama,
      nominal: nominalV,
      tanggal: tanggalV
    })).then(r => { if (!r.ok) console.log('WA notifikasi gagal:', r.error); });
  }
  notifyChange();
  res.json({
    ok: true,
    struk: {
      id: Number(info.lastInsertRowid),
      nis: (s && s.nis) ? s.nis : '',
      nama: namaPihakV || (s ? s.nama : ''),
      kelas: (s && s.kelas) ? s.kelas : '',
      pihakLain: !s,
      jenis: j.nama,
      nominal: nominalV,
      tanggal: tanggalV,
      keterangan: String(keterangan || '').trim(),
      createdAt,
      sekolah: settingGet('nama_sekolah') || '',
      tahunAjaran: settingGet('tahun_ajaran') || '',
      adminNama: settingGet('admin_nama') || '',
      totalTerkumpul: db.prepare('SELECT IFNULL(SUM(nominal), 0) AS t FROM pembayaran').get().t
    }
  });
});

app.post('/api/pembayaran/:id', (req, res) => {
  const id = Number(req.params.id);
  const cur = db.prepare('SELECT * FROM pembayaran WHERE id = ?').get(id);
  if (!cur) return res.status(404).json({ error: 'Pembayaran tidak ditemukan' });
  const { siswa_id, jenis_id, nominal, tanggal, keterangan, nama_pihak } = req.body || {};
  const siswaV = siswa_id !== undefined ? (siswa_id ? Number(siswa_id) : null) : cur.siswa_id;
  const jenisV = jenis_id !== undefined ? Number(jenis_id) : cur.jenis_id;
  const nominalV = nominal !== undefined ? Number(nominal) : cur.nominal;
  if (!jenisV) return res.status(400).json({ error: 'Jenis sumbangan wajib diisi' });
  if (!nominalV || nominalV <= 0) return res.status(400).json({ error: 'Nominal harus lebih dari 0' });
  const namaPihakV = nama_pihak !== undefined ? String(nama_pihak).trim() : (cur.nama_pihak || '');
  if (siswaV) {
    const s = db.prepare('SELECT id FROM siswa WHERE id = ?').get(siswaV);
    if (!s) return res.status(400).json({ error: 'Siswa tidak ditemukan' });
  } else if (!namaPihakV) {
    return res.status(400).json({ error: 'Siswa atau nama pihak wajib diisi' });
  }
  const j = db.prepare('SELECT id FROM jenis_sumbangan WHERE id = ?').get(jenisV);
  if (!j) return res.status(400).json({ error: 'Jenis sumbangan tidak ditemukan' });
  const tanggalV = tanggal !== undefined ? tanggal : cur.tanggal;
  const keteranganV = keterangan !== undefined ? String(keterangan).trim() : (cur.keterangan || '');
  db.prepare('UPDATE pembayaran SET siswa_id = ?, jenis_id = ?, nominal = ?, tanggal = ?, keterangan = ?, nama_pihak = ? WHERE id = ?')
    .run(siswaV || null, jenisV, nominalV, tanggalV, keteranganV, namaPihakV, id);
  notifyChange();
  res.json({ ok: true });
});

app.delete('/api/pembayaran/:id', (req, res) => {
  db.prepare('DELETE FROM pembayaran WHERE id = ?').run(Number(req.params.id));
  notifyChange();
  res.json({ ok: true });
});

// ---------- Kwitansi / bukti pembayaran ----------
const ANGKA = ['Nol', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'];

function bilinear(n) {
  const s = String(n);
  const len = s.length;
  if (len === 1) return ANGKA[Number(s)];
  if (len === 2) return ANGKA[Number(s[0])] + ' Puluh' + (s[1] === '0' ? '' : ' ' + ANGKA[Number(s[1])]);
  if (len === 3) {
    const r = ANGKA[Number(s[0])] + ' Ratus';
    return s.slice(1) === '00' ? r : r + ' ' + bilinear(s.slice(1));
  }
  return n;
}

function Process(n) {
  n = Math.floor(Math.abs(Number(n) || 0));
  const satuan = ['', ' Satu', ' Dua', ' Tiga', ' Empat', ' Lima', ' Enam', ' Tujuh', ' Delapan', ' Sembilan'];
  const konversi = ['', ' Ribu', ' Juta', ' Milyar', ' Trilyun'];
  if (n === 0) return 'Nol Rupiah';
  let hasil = '';
  let idx = 0;
  let sisa = n;
  while (sisa > 0) {
    const bagian = sisa % 1000;
    if (bagian > 0) hasil = bilinear(bagian) + satuan[bagian % 10] + konversi[idx] + hasil;
    sisa = Math.floor(sisa / 1000);
    idx++;
  }
  return hasil.trim() + ' Rupiah';
}

function dataKwitansi(id) {
  const r = db.prepare(`
    SELECT p.*, s.nis, s.nama AS nama_siswa, s.kelas, s.no_hp, j.nama AS jenis
    FROM pembayaran p
    LEFT JOIN siswa s ON s.id = p.siswa_id
    LEFT JOIN jenis_sumbangan j ON j.id = p.jenis_id
    WHERE p.id = ?
  `).get(Number(id));
  if (!r) return null;
  const tahun = (settingGet('tahun_ajaran') || '').split('/')[0] || String(new Date().getFullYear());
  return {
    id: r.id,
    no: `KW/${tahun}/${String(r.id).padStart(5, '0')}`,
    nis: r.nis || '',
    nama: r.nama_siswa || r.nama_pihak || '',
    kelas: r.kelas || '',
    pihakLain: !r.siswa_id,
    jenis: r.jenis || '-',
    nominal: Number(r.nominal || 0),
    tanggal: r.tanggal,
    periode: r.periode || '',
    keterangan: r.keterangan || '',
    created_at: r.created_at,
    sekolah: settingGet('nama_sekolah') || 'GotongRoyong',
    tahunAjaran: settingGet('tahun_ajaran') || '',
    logo: settingGet('logo') || '',
    adminNama: settingGet('admin_nama') || 'Admin',
    totalTerkumpul: db.prepare('SELECT IFNULL(SUM(nominal), 0) AS t FROM pembayaran').get().t,
    saldo: db.prepare('SELECT IFNULL(SUM(nominal), 0) - IFNULL((SELECT SUM(nominal) FROM penggunaan_dana), 0) AS t FROM pembayaran').get().t,
    terbilang: Process(r.nominal)
  };
}

function kwitansiPayload(ids) {
  const list = [...new Set((Array.isArray(ids) ? ids : []).map(Number).filter(Boolean))].slice(0, 100);
  return list.map(dataKwitansi).filter(Boolean);
}

app.get('/api/kwitansi/:id', (req, res) => {
  const data = dataKwitansi(req.params.id);
  if (!data) return res.status(404).json({ error: 'Data pembayaran tidak ditemukan' });
  res.json(data);
});

app.post('/api/kwitansi', (req, res) => {
  const ids = String((req.body || {}).ids || '').split(',');
  const list = kwitansiPayload(ids);
  if (!list.length) return res.status(400).json({ error: 'Tidak ada data kwitansi yang valid' });
  res.json({ rows: list });
});

app.post('/api/kwitansi/:id/wa', async (req, res) => {
  const data = dataKwitansi(req.params.id);
  if (!data) return res.status(404).json({ error: 'Data pembayaran tidak ditemukan' });
  const row = db.prepare('SELECT no_hp FROM siswa WHERE id = (SELECT siswa_id FROM pembayaran WHERE id = ?)').get(Number(req.params.id));
  const noHp = String((req.body || {}).no_hp || (row && row.no_hp) || '').trim();
  if (!noHp) return res.status(400).json({ error: 'Nomor WhatsApp wali murid belum tersedia' });
  if (!waState.ready) return res.status(400).json({ error: 'WhatsApp belum terhubung' });
  const teks = [
    `Assalamu'alaikum Bapak/Ibu wali dari *${data.nama}*${data.kelas ? ' (' + data.kelas + ')' : ''},`,
    '',
    `Berikut bukti pembayaran sumbangan di *${data.sekolah}*.`,
    `ðŸ§¾ No. Kwitansi : *${data.no}*`,
    `ðŸ“‹ Jenis         : *${data.jenis}*`,
    `ðŸ’° Nominal       : *Rp ${data.nominal.toLocaleString('id-ID')}*`,
    `ðŸ“… Tanggal       : ${data.tanggal}`,
    '',
    'Terima kasih atas partisipasi Bapak/Ibu.'
  ].join('\n');
  const hasil = await kirimWa(noHp, teks);
  if (!hasil.ok) return res.status(400).json({ error: hasil.error });
  res.json({ ok: true });
});

// --- Admin: logo sekolah ---
app.post('/api/logo', (req, res) => {
  const { data } = req.body || {};
  if (!data) return res.status(400).json({ error: 'Data logo kosong' });
  let buf, ext = 'png';
  const m = String(data).match(/^data:image\/([a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (m) {
    ext = m[1].split('+')[0].toLowerCase();
    buf = Buffer.from(m[2], 'base64');
  } else if (/^[A-Za-z0-9+/=\s]+$/.test(String(data))) {
    buf = Buffer.from(String(data), 'base64');
  } else {
    return res.status(400).json({ error: 'Format gambar tidak valid' });
  }
  if (!['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'].includes(ext)) {
    return res.status(400).json({ error: 'Format gambar tidak didukung' });
  }
  if (!buf.length) return res.status(400).json({ error: 'Gambar kosong' });
  const old = settingGet('logo');
  if (old && old.startsWith('/uploads/')) {
    try { fs.unlinkSync(path.join(UPLOAD_DIR, path.basename(old))); } catch (e) {}
  }
  const name = 'logo.' + ext;
  fs.writeFileSync(path.join(UPLOAD_DIR, name), buf);
  settingSet('logo', '/uploads/' + name);
  notifyChange();
  res.json({ ok: true, logo: '/uploads/' + name });
});

app.delete('/api/logo', (req, res) => {
  const old = settingGet('logo');
  if (old && old.startsWith('/uploads/')) {
    try { fs.unlinkSync(path.join(UPLOAD_DIR, path.basename(old))); } catch (e) {}
  }
  db.prepare('DELETE FROM setting WHERE key = ?').run('logo');
  notifyChange();
  res.json({ ok: true });
});

// --- Admin: penggunaan dana ---
const DOKUMEN_EXT = ['pdf', 'jpg', 'jpeg', 'png', 'gif', 'webp', 'doc', 'docx', 'xls', 'xlsx'];

function simpanDokumen(dataB64, namaAsli) {
  if (!dataB64) return null;
  let buf;
  const m = String(dataB64).match(/^data:[^;]+;base64,(.+)$/);
  if (m) buf = Buffer.from(m[1], 'base64');
  else if (/^[A-Za-z0-9+/=\s]+$/.test(String(dataB64))) buf = Buffer.from(String(dataB64), 'base64');
  else return null;
  if (!buf.length || buf.length > 15 * 1024 * 1024) return null;
  const clean = path.basename(String(namaAsli || 'dokumen')).replace(/[^\w.\-]+/g, '_');
  const ext = (clean.split('.').pop() || '').toLowerCase();
  if (!DOKUMEN_EXT.includes(ext)) return null;
  const name = `peng-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;
  fs.writeFileSync(path.join(UPLOAD_DIR, name), buf);
  return '/uploads/' + name;
}

function hapusFileDokumen(url) {
  if (url && url.startsWith('/uploads/')) {
    try { fs.unlinkSync(path.join(UPLOAD_DIR, path.basename(url))); } catch (e) {}
  }
}

app.get('/api/penggunaan', (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 100, 300);
  const rows = db.prepare(`
    SELECT * FROM penggunaan_dana
    ORDER BY tanggal DESC, id DESC
    LIMIT ?
  `).all(limit);
  res.json(rows);
});

app.post('/api/penggunaan', (req, res) => {
  const { tanggal, keterangan, nominal, dokumen_data, dokumen_nama } = req.body || {};
  if (!keterangan || !String(keterangan).trim()) return res.status(400).json({ error: 'Keterangan wajib diisi' });
  if (!nominal || Number(nominal) <= 0) return res.status(400).json({ error: 'Nominal harus lebih dari 0' });
  const dokumen = simpanDokumen(dokumen_data, dokumen_nama);
  db.prepare('INSERT INTO penggunaan_dana (tanggal, keterangan, nominal, dokumen, created_at) VALUES (?, ?, ?, ?, ?)')
    .run(tanggal || fmtDate(new Date()), String(keterangan).trim(), Number(nominal), dokumen, new Date().toISOString());
  notifyChange();
  res.json({ ok: true });
});

app.delete('/api/penggunaan/:id', (req, res) => {
  const id = Number(req.params.id);
  const cur = db.prepare('SELECT dokumen FROM penggunaan_dana WHERE id = ?').get(id);
  if (cur) hapusFileDokumen(cur.dokumen);
  db.prepare('DELETE FROM penggunaan_dana WHERE id = ?').run(id);
  notifyChange();
  res.json({ ok: true });
});

// --- Rekap per kelas ---
app.get('/api/rekap/kelas', (req, res) => {
  const rows = db.prepare(`
    SELECT s.kelas,
           COUNT(DISTINCT s.id) AS jumlah_siswa,
           COUNT(DISTINCT p.siswa_id) AS siswa_bayar,
           IFNULL(SUM(p.nominal), 0) AS total
    FROM siswa s
    LEFT JOIN pembayaran p ON p.siswa_id = s.id
    GROUP BY s.kelas
    ORDER BY s.kelas
  `).all();
  res.json(rows.map(r => ({ ...r, kelas: r.kelas || '(tanpa kelas)' })));
});

// --- Rekap bulanan (tren pemasukan & penggunaan per bulan) ---
const NAMA_BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function labelBulan(kunci) {
  const [y, m] = String(kunci || '').split('-');
  const idx = Number(m) - 1;
  if (!y || idx < 0 || idx > 11) return kunci || '';
  return `${NAMA_BULAN[idx]} ${y}`;
}

function normalisasiBulan(v) {
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const d = s ? new Date(s) : new Date();
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function daftarBulan(n = 6, akhir = null) {
  const end = akhir ? new Date(akhir + '-01T00:00:00') : new Date();
  if (Number.isNaN(end.getTime())) return [];
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(end.getFullYear(), end.getMonth() - i, 1);
    out.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  return out;
}

function rekapBulanan(dari, sampai) {
  const mulai = normalisasiBulan(dari) || daftarBulan(6)[0];
  const akhir = normalisasiBulan(sampai) || daftarBulan(6)[5];
  const kunci = [];
  const [ay, am] = mulai.split('-').map(Number);
  const [by, bm] = akhir.split('-').map(Number);
  let y = ay, m = am;
  let guard = 0;
  while ((y < by || (y === by && m <= bm)) && guard++ < 240) {
    kunci.push(`${y}-${String(m).padStart(2, '0')}`);
    m++; if (m > 12) { m = 1; y++; }
  }
  const inClaus = kunci.map(() => '?').join(',') || "''";
  const arg = kunci;

  const bayar = db.prepare(`
    SELECT strftime('%Y-%m', tanggal) AS bulan,
           IFNULL(SUM(nominal), 0) AS total,
           COUNT(*) AS transaksi,
           COUNT(DISTINCT siswa_id) AS siswa_bayar
    FROM pembayaran
    WHERE strftime('%Y-%m', tanggal) IN (${inClaus})
    GROUP BY bulan
  `).all(...arg);
  const pakai = db.prepare(`
    SELECT strftime('%Y-%m', tanggal) AS bulan, IFNULL(SUM(nominal), 0) AS total
    FROM penggunaan_dana
    WHERE strftime('%Y-%m', tanggal) IN (${inClaus})
    GROUP BY bulan
  `).all(...arg);
  const jenis = db.prepare(`
    SELECT strftime('%Y-%m', p.tanggal) AS bulan, j.nama,
           IFNULL(SUM(p.nominal), 0) AS total
    FROM pembayaran p
    JOIN jenis_sumbangan j ON j.id = p.jenis_id
    WHERE strftime('%Y-%m', p.tanggal) IN (${inClaus})
    GROUP BY bulan, j.id
    ORDER BY bulan, j.id
  `).all(...arg);

  const petaBayar = new Map(bayar.map(r => [r.bulan, r]));
  const petaPakai = new Map(pakai.map(r => [r.bulan, r]));
  const petaJenis = new Map();
  for (const r of jenis) {
    if (!petaJenis.has(r.bulan)) petaJenis.set(r.bulan, []);
    petaJenis.get(r.bulan).push({ nama: r.nama, total: r.total });
  }

  let saldo = 0;
  const rows = kunci.map(b => {
    const by = petaBayar.get(b);
    const pk = petaPakai.get(b);
    const total = by ? Number(by.total) : 0;
    const penggunaan = pk ? Number(pk.total) : 0;
    saldo += total - penggunaan;
    return {
      bulan: b,
      label: labelBulan(b),
      total,
      penggunaan,
      saldo: saldo,
      transaksi: by ? Number(by.transaksi) : 0,
      siswa_bayar: by ? Number(by.siswa_bayar) : 0,
      jenis: petaJenis.get(b) || []
    };
  });
  return rows;
}

app.get('/api/rekap/bulanan', (req, res) => {
  const rows = rekapBulanan(req.query.dari, req.query.sampai);
  res.json({
    rows,
    total: rows.reduce((s, r) => s + r.total, 0),
    penggunaan: rows.reduce((s, r) => s + r.penggunaan, 0),
    saldo: rows.length ? rows[rows.length - 1].saldo : 0
  });
});

app.get('/api/laporan/rekap-bulanan', (req, res) => {
  const rows = rekapBulanan(req.query.dari, req.query.sampai);
  const aoa = [
    ['No', 'Bulan', 'Pemasukan (Rp)', 'Penggunaan Dana (Rp)', 'Saldo Kumulatif (Rp)', 'Transaksi', 'Siswa Bayar'],
    ...rows.map((r, i) => [i + 1, r.label, r.total, r.penggunaan, r.saldo, r.transaksi, r.siswa_bayar])
  ];
  aoa.push([]);
  aoa.push(['', 'TOTAL', rows.reduce((s, r) => s + r.total, 0), rows.reduce((s, r) => s + r.penggunaan, 0), '', '', '']);
  sendXlsx(res, aoa, 'laporan_rekap_bulanan.xls', 'Rekap Bulanan');
});

// --- Export Excel laporan ---
function sendXlsx(res, aoa, filename, sheetName = 'Laporan') {
  const ws = {};
  const nCols = aoa[0].length;
  ws['!cols'] = aoa[0].map((_, i) => ({ wch: Math.max(12, Math.min(32, String(aoa[0][i] || '').length + 6)) }));
  ws['!rows'] = [{ hpt: 26 }, { hpt: 22 }];
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length, c: nCols - 1 } });
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: nCols - 1 } }];
  const borderThin = { style: 'thin', color: { rgb: 'CBD5E1' } };
  const borderAll = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };
  const sekolah = settingGet('nama_sekolah') || '';

  ws[XLSX_ADDR(0, 0)] = {
    t: 's', v: `${sheetName.toUpperCase()} â€” ${sekolah || 'GotongRoyong'}`,
    s: { font: { bold: true, sz: 14, color: { rgb: '1E293B' } }, alignment: { vertical: 'center', horizontal: 'left' } }
  };
  const headerStyle = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { fgColor: { rgb: '4F46E5' } },
    alignment: { vertical: 'center', horizontal: 'center' },
    border: borderAll
  };
  const bodyStyle = { alignment: { vertical: 'center', horizontal: 'left' }, border: borderAll };
  const bodyNumStyle = { alignment: { vertical: 'center', horizontal: 'right' }, border: borderAll };

  aoa[0].forEach((h, c) => {
    ws[XLSX_ADDR(1, c)] = { t: 's', v: String(h), s: headerStyle };
  });
  for (let r = 1; r < aoa.length; r++) {
    aoa[r].forEach((v, c) => {
      const isNum = typeof v === 'number';
      ws[XLSX_ADDR(r + 1, c)] = { t: isNum ? 'n' : 's', v, s: isNum ? bodyNumStyle : bodyStyle };
    });
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xls' });
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.ms-excel');
  res.send(buf);
}

function buildFilters(req) {
  const dari = req.query.dari || null;
  const sampai = req.query.sampai || null;
  const kelas = req.query.kelas || null;
  const jenis = req.query.jenis ? Number(req.query.jenis) : null;
  const q = String(req.query.q || '').trim();
  const conds = [];
  const args = [];
  if (dari) { conds.push('p.tanggal >= ?'); args.push(String(dari)); }
  if (sampai) { conds.push('p.tanggal <= ?'); args.push(String(sampai)); }
  if (kelas) { conds.push('s.kelas = ?'); args.push(String(kelas)); }
  if (jenis) { conds.push('p.jenis_id = ?'); args.push(jenis); }
  if (q) {
    conds.push("(COALESCE(s.nama, p.nama_pihak, '') LIKE ? OR IFNULL(s.nis, '') LIKE ? OR IFNULL(s.kelas, '') LIKE ? OR IFNULL(p.keterangan, '') LIKE ?)");
    const like = `%${q}%`;
    args.push(like, like, like, like);
  }
  return { where: conds.length ? 'WHERE ' + conds.join(' AND ') : '', args };
}

app.get('/api/laporan/pembayaran', (req, res) => {
  const { where, args } = buildFilters(req);
  const rows = db.prepare(`
    SELECT p.id, s.nis, COALESCE(s.nama, p.nama_pihak, '') AS nama, s.kelas, j.nama AS jenis,
           p.nominal, p.tanggal, p.keterangan, p.created_at
    FROM pembayaran p
    LEFT JOIN siswa s ON s.id = p.siswa_id
    LEFT JOIN jenis_sumbangan j ON j.id = p.jenis_id
    ${where}
    ORDER BY p.tanggal, p.id
  `).all(...args);
  const total = rows.reduce((s, r) => s + Number(r.nominal || 0), 0);
  const aoa = [
    ['No', 'NIS', 'Nama', 'Kelas', 'Jenis Sumbangan', 'Nominal (Rp)', 'Tanggal', 'Keterangan'],
    ...rows.map((r, i) => [i + 1, r.nis || '', r.nama, r.kelas || '', r.jenis || '-', r.nominal, r.tanggal, r.keterangan || ''])
  ];
  aoa.push([]);
  aoa.push(['', '', '', '', 'TOTAL', total, '', '']);
  sendXlsx(res, aoa, 'laporan_pembayaran.xls', 'Pembayaran');
});

app.get('/api/laporan/penggunaan', (req, res) => {
  const dari = req.query.dari || null;
  const sampai = req.query.sampai || null;
  const conds = [];
  const args = [];
  if (dari) { conds.push('tanggal >= ?'); args.push(String(dari)); }
  if (sampai) { conds.push('tanggal <= ?'); args.push(String(sampai)); }
  const where = conds.length ? 'WHERE ' + conds.join(' AND ') : '';
  const rows = db.prepare(`
    SELECT * FROM penggunaan_dana
    ${where}
    ORDER BY tanggal, id
  `).all(...args);
  const aoa = [
    ['No', 'Tanggal', 'Keterangan', 'Nominal (Rp)'],
    ...rows.map((r, i) => [i + 1, r.tanggal, r.keterangan, r.nominal])
  ];
  sendXlsx(res, aoa, 'laporan_penggunaan_dana.xls', 'Penggunaan Dana');
});

app.get('/api/laporan/rekap-kelas', (req, res) => {
  const dari = req.query.dari || null;
  const sampai = req.query.sampai || null;
  const conds = [];
  const args = [];
  if (dari) { conds.push('p.tanggal >= ?'); args.push(String(dari)); }
  if (sampai) { conds.push('p.tanggal <= ?'); args.push(String(sampai)); }
  const where = conds.length ? 'AND ' + conds.join(' AND ') : '';
  const rows = db.prepare(`
    SELECT s.kelas,
           COUNT(DISTINCT s.id) AS jumlah_siswa,
           COUNT(DISTINCT p.siswa_id) AS siswa_bayar,
           IFNULL(SUM(p.nominal), 0) AS total
    FROM siswa s
    LEFT JOIN pembayaran p ON p.siswa_id = s.id ${where}
    GROUP BY s.kelas
    ORDER BY s.kelas
  `).all(...args);
  const totalSiswa = rows.reduce((a, r) => a + r.jumlah_siswa, 0);
  const totalBayar = rows.reduce((a, r) => a + r.siswa_bayar, 0);
  const totalNominal = rows.reduce((a, r) => a + r.total, 0);
  const aoa = [
    ['No', 'Kelas', 'Jumlah Siswa', 'Siswa Sudah Bayar', 'Belum Bayar', 'Persentase (%)', 'Total Terkumpul (Rp)'],
    ...rows.map((r, i) => {
      const pct = r.jumlah_siswa > 0 ? Math.round((r.siswa_bayar / r.jumlah_siswa) * 100) : 0;
      return [i + 1, r.kelas || '(tanpa kelas)', r.jumlah_siswa, r.siswa_bayar, r.jumlah_siswa - r.siswa_bayar, pct, r.total];
    }),
    [],
    ['', 'TOTAL', totalSiswa, totalBayar, totalSiswa - totalBayar,
      totalSiswa > 0 ? Math.round((totalBayar / totalSiswa) * 100) : 0, totalNominal]
  ];
  sendXlsx(res, aoa, 'laporan_rekap_kelas.xls', 'Rekap per Kelas');
});

// --- Backup & Restore ---
const TABEL_DIBUTUHKAN = ['setting', 'jenis_sumbangan', 'siswa', 'pembayaran', 'penggunaan_dana'];

app.get('/api/backup/download', (req, res) => {
  const tmp = path.join(DATA_DIR, `backup-${Date.now()}.db`);
  try {
    db.exec(`VACUUM INTO '${tmp.replace(/'/g, "''")}'`);
    const tanggal = new Date().toISOString().slice(0, 10);
    res.download(tmp, `backup-sumbangan-${tanggal}.sqlite`, (err) => {
      try { fs.unlinkSync(tmp); } catch (e) {}
      if (err) console.error(err);
    });
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (_) {}
    res.status(500).json({ error: 'Gagal membuat backup: ' + e.message });
  }
});

app.post('/api/backup/restore', (req, res) => {
  const { data } = req.body || {};
  if (!data) return res.status(400).json({ error: 'File backup tidak ditemukan' });
  let buf;
  const m = String(data).match(/^data:[^;]+;base64,(.+)$/);
  if (m) buf = Buffer.from(m[1], 'base64');
  else if (/^[A-Za-z0-9+/=\s]+$/.test(String(data))) buf = Buffer.from(String(data), 'base64');
  else return res.status(400).json({ error: 'Data backup tidak valid' });
  if (!buf.length) return res.status(400).json({ error: 'File backup kosong' });
  if (buf.length > 50 * 1024 * 1024) return res.status(400).json({ error: 'File backup maks 50MB' });

  const tmp = path.join(DATA_DIR, `restore-${Date.now()}.db`);
  let namaTabel;
  try {
    fs.writeFileSync(tmp, buf);
    const check = new DatabaseSync(tmp, { readOnly: true });
    namaTabel = check.prepare("SELECT name FROM sqlite_master WHERE type='table'").all().map(r => r.name);
    check.close();
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (_) {}
    return res.status(400).json({ error: 'File bukan backup yang valid.' });
  }
  const kurang = TABEL_DIBUTUHKAN.filter(t => !namaTabel.includes(t));
  if (kurang.length) {
    try { fs.unlinkSync(tmp); } catch (_) {}
    return res.status(400).json({ error: 'File bukan backup valid (tabel hilang: ' + kurang.join(', ') + ')' });
  }
  try {
    try {
      db.exec(`VACUUM INTO '${path.join(DATA_DIR, `safety-${Date.now()}.db`).replace(/'/g, "''")}'`);
    } catch (_) {}
    db.close();
    fs.copyFileSync(tmp, DB_PATH);
    db = initDb();
    fs.unlinkSync(tmp);
    notifyChange();
    res.json({ ok: true, pesan: 'Data berhasil dipulihkan.' });
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (_) {}
    res.status(500).json({ error: 'Gagal memulihkan data: ' + e.message });
  }
});

// --- WhatsApp: status & kontrol ---
app.get('/api/wa/status', (req, res) => {
  res.json({
    lib: !!WA_LIB,
    enabled: settingGet('wa_enabled') === '1',
    status: waState.status,
    ready: waState.ready,
    nomor: waState.nomor,
    admin: settingGet('wa_admin') || '',
    error: waState.error
  });
});

app.post('/api/wa/settings', (req, res) => {
  const { enabled, admin } = req.body || {};
  if (admin !== undefined) {
    const bersih = String(admin).split(/[,\s;]+/).map((x) => x.replace(/[^\d]/g, '')).filter((x) => x.length >= 8);
    settingSet('wa_admin', bersih.join(','));
  }
  if (req.body.pesan_pengingat !== undefined) {
    settingSet('wa_pesan_pengingat', String(req.body.pesan_pengingat).slice(0, 2000));
  }
  if (req.body.jeda !== undefined) {
    settingSet('wa_pengingat_jeda', String(Math.min(Math.max(Number(req.body.jeda) || 2500, 800), 60000)));
  }
  if (req.body.batas !== undefined) {
    settingSet('wa_pengingat_batas', String(Math.min(Math.max(Number(req.body.batas) || 0, 0), 500)));
  }
  if (enabled !== undefined) {
    if (enabled) {
      settingSet('wa_enabled', '1');
      waStart();
    } else {
      settingSet('wa_enabled', '0');
      waStop();
    }
  }
  notifyChange();
  res.json({ ok: true });
});

app.post('/api/wa/logout', async (req, res) => {
  try {
    if (waClient) {
      try { await waClient.logout(); } catch (e) {}
      try { waClient.end(); } catch (e) {}
      waClient = null;
    }
  } catch (e) {}
  try {
    if (fs.existsSync(WA_SESSION_DIR)) {
      fs.rmSync(WA_SESSION_DIR, { recursive: true, force: true });
    }
  } catch (e) {
    return res.status(500).json({ error: 'Gagal menghapus sesi: ' + e.message });
  }
  settingSet('wa_enabled', '1');
  waStart();
  notifyChange();
  res.json({ ok: true, pesan: 'Sesi WhatsApp direset. Pindai QR baru untuk menghubungkan.' });
});

app.get('/api/wa/qr', async (req, res) => {
  if (!waState.qr) return res.json({ qr: null });
  if (!QR_LIB) return res.json({ qr: waState.qr, plain: true });
  try {
    const dataUrl = await QR_LIB.toDataURL(waState.qr);
    res.json({ qr: dataUrl });
  } catch (e) {
    res.json({ qr: null });
  }
});

app.post('/api/wa/test', async (req, res) => {
  const { nomor, pesan } = req.body || {};
  const n = String(nomor || '').trim();
  if (!n) return res.status(400).json({ error: 'Nomor tujuan belum diisi' });
  const teks = String(pesan || '').trim() || 'Pesan uji coba dari aplikasi sumbangan sekolah.';
  const hasil = await kirimWa(n, teks);
  if (!hasil.ok) return res.status(400).json({ error: hasil.error });
  res.json({ ok: true });
});

// ---------- WA blast pengingat ----------
const POLA_BULAN = /^\d{4}-\d{2}$/;

function kandidatPengingat({ bulan, kelas, jenis_id, hanyaBelumBayar = true }) {
  const b = normalisasiBulan(bulan);
  if (!b || !POLA_BULAN.test(b)) return null;
  const akhirBulan = new Date(Number(b.slice(0, 4)), Number(b.slice(5, 7)), 0).getDate();
  const dari = `${b}-01`;
  const sampai = `${b}-${String(akhirBulan).padStart(2, '0')}`;

  const conds = ["IFNULL(TRIM(s.no_hp), '') != ''"];
  const args = [];
  if (kelas) { conds.push('s.kelas = ?'); args.push(String(kelas)); }
  if (hanyaBelumBayar) {
    conds.push(`NOT EXISTS (
      SELECT 1 FROM pembayaran p
      WHERE p.siswa_id = s.id AND p.tanggal >= ? AND p.tanggal <= ?
        AND (? = 0 OR p.jenis_id = ?)
    )`);
    args.push(dari, sampai, Number(jenis_id) || 0, Number(jenis_id) || 0);
  }

  const rows = db.prepare(`
    SELECT s.id, s.nis, s.nama, s.kelas, s.no_hp
    FROM siswa s
    WHERE ${conds.join(' AND ')}
    ORDER BY s.kelas, s.nama
  `).all(...args);
  return { bulan: b, dari, sampai, rows };
}

function pesanPengingatDefault(row, bulan) {
  const sekolah = settingGet('nama_sekolah') || 'Sekolah';
  return [
    `Assalamu'alaikum Bapak/Ibu wali dari *${row.nama}*${row.kelas ? ' (' + row.kelas + ')' : ''},`,
    '',
    `Mengingat bahwa sumbangan bulan *${labelBulan(bulan)}* di *${sekolah}* belum tercatat.`,
    'Mohon konfirmasi bila sudah transfer, atau informasikan bila akan terlambat.',
    '',
    'Terima kasih atas partisipasi Bapak/Ibu.'
  ].join('\n');
}

let pengingatJob = {
  jalan: false, selesai: false, bulan: null, total: 0,
  terkirim: 0, gagal: 0, mulai: null, detail: []
};

app.get('/api/wa/pengingat/kandidat', (req, res) => {
  const hasil = kandidatPengingat({
    bulan: req.query.bulan,
    kelas: req.query.kelas || null,
    jenis_id: req.query.jenis || null,
    hanyaBelumBayar: req.query.semua !== '1'
  });
  if (!hasil) return res.status(400).json({ error: 'Bulan tidak valid (format YYYY-MM)' });
  res.json({
    bulan: hasil.bulan,
    periode: `${hasil.dari} s/d ${hasil.sampai}`,
    jumlah: hasil.rows.length,
    rows: hasil.rows
  });
});

app.get('/api/wa/pengingat/status', (req, res) => {
  res.json(pengingatJob);
});

app.get('/api/wa/pengingat/riwayat', (req, res) => {
  const limit = Math.min(Math.max(Number(req.query.limit) || 100, 1), 500);
  const bulan = String(req.query.bulan || '').trim();
  const conds = [];
  const args = [];
  if (POLA_BULAN.test(bulan)) { conds.push('g.bulan = ?'); args.push(bulan); }
  const where = conds.length ? 'WHERE ' + conds.join(' AND ') : '';
  const rows = db.prepare(`
    SELECT g.*, s.nama, s.kelas, s.no_hp
    FROM pengingat g
    LEFT JOIN siswa s ON s.id = g.siswa_id
    ${where}
    ORDER BY g.id DESC
    LIMIT ?
  `).all(...args, limit);
  res.json(rows);
});

app.post('/api/wa/pengingat', async (req, res) => {
  if (!WA_LIB) return res.status(400).json({ error: 'Library WhatsApp tidak tersedia' });
  if (!waState.ready) return res.status(400).json({ error: 'WhatsApp belum terhubung. Hubungkan lewat menu Pengaturan.' });
  if (pengingatJob.jalan) return res.status(409).json({ error: 'Pengiriman pengingat sedang berjalan.' });

  const body = req.body || {};
  const hasil = kandidatPengingat({
    bulan: body.bulan,
    kelas: body.kelas || null,
    jenis_id: body.jenis || null,
    hanyaBelumBayar: body.semua !== true
  });
  if (!hasil) return res.status(400).json({ error: 'Bulan tidak valid (format YYYY-MM)' });

  let kandidat = hasil.rows;
  if (Array.isArray(body.ids) && body.ids.length) {
    const dipilih = body.ids.map(Number);
    kandidat = kandidat.filter(r => dipilih.includes(Number(r.id)));
  }
  const batas = Math.min(Math.max(Number(body.batas) || Number(settingGet('wa_pengingat_batas')) || 0, 0), 500);
  if (batas) kandidat = kandidat.slice(0, batas);
  if (!kandidat.length) {
    return res.status(400).json({ error: 'Tidak ada penerima yang cocok. Periksa bulan, kelas, dan filter.' });
  }

  const jedaMs = Math.min(Math.max(Number(body.jeda) || Number(settingGet('wa_pengingat_jeda')) || 2500, 800), 60000);
  const pesanKustom = String(body.pesan || settingGet('wa_pesan_pengingat') || '').trim();

  pengingatJob = {
    jalan: true, selesai: false, bulan: hasil.bulan, total: kandidat.length,
    terkirim: 0, gagal: 0, mulai: new Date().toISOString(), detail: []
  };

  const ins = db.prepare(
    'INSERT INTO pengingat (siswa_id, bulan, periode, pesan, status, hasil, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)'
  );
  const createdAt = new Date().toISOString();

  for (let i = 0; i < kandidat.length; i++) {
    if (!pengingatJob.jalan) break;
    const row = kandidat[i];
    const teks = pesanKustom || pesanPengingatDefault(row, hasil.bulan);
    const hasilKirim = await kirimWa(row.no_hp, teks);
    const status = hasilKirim.ok ? 'terkirim' : 'gagal';
    const catatan = hasilKirim.ok ? 'ok' : (hasilKirim.error || 'gagal');
    try {
      ins.run(row.id, hasil.bulan, `${hasil.dari} s/d ${hasil.sampai}`, teks, status, catatan, createdAt);
    } catch (e) {
      console.log('[WA pengingat] gagal menyimpan riwayat:', e.message);
    }
    if (hasilKirim.ok) pengingatJob.terkirim++;
    else pengingatJob.gagal++;
    pengingatJob.detail.push({
      nama: row.nama, kelas: row.kelas, no_hp: row.no_hp,
      status, error: hasilKirim.ok ? null : catatan
    });
    broadcast({ type: 'pengingat' });
    if (i < kandidat.length - 1) await new Promise(r => setTimeout(r, jedaMs));
  }

  pengingatJob.jalan = false;
  pengingatJob.selesai = true;
  broadcast({ type: 'pengingat' });
  res.json({ ok: true, total: pengingatJob.terkirim + pengingatJob.gagal });
});

app.post('/api/wa/pengingat/stop', (req, res) => {
  if (!pengingatJob.jalan) return res.status(400).json({ error: 'Tidak ada pengiriman yang sedang berjalan' });
  pengingatJob.jalan = false;
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`GotongRoyong berjalan di http://localhost:${PORT}`);
  console.log(`Halaman publik  : http://localhost:${PORT}/`);
  console.log(`Halaman admin   : http://localhost:${PORT}/admin.html`);
  console.log(`Password admin  : ${ADMIN_PASSWORD}`);
});
