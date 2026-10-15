#!/usr/bin/env node
/**
 * CEMIG auto-register moduleId audit
 * ==================================
 *
 * For every meter in the CEMIG meter list, reads the auto-registration records
 * (the "Registro de dispositivo" screen = MySQL table `ops_auto_register`, column
 * `register_msg` = "Conteúdo" JSON), orders them by `register_time` ascending and
 * tracks the `moduleId` inside the JSON:
 *
 *   - moduleId changes over time        -> CONFLICT   (meter is skipped from then on)
 *   - empty / invalid JSON / no moduleId -> ANOMALY    (counted, meter keeps being analysed)
 *   - meter has no records at all        -> NO_RECORDS (anomaly at meter level)
 *
 * The run is resumable: results go to state/results.jsonl, progress to
 * state/progress.json. Ctrl+C finishes the current batch, saves, builds the
 * dashboard and exits. Run again to continue from where it stopped.
 *
 * Commands
 *   node audit.mjs run    [--source mysql|api] [--db <ami schema>] [--meters mysql|api|file]
 *                         [--meters-file list.txt] [--batch 200] [--limit N] [--no-build]
 *                         [--collections ops_auto_register_202608,...] [--from 'YYYY-MM-DD'] [--to 'YYYY-MM-DD']
 *                         [--mongo-uri mongodb://host:27017/?directConnection=true] [--mongo-db ami]
 *
 * NOTE: ops_auto_register is a SPLIT table (registry db_split_table_def_n, split_uom 2 = monthly). Its consumers
 *       all read through mongoTemplate, so the rows live in MongoDB collections ops_auto_register_yyyyMM;
 *       the MySQL table of the same name is an empty template (confirmed empty in prod on 2026-10-15).
 *   node audit.mjs build  [--dashboard-max-ok N]      rebuild dashboard/ + output/*.xlsx from state
 *   node audit.mjs status                               show progress
 *   node audit.mjs probe  [--meter PMD268024393]        test MySQL + Mongo connectivity, show collections,
 *                                                       document shape and one meter's records (run this first)
 *   node audit.mjs reset  [--keep-meters]               delete state (asks for --yes)
 *
 * Sources
 *   --source mongo (default): MongoDB db 'ami' (MONGO_URI / --mongo-uri, --mongo-db), monthly collections
 *                             ops_auto_register_yyyyMM resolved from the server (--collections a,b overrides)
 *   --source mysql          : MySQL ops_auto_register_yyyyMM tables (only if a deployment keeps them in MySQL;
 *                             schema 'ami' on the cemig-ext-data server, --db overrides)
 *   --source api            : same endpoint the screen uses, GET /ami/collection/V1/register/result
 *                             (slow: one PowerShell request per meter page; logs in with the
 *                             shared SanPlat account, which invalidates other sessions)
 *   --meters mysql (default): archive_meter in the AMI schema (delete_flag = 0)
 *   --meters api            : MDC /ami/archives/V1/meters, paginated (same as fetch-meters-mdc.js; logs in)
 *   --meters file           : --meters-file <txt, one meter per line>
 *
 * Reuses ../cemig-ext-data (auth.js, config.js, .env) for credentials and transport.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync, appendFileSync, renameSync, unlinkSync, rmSync, copyFileSync, createReadStream } from 'node:fs';
import { createInterface } from 'node:readline';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import mysql from 'mysql2/promise';
import { MongoClient, ObjectId } from 'mongodb';
import ExcelJS from 'exceljs';

// ───────────────────────────────────────────────────────────── paths & constants
const ROOT = dirname(fileURLToPath(import.meta.url));
const EXT_DATA_DIR = resolve(ROOT, '..', 'cemig-ext-data');
const STATE_DIR = join(ROOT, 'state');
const DASH_DIR = join(ROOT, 'dashboard');
const VENDOR_DIR = join(DASH_DIR, 'vendor');
const OUTPUT_DIR = join(ROOT, 'output');
const F = {
  meters: join(STATE_DIR, 'meters.jsonl'),
  metersProgress: join(STATE_DIR, 'meters-progress.json'),
  progress: join(STATE_DIR, 'progress.json'),
  results: join(STATE_DIR, 'results.jsonl'),
  log: join(STATE_DIR, 'run.log'),
  dataJs: join(DASH_DIR, 'data.js'),
  indexHtml: join(DASH_DIR, 'index.html'),
  xlsx: join(OUTPUT_DIR, 'auto-register-audit.xlsx'),
  summaryJson: join(OUTPUT_DIR, 'summary.json'),
};

const STATUS = ['conflict', 'ok', 'ok_anomalies', 'no_module_id', 'no_records'];
const STATUS_LABEL = {
  conflict: 'Conflito (moduleId mudou)',
  ok: 'Consistente',
  ok_anomalies: 'Consistente com anomalias',
  no_module_id: 'Sem moduleId em nenhum registro',
  no_records: 'Sem registros',
};
const ANOMALY_TYPES = ['empty_content', 'invalid_json', 'missing_module_id', 'no_records'];
const ANOMALY_LABEL = {
  empty_content: 'Conteúdo vazio',
  invalid_json: 'Conteúdo não é JSON',
  missing_module_id: 'JSON sem moduleId',
  no_records: 'Medidor sem registros',
};
const REG_STATUS_LABEL = { 0: 'Não processado', 1: 'Sucesso', 2: 'Falha', 3: 'Inválido' };
const DEVICE_TYPE_LABEL = { '01': 'DCU', '02': 'Medidor', '05': 'Caixa de medição' };

// same server/user as cemig-ext-data's MYSQL_CONFIG (db nms); the AMI schema is 'ami' (see cemig-ext-data/src/fetch-archive-poc.mjs)
const DEFAULT_AMI_SCHEMA = 'ami';

// Where the screen's rows really live: MongoDB (db `ami`), monthly collections ops_auto_register_yyyyMM,
// per the split registry db_split_table_def_n (its consumers all go through mongoTemplate).
// CEMIG Mongo replica set: 10.30.104.31 (primary, same subnet as the MySQL server, reachable from nansenu01) and
// 10.30.132.30 (secondary; what Projects/mongodb-loadprofile-jump used). No credentials; direct connection; secondaryPreferred reads.
const DEFAULT_MONGO_HOSTS = ['10.30.104.31:27017', '10.30.132.30:27017'];   // primary first: it is the one reachable from the CEMIG-side machines (confirmed 2026-10-15)
const DEFAULT_MONGO_DB = process.env.MONGO_DB || 'ami';
const DEFAULT_MONGO_URI = process.env.MONGO_URI || `mongodb://${DEFAULT_MONGO_HOSTS[0]}/${DEFAULT_MONGO_DB}?directConnection=true`;
const MONGO_CLIENT_OPTIONS = { directConnection: true, readPreference: 'secondaryPreferred', serverSelectionTimeoutMS: 20000 };

const DEFAULTS = {
  source: 'mongo',
  meters: 'mysql',
  batch: 200,
  pageSize: 200,         // MDC meters endpoint
  apiPageSize: 500,      // register/result endpoint
  maxSamples: 20,        // anomaly samples stored per meter
  maxTimeline: 200,      // timeline rows stored for conflict meters
  dashboardMaxOk: Infinity,
  maxConsecutiveErrors: 7,
};

// ───────────────────────────────────────────────────────────── small helpers
function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (key.startsWith('no-')) args[key.slice(3)] = false;
      else if (next !== undefined && !next.startsWith('--')) { args[key] = next; i++; }
      else args[key] = true;
    } else args._.push(a);
  }
  return args;
}

function ensureDirs() { for (const d of [STATE_DIR, DASH_DIR, VENDOR_DIR, OUTPUT_DIR]) mkdirSync(d, { recursive: true }); }

function readJson(file, fallback) {
  if (!existsSync(file)) return fallback;
  try { return JSON.parse(readFileSync(file, 'utf8')); } catch { return fallback; }
}

function writeJsonAtomic(file, obj) {
  const tmp = file + '.tmp';
  writeFileSync(tmp, JSON.stringify(obj, null, 2), 'utf8');
  renameSync(tmp, file);
}

const nowIso = () => new Date().toISOString();
const pad = (n) => String(n).padStart(2, '0');
function localStamp(d = new Date()) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function log(msg) {
  const line = `[${localStamp()}] ${msg}`;
  console.log(line);
  try { appendFileSync(F.log, line + '\n', 'utf8'); } catch { /* ignore */ }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const chunk = (arr, n) => { const out = []; for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n)); return out; };
const monthOf = (ts) => (ts ? String(ts).slice(0, 7) : null);
const trimStr = (v) => (v === undefined || v === null ? '' : String(v).trim());

/** Normalises any date representation to 'YYYY-MM-DD HH:mm:ss' (local wall-clock). */
function normTime(v) {
  if (v === undefined || v === null || v === '') return null;
  if (v instanceof Date) return localStamp(v);
  const s = String(v).trim().replace('T', ' ');
  return s.length > 19 ? s.slice(0, 19) : s;
}

async function* readJsonl(file) {
  if (!existsSync(file)) return;
  const rl = createInterface({ input: createReadStream(file, 'utf8'), crlfDelay: Infinity });
  for await (const line of rl) {
    const t = line.trim();
    if (!t) continue;
    try { yield JSON.parse(t); } catch { /* skip corrupt line (interrupted write) */ }
  }
}

// ───────────────────────────────────────────────────────────── cemig-ext-data bridge
let ext = null;
async function loadExt() {
  if (ext) return ext;
  const authPath = join(EXT_DATA_DIR, 'src', 'auth.js');
  const configPath = join(EXT_DATA_DIR, 'src', 'config.js');
  if (!existsSync(authPath) || !existsSync(configPath)) {
    throw new Error(`cemig-ext-data not found at ${EXT_DATA_DIR} (needs src/auth.js, src/config.js and .env)`);
  }
  const [auth, config] = await Promise.all([import(pathToFileURL(authPath).href), import(pathToFileURL(configPath).href)]);
  ext = { auth, config };
  return ext;
}

// ───────────────────────────────────────────────────────────── MySQL
async function openMysql(database) {
  const { config } = await loadExt();
  const cfg = { ...config.MYSQL_CONFIG, dateStrings: true, connectTimeout: 30000 };
  if (database) cfg.database = database;
  return mysql.createConnection(cfg);
}

/** Finds the schema that holds ops_auto_register on the configured MySQL server. */
async function detectAmiSchema(explicit) {
  if (explicit) return explicit;
  if (process.env.AMI_MYSQL_DATABASE) return process.env.AMI_MYSQL_DATABASE;
  const conn = await openMysql();
  try {
    const [rows] = await conn.query(
      "SELECT table_schema AS s, SUM(table_rows) AS n FROM information_schema.tables WHERE table_name LIKE 'ops\\_auto\\_register%' GROUP BY table_schema ORDER BY (table_schema = ?) DESC, n DESC", [DEFAULT_AMI_SCHEMA]
    );
    if (!rows.length) throw new Error('no schema on this MySQL server has table ops_auto_register; pass --db <schema> or set AMI_MYSQL_DATABASE');
    if (rows.length > 1) log(`several schemas have ops_auto_register: ${rows.map((r) => `${r.s} (~${r.n} rows)`).join(', ')} — using ${rows[0].s}; override with --db`);
    return rows[0].s;
  } finally { await conn.end(); }
}

/** Reusable MySQL session with automatic reconnect + retry. */
function mysqlSession(database) {
  let conn = null;
  return {
    async query(sql, params, attempts = 4) {
      let lastErr;
      for (let i = 1; i <= attempts; i++) {
        try {
          if (!conn) conn = await openMysql(database);
          const [rows] = await conn.query(sql, params);
          return rows;
        } catch (err) {
          lastErr = err;
          log(`mysql error (attempt ${i}/${attempts}): ${err.message}`);
          try { await conn?.end(); } catch { /* ignore */ }
          conn = null;
          if (i < attempts) await sleep(Math.min(60000, 5000 * i));
        }
      }
      throw lastErr;
    },
    async close() { try { await conn?.end(); } catch { /* ignore */ } conn = null; },
  };
}

// ───────────────────────────────────────────────────────────── meter list
const METER_KEYS = ['meterNo', 'deviceNo', 'meter_no', 'device_no'];
const META_KEYS = ['pocNo', 'dcuNo', 'meterStatus', 'commType', 'meterTypeName', 'orgNo', 'runTime'];

function pickMeter(rec) {
  const key = METER_KEYS.find((k) => rec[k] !== undefined && rec[k] !== null && String(rec[k]).trim() !== '');
  if (!key) return null;
  const out = { meterNo: String(rec[key]).trim() };
  for (const k of META_KEYS) if (rec[k] !== undefined && rec[k] !== null && rec[k] !== '') out[k] = rec[k];
  return out;
}

async function fetchMetersApi(opts) {
  const { auth, config } = await loadExt();
  const progress = readJson(F.metersProgress, { lastPage: 0, totalPages: null, totalCount: null, source: 'api' });
  if (progress.totalPages && progress.lastPage >= progress.totalPages) { log('meter list already complete (api)'); return; }
  log('logging in to MDC for the meter list (this invalidates other sessions of the same account)...');
  const token = await auth.cemigLoginMdc();
  let page = progress.lastPage + 1;
  if (page === 1 && existsSync(F.meters)) unlinkSync(F.meters);
  let errors = 0;
  for (;;) {
    if (progress.totalPages && page > progress.totalPages) break;
    const params = new URLSearchParams({ orgNo: '1', orgName: 'CEMIG', page: String(page), size: String(opts.pageSize) });
    const url = `${config.CEMIG_MDC_URL}/ami/archives/V1/meters?${params}`;
    let data;
    try { data = await auth.cemigRequest(token, url, { authMode: 'cookie' }); } catch (err) {
      errors++; log(`meters page ${page}: error ${errors}/${opts.maxConsecutiveErrors} — ${err.message || err}`);
      if (errors >= opts.maxConsecutiveErrors) throw new Error('too many consecutive errors fetching the meter list');
      await sleep(5000); continue;
    }
    errors = 0;
    const result = data.result || {};
    const content = result.content || [];
    if (!progress.totalPages) { progress.totalPages = result.pageCount; progress.totalCount = result.count; log(`meter list: ${progress.totalCount} meters in ${progress.totalPages} pages`); }
    if (!content.length) { log(`meters page ${page}: empty — end`); progress.totalPages = page - 1; writeJsonAtomic(F.metersProgress, progress); break; }
    if (page === 1) log(`meter record keys: ${Object.keys(content[0]).join(', ')}`);
    const lines = content.map(pickMeter).filter(Boolean).map((m) => JSON.stringify(m)).join('\n') + '\n';
    appendFileSync(F.meters, lines, 'utf8');
    progress.lastPage = page; writeJsonAtomic(F.metersProgress, progress);
    log(`meters page ${page}/${progress.totalPages}: ${content.length} records`);
    page++;
  }
}

async function fetchMetersMysql(schema) {
  const progress = readJson(F.metersProgress, { lastPage: 0, totalPages: null, totalCount: null, source: 'mysql' });
  if (progress.done) { log('meter list already complete (mysql)'); return; }
  const db = mysqlSession(schema);
  try {
    if (existsSync(F.meters)) unlinkSync(F.meters);
    const PAGE = 50000; let offset = 0; let total = 0;
    for (;;) {
      const rows = await db.query(
        `SELECT meter_no AS meterNo, dcu_no AS dcuNo, meter_status AS meterStatus, comm_type AS commType, org_no AS orgNo
           FROM archive_meter WHERE delete_flag = 0 ORDER BY meter_no LIMIT ${PAGE} OFFSET ${offset}`
      );
      if (!rows.length) break;
      appendFileSync(F.meters, rows.map((r) => JSON.stringify(r)).join('\n') + '\n', 'utf8');
      total += rows.length; offset += rows.length;
      log(`archive_meter: ${total} meters so far`);
      if (rows.length < PAGE) break;
    }
    progress.done = true; progress.totalCount = total; writeJsonAtomic(F.metersProgress, progress);
  } finally { await db.close(); }
}

function importMetersFile(file) {
  if (!file || !existsSync(file)) throw new Error(`--meters-file not found: ${file}`);
  const lines = readFileSync(file, 'utf8').split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith('#'));
  writeFileSync(F.meters, lines.map((m) => JSON.stringify({ meterNo: m })).join('\n') + '\n', 'utf8');
  writeJsonAtomic(F.metersProgress, { done: true, totalCount: lines.length, source: 'file', file });
  log(`meter list imported from ${file}: ${lines.length} meters`);
}

async function loadMeterList() {
  const meters = []; const seen = new Set();
  for await (const m of readJsonl(F.meters)) {
    if (!m.meterNo || seen.has(m.meterNo)) continue;
    seen.add(m.meterNo); meters.push(m);
  }
  return meters;
}

// ───────────────────────────────────────────────────────────── register records
const REG_COLUMNS = `id, device_no AS deviceNo, device_category AS deviceCategory, register_time AS registerTime,
  register_status AS registerStatus, register_result AS registerResult, register_msg AS registerMsg,
  upper_device_type AS upperDeviceType, upper_device_no AS upperDeviceNo, ip, port, mac_address AS macAddress`;
const BASE_TABLE = 'ops_auto_register';

/**
 * ops_auto_register is a SPLIT table in SanPlat (db_split_table_def_n: split_uom 2 = one physical table per
 * month, named ops_auto_register_yyyyMM; db_split_table_detail_n lists them with their time window).
 * The base table is only the template and stays empty. Resolution order:
 *   1. --tables a,b,c (explicit)      2. db_split_table_detail_n rows for def_name = base
 *   3. SHOW TABLES LIKE 'ops_auto_register_%'      4. the base table itself, if it has rows
 * Returns [{ name, start, end }] in chronological order (start/end null when unknown).
 */
async function resolveRegisterTables(db, { explicit = null, from = null, to = null } = {}) {
  let tables = [];
  if (explicit) tables = String(explicit).split(',').map((t) => ({ name: t.trim(), start: null, end: null })).filter((t) => t.name);
  if (!tables.length) {
    try {
      const rows = await db.query(
        'SELECT table_name AS name, start_time AS start, end_time AS end FROM db_split_table_detail_n WHERE def_name = ? ORDER BY start_time, table_name', [BASE_TABLE], 1
      );
      tables = rows.map((r) => ({ name: r.name, start: normTime(r.start), end: normTime(r.end) }));
      if (tables.length) log(`split registry: ${tables.length} physical tables for ${BASE_TABLE} (${tables[0].name} .. ${tables[tables.length - 1].name})`);
    } catch (err) { log(`split registry not readable (${err.message}) — falling back to SHOW TABLES`); }
  }
  if (!tables.length) {
    const rows = await db.query(`SHOW TABLES LIKE '${BASE_TABLE}\\_%'`);
    tables = rows.map((r) => ({ name: String(Object.values(r)[0]), start: null, end: null })).filter((t) => /^ops_auto_register_\d{4,8}$/.test(t.name)).sort((a, b) => a.name.localeCompare(b.name));
    if (tables.length) log(`SHOW TABLES: ${tables.length} physical tables for ${BASE_TABLE}`);
  }
  // keep only tables that actually exist (registry rows may reference dropped/expired tables)
  const existing = new Set((await db.query(`SHOW TABLES LIKE '${BASE_TABLE}%'`)).map((r) => String(Object.values(r)[0])));
  const missing = tables.filter((t) => !existing.has(t.name)).map((t) => t.name);
  if (missing.length) log(`ignoring ${missing.length} registry tables that do not exist: ${missing.slice(0, 5).join(', ')}${missing.length > 5 ? ', …' : ''}`);
  tables = tables.filter((t) => existing.has(t.name));
  // time-window pruning (only when the registry gave us the window)
  if (from || to) tables = tables.filter((t) => (!t.end || !from || t.end >= from) && (!t.start || !to || t.start <= to));
  // the base table itself, only if it holds rows
  if (existing.has(BASE_TABLE)) {
    const probe = await db.query(`SELECT 1 FROM ${BASE_TABLE} LIMIT 1`);
    if (probe.length) { log(`base table ${BASE_TABLE} is not empty — included`); tables.push({ name: BASE_TABLE, start: null, end: null }); }
  }
  if (!tables.length) throw new Error(`no physical table found for ${BASE_TABLE} (registry, SHOW TABLES and base table are all empty) — check the schema (--db) or pass --tables`);
  log(`register tables: ${tables.map((t) => t.name).join(', ')}`);
  return tables;
}

async function fetchRecordsMysql(db, meterNos, tables) {
  // one round trip per batch: UNION ALL over the monthly tables, then a client-side sort
  const sql = tables.map((t) => `SELECT ${REG_COLUMNS} FROM ${t.name} WHERE device_no IN (?)`).join('\nUNION ALL\n');
  const rows = await db.query(sql, tables.map(() => meterNos));
  const byMeter = new Map();
  for (const r of rows) {
    r.registerTime = normTime(r.registerTime);
    const list = byMeter.get(r.deviceNo) || []; list.push(r); byMeter.set(r.deviceNo, list);
  }
  for (const list of byMeter.values()) list.sort((a, b) => (a.registerTime || '').localeCompare(b.registerTime || '') || String(a.id).localeCompare(String(b.id)));
  return byMeter;
}

// ───────────────────────────────────────────────────────────── register records (MongoDB — the real source)
/**
 * Opens the CEMIG Mongo. With no explicit --mongo-uri / MONGO_URI it tries each default host in turn
 * (primary first, then secondary) so a machine that only routes to one subnet still gets through.
 */
async function mongoSession(uri = null, dbName = DEFAULT_MONGO_DB) {
  const explicit = uri || process.env.MONGO_URI || null;
  const candidates = explicit ? [explicit] : DEFAULT_MONGO_HOSTS.map((h) => `mongodb://${h}/${dbName}?directConnection=true`);
  const failures = [];
  for (const cand of candidates) {
    const masked = cand.replace(/\/\/[^@]*@/, '//***@');
    const client = new MongoClient(cand, MONGO_CLIENT_OPTIONS);
    log(`mongo: connecting to ${masked} (db ${dbName}) …`);
    try {
      await client.connect();
      await client.db(dbName).command({ ping: 1 });
      log(`mongo: connected to ${masked} db ${dbName}`);
      return { client, db: client.db(dbName), dbName, uri: masked, async close() { try { await client.close(); } catch { /* ignore */ } } };
    } catch (err) {
      // surface the per-server reason (ECONNREFUSED / ETIMEDOUT / auth) instead of the generic "server selection timed out"
      const servers = err?.reason?.servers ? [...err.reason.servers.values()].map((s) => `${s.address}: ${s.error?.message || s.type || 'unknown'}`) : [];
      const why = `${masked}: ${err.message}${servers.length ? ` — ${servers.join('; ')}` : ''}`;
      failures.push(why); log(`mongo: ${why}`);
      try { await client.close(); } catch { /* ignore */ }
    }
  }
  throw new Error(`cannot reach MongoDB (${failures.length} host${failures.length > 1 ? 's' : ''} tried). Check the CEMIG VPN/route (Test-NetConnection <host> -Port 27017) or pass --mongo-uri / set MONGO_URI. Details: ${failures.join(' | ')}`);
}

/** Time window of a split collection from its suffix: _yyyyMMdd (day), _yyyyMM (month), _yyyy (year). */
function suffixWindow(name) {
  const m = name.match(/_(\d{4})(\d{2})?(\d{2})?$/);
  if (!m) return { start: null, end: null };
  const y = +m[1], mo = m[2] ? +m[2] : null, d = m[3] ? +m[3] : null;
  const p = (n) => String(n).padStart(2, '0');
  if (d) return { start: `${y}-${p(mo)}-${p(d)} 00:00:00`, end: `${y}-${p(mo)}-${p(d)} 23:59:59` };
  if (mo) { const last = new Date(y, mo, 0).getDate(); return { start: `${y}-${p(mo)}-01 00:00:00`, end: `${y}-${p(mo)}-${p(last)} 23:59:59` }; }
  return { start: `${y}-01-01 00:00:00`, end: `${y}-12-31 23:59:59` };
}

/**
 * Lists the physical Mongo collections of ops_auto_register: ops_auto_register_<suffix> in chronological order,
 * plus the base collection if it holds documents. --collections (explicit) overrides; --from/--to prune.
 */
async function resolveRegisterCollections(mongo, { explicit = null, from = null, to = null } = {}) {
  let names;
  if (explicit) names = String(explicit).split(',').map((s) => s.trim()).filter(Boolean);
  else {
    const all = (await mongo.db.listCollections({ name: { $regex: `^${BASE_TABLE}` } }, { nameOnly: true }).toArray()).map((c) => c.name);
    names = all.filter((n) => new RegExp(`^${BASE_TABLE}_\\d{4,8}$`).test(n)).sort((a, b) => a.localeCompare(b));
    if (all.includes(BASE_TABLE)) {
      const n = await mongo.db.collection(BASE_TABLE).estimatedDocumentCount();
      if (n > 0) { log(`base collection ${BASE_TABLE} holds ~${n} documents — included`); names.push(BASE_TABLE); }
    }
  }
  let cols = names.map((name) => ({ name, ...suffixWindow(name) }));
  if (from || to) cols = cols.filter((c) => (!c.end || !from || c.end >= from) && (!c.start || !to || c.start <= to));
  if (!cols.length) throw new Error(`no collection ${BASE_TABLE}_* found in Mongo db ${mongo.dbName} (check --mongo-uri / --mongo-db, or pass --collections)`);
  log(`register collections (${cols.length}): ${cols.map((c) => c.name).join(', ')}`);
  return cols;
}

const MONGO_INFO_CACHE = new Map();
/**
 * Samples one document of a collection to learn its field naming (snake_case @Field names are expected:
 * device_no, register_time ...; camelCase is tolerated) and whether register_time is a BSON Date or a string.
 * Needed so filters ($in on the meter field, time windows) hit the real field with the right BSON type.
 */
async function mongoCollectionInfo(mongo, name) {
  if (MONGO_INFO_CACHE.has(name)) return MONGO_INFO_CACHE.get(name);
  const doc = await mongo.db.collection(name).findOne({}, { projection: { device_no: 1, deviceNo: 1, register_time: 1, registerTime: 1, device_category: 1, deviceCategory: 1 } });
  const camel = !!doc && doc.deviceNo !== undefined && doc.device_no === undefined;
  const info = {
    sampled: !!doc, meterField: camel ? 'deviceNo' : 'device_no', timeField: camel ? 'registerTime' : 'register_time', categoryField: camel ? 'deviceCategory' : 'device_category',
    timeIsDate: !!doc && (doc.register_time ?? doc.registerTime) instanceof Date,
  };
  if (doc && !MONGO_INFO_CACHE.size) log(`mongo document shape (${name}): meter field ${info.meterField}, ${info.timeField} stored as ${info.timeIsDate ? 'Date' : typeof (doc.register_time ?? doc.registerTime)}`);
  MONGO_INFO_CACHE.set(name, info);
  return info;
}
const mongoMeterField = async (mongo, name) => (await mongoCollectionInfo(mongo, name)).meterField;
/** Builds a {$gte,$lte} bound for the time field with the BSON type the collection actually uses. */
function mongoTimeBound(info, from, to) {
  const conv = (s) => (info.timeIsDate ? new Date(String(s).replace(' ', 'T')) : String(s));
  const b = {}; if (from) b.$gte = conv(from); if (to) b.$lte = conv(to); return b;
}
const MONGO_PROJECTION = { _id: 1, device_no: 1, deviceNo: 1, device_category: 1, deviceCategory: 1, register_time: 1, registerTime: 1, register_status: 1, registerStatus: 1, register_result: 1, registerResult: 1,
  register_msg: 1, registerMsg: 1, upper_device_type: 1, upperDeviceType: 1, upper_device_no: 1, upperDeviceNo: 1, lastUpperDeviceNo: 1, ip: 1, port: 1, mac_address: 1, macAddress: 1, wiring_phase: 1, wiringPhase: 1 };

function normalizeMongoRow(d) {
  const g = (...ks) => { for (const k of ks) if (d[k] !== undefined && d[k] !== null) return d[k]; return null; };
  return {
    id: d._id instanceof ObjectId ? d._id.toHexString() : String(d._id), deviceNo: trimStr(g('device_no', 'deviceNo')), deviceCategory: g('device_category', 'deviceCategory'),
    registerTime: normTime(g('register_time', 'registerTime')), registerStatus: g('register_status', 'registerStatus'), registerResult: g('register_result', 'registerResult'),
    registerMsg: g('register_msg', 'registerMsg'), upperDeviceType: g('upper_device_type', 'upperDeviceType'), upperDeviceNo: g('upper_device_no', 'upperDeviceNo', 'lastUpperDeviceNo'),
    ip: g('ip'), port: g('port'), macAddress: g('mac_address', 'macAddress'), wiringPhase: g('wiring_phase', 'wiringPhase'),
  };
}

async function fetchRecordsMongo(mongo, meterNos, collections, attempts = 4) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      const perCol = await Promise.all(collections.map(async (c) => {
        const field = await mongoMeterField(mongo, c.name);
        return mongo.db.collection(c.name).find({ [field]: { $in: meterNos } }, { projection: MONGO_PROJECTION }).toArray();
      }));
      const byMeter = new Map();
      for (const docs of perCol) for (const d of docs) {
        const r = normalizeMongoRow(d);
        const list = byMeter.get(r.deviceNo) || []; list.push(r); byMeter.set(r.deviceNo, list);
      }
      for (const list of byMeter.values()) list.sort((a, b) => (a.registerTime || '').localeCompare(b.registerTime || '') || a.id.localeCompare(b.id));
      return byMeter;
    } catch (err) {
      lastErr = err; log(`mongo error (attempt ${i}/${attempts}): ${err.message}`);
      if (i < attempts) await sleep(Math.min(60000, 5000 * i));
    }
  }
  throw lastErr;
}

function apiSession() {
  let token = null;
  return {
    async request(url) {
      const { auth } = await loadExt();
      if (!token) { log('logging in to MDC (register/result API)...'); token = await auth.cemigLoginMdc(); }
      try { return await auth.cemigRequest(token, url, { authMode: 'cookie' }); } catch (err) {
        if (/AMI-00013|Sess/i.test(String(err.message))) { token = null; }
        throw err;
      }
    },
  };
}

async function fetchRecordsApi(api, meterNos, opts) {
  const { config } = await loadExt();
  const byMeter = new Map();
  const endTime = localStamp(new Date(Date.now() + 86400000));
  for (const meterNo of meterNos) {
    const rows = [];
    let page = 1; let errors = 0;
    for (;;) {
      const params = new URLSearchParams({ deviceNo: meterNo, startTime: '2000-01-01 00:00:00', endTime, page: String(page), size: String(opts.apiPageSize) });
      const url = `${config.CEMIG_MDC_URL}/ami/collection/V1/register/result?${params}`;
      let data;
      try { data = await api.request(url); } catch (err) {
        errors++; log(`register/result ${meterNo} p${page}: error ${errors}/${opts.maxConsecutiveErrors} — ${err.message || err}`);
        if (errors >= opts.maxConsecutiveErrors) throw new Error(`too many consecutive API errors (meter ${meterNo})`);
        await sleep(3000); continue;
      }
      errors = 0;
      const result = data.result || {};
      const content = Array.isArray(result) ? result : (result.content || []);
      for (const c of content) {
        rows.push({
          id: c.id ?? null, deviceNo: c.deviceNo ?? meterNo, deviceCategory: c.deviceCategory ?? null, registerTime: normTime(c.registerTime),
          registerStatus: c.registerStatus ?? null, registerResult: c.registerResult ?? null, registerMsg: c.registerMsg ?? null,
          upperDeviceType: c.lastUpperDeviceType ?? c.upperDeviceType ?? null, upperDeviceNo: c.lastUpperDeviceNo ?? c.upperDeviceNo ?? null,
          ip: c.ip ?? null, port: c.port ?? null, macAddress: c.macAddress ?? null,
        });
      }
      const pageCount = result.pageCount ?? (content.length < opts.apiPageSize ? page : page + 1);
      if (!content.length || page >= pageCount) break;
      page++;
    }
    rows.sort((a, b) => (a.registerTime || '').localeCompare(b.registerTime || '') || String(a.id || '').localeCompare(String(b.id || '')));
    if (rows.length) byMeter.set(meterNo, rows);
  }
  return byMeter;
}

// ───────────────────────────────────────────────────────────── analysis (the rule)
function parseContent(msg) {
  const s = trimStr(msg);
  if (!s) return { anomaly: 'empty_content', json: null };
  let json;
  try { json = JSON.parse(s); } catch { return { anomaly: 'invalid_json', json: null }; }
  if (!json || typeof json !== 'object' || Array.isArray(json)) return { anomaly: 'invalid_json', json: null };
  const moduleId = trimStr(json.moduleId);
  if (!moduleId) return { anomaly: 'missing_module_id', json };
  return { anomaly: null, json, moduleId };
}

/**
 * rows must be ordered by registerTime ascending.
 * Returns one result object per meter (what goes into results.jsonl).
 */
function analyzeMeter(meter, rows, opts) {
  const res = {
    meterNo: meter.meterNo, meta: Object.fromEntries(Object.entries(meter).filter(([k]) => k !== 'meterNo')),
    status: 'ok', records: rows.length, firstAt: null, lastAt: null, successes: 0, failures: 0, regStatus: {},
    moduleId: null, moduleSince: null, moduleCount: 0, upperDevices: [], macs: [], ips: [],
    conflict: null, rowsSkippedAfterConflict: 0,
    anomalies: { empty_content: 0, invalid_json: 0, missing_module_id: 0 }, anomalyTotal: 0, anomaliesByMonth: {}, anomalySamples: [],
    recordsByMonth: {}, timeline: null, analyzedAt: nowIso(),
  };
  if (!rows.length) {
    res.status = 'no_records'; res.anomalyTotal = 1; res.anomalies.no_records = 1;
    return res;
  }
  res.firstAt = rows[0].registerTime; res.lastAt = rows[rows.length - 1].registerTime;
  const upper = new Set(); const macs = new Set(); const ips = new Set();
  const timeline = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const m = monthOf(r.registerTime); if (m) res.recordsByMonth[m] = (res.recordsByMonth[m] || 0) + 1;
    const st = trimStr(r.registerStatus); res.regStatus[st || '?'] = (res.regStatus[st || '?'] || 0) + 1;
    if (st === '1') res.successes++; else if (st === '2') res.failures++;
    if (r.upperDeviceNo) upper.add(String(r.upperDeviceNo)); if (r.macAddress) macs.add(String(r.macAddress)); if (r.ip) ips.add(String(r.ip));
    const c = parseContent(r.registerMsg);
    const tl = [r.registerTime, st, c.moduleId || null, c.anomaly, trimStr(r.registerResult).slice(0, 120), r.upperDeviceNo || null, c.json?.dateTime || null];
    timeline.push(tl);
    if (c.anomaly) {
      res.anomalies[c.anomaly]++; res.anomalyTotal++;
      if (m) res.anomaliesByMonth[m] = (res.anomaliesByMonth[m] || 0) + 1;
      if (res.anomalySamples.length < opts.maxSamples) res.anomalySamples.push({ at: r.registerTime, type: c.anomaly, regStatus: st, result: trimStr(r.registerResult).slice(0, 200), content: trimStr(r.registerMsg).slice(0, 200) });
      continue;
    }
    if (res.moduleId === null) { res.moduleId = c.moduleId; res.moduleSince = r.registerTime; res.moduleCount = 1; continue; }
    if (c.moduleId === res.moduleId) { res.moduleCount++; continue; }
    // moduleId changed -> conflict; stop analysing this meter
    res.status = 'conflict';
    res.conflict = {
      at: r.registerTime, month: m, index: i, previousModuleId: res.moduleId, newModuleId: c.moduleId,
      previousSince: res.moduleSince, previousCount: res.moduleCount, regStatus: st, result: trimStr(r.registerResult).slice(0, 200),
      ip: r.ip || null, port: r.port ?? null, macAddress: r.macAddress || null, upperDeviceNo: r.upperDeviceNo || null,
      msgDateTime: c.json?.dateTime || null, meterNoInContent: c.json?.meterNo || null,
    };
    res.rowsSkippedAfterConflict = rows.length - i - 1;
    // keep a window of the timeline around the conflict: up to maxTimeline rows before it and 5 after
    const after = rows.slice(i + 1, i + 6).map((x) => { const cc = parseContent(x.registerMsg); return [x.registerTime, trimStr(x.registerStatus), cc.moduleId || null, cc.anomaly, trimStr(x.registerResult).slice(0, 120), x.upperDeviceNo || null, cc.json?.dateTime || null]; });
    res.timeline = timeline.slice(Math.max(0, timeline.length - opts.maxTimeline)).concat(after);
    break;
  }
  if (res.status !== 'conflict') {
    if (res.moduleId === null) res.status = 'no_module_id';
    else if (res.anomalyTotal > 0) res.status = 'ok_anomalies';
  }
  res.upperDevices = [...upper]; res.macs = [...macs]; res.ips = [...ips];
  return res;
}

// ───────────────────────────────────────────────────────────── run (resumable loop)
let stopRequested = false;
function installSignalHandlers() {
  const onSignal = (sig) => {
    if (stopRequested) { log(`${sig} again — exiting immediately`); process.exit(130); }
    stopRequested = true;
    log(`${sig} received — finishing the current batch, saving state and building the dashboard. Press again to force quit.`);
  };
  for (const s of ['SIGINT', 'SIGTERM', 'SIGBREAK']) { try { process.on(s, () => onSignal(s)); } catch { /* not supported */ } }
}

async function cmdRun(args) {
  ensureDirs();
  const opts = { ...DEFAULTS, ...args };
  opts.batch = Number(opts.batch) || DEFAULTS.batch;
  opts.limit = args.limit !== undefined ? Number(args.limit) : Infinity;
  opts.maxSamples = Number(args['max-samples']) || DEFAULTS.maxSamples;
  opts.maxTimeline = Number(args['max-timeline']) || DEFAULTS.maxTimeline;
  if (!['mongo', 'mysql', 'api'].includes(opts.source)) throw new Error(`--source must be mongo, mysql or api (got ${opts.source})`);
  if (!['api', 'mysql', 'file'].includes(opts.meters)) throw new Error(`--meters must be api, mysql or file (got ${opts.meters})`);
  installSignalHandlers();
  log(`run: source=${opts.source} meters=${opts.meters} batch=${opts.batch}${Number.isFinite(opts.limit) ? ` limit=${opts.limit}` : ''}`);

  let schema = null;
  if (opts.source === 'mysql' || opts.meters === 'mysql') { schema = await detectAmiSchema(args.db); log(`AMI MySQL schema: ${schema}`); }

  // 1. meter list (fetched once, cached in state/meters.jsonl)
  if (opts.meters === 'file') { if (!readJson(F.metersProgress, {}).done || args['meters-file']) importMetersFile(args['meters-file']); }
  else if (opts.meters === 'mysql') await fetchMetersMysql(schema);
  else await fetchMetersApi(opts);
  const meters = await loadMeterList();
  if (!meters.length) throw new Error('meter list is empty');
  log(`meter list: ${meters.length} unique meters`);

  // 2. progress
  const progress = readJson(F.progress, { nextIndex: 0, processed: 0, startedAt: nowIso(), runs: 0 });
  if (progress.metersTotal && progress.metersTotal !== meters.length) log(`note: meter list size changed (${progress.metersTotal} -> ${meters.length}); continuing by index ${progress.nextIndex}`);
  progress.metersTotal = meters.length; progress.source = opts.source; progress.metersSource = opts.meters; progress.schema = schema; progress.runs++;
  progress.updatedAt = nowIso(); delete progress.finishedAt;
  writeJsonAtomic(F.progress, progress);
  if (progress.nextIndex >= meters.length) { log('all meters already analysed — nothing to do (use `reset` to start over)'); if (opts.build !== false) await cmdBuild(args); return; }
  log(`resuming at index ${progress.nextIndex} (${progress.processed} meters already analysed)`);

  // 3. loop
  const db = opts.source === 'mysql' ? mysqlSession(schema) : null;
  const api = opts.source === 'api' ? apiSession() : null;
  const mongo = opts.source === 'mongo' ? await mongoSession(args['mongo-uri'] || null, args['mongo-db'] || DEFAULT_MONGO_DB) : null;
  const tables = db ? await resolveRegisterTables(db, { explicit: args.tables, from: args.from, to: args.to })
    : mongo ? await resolveRegisterCollections(mongo, { explicit: args.collections, from: args.from, to: args.to }) : null;
  progress.store = mongo ? `mongodb ${mongo.dbName}@${new URL(mongo.uri).host}` : db ? `mysql ${schema}` : 'mdc api';
  if (tables) progress.tables = tables.map((t) => t.name);
  writeJsonAtomic(F.progress, progress);
  const t0 = Date.now(); let doneThisRun = 0;
  try {
    while (progress.nextIndex < meters.length && doneThisRun < opts.limit && !stopRequested) {
      const size = Math.min(opts.batch, meters.length - progress.nextIndex, opts.limit - doneThisRun);
      const batch = meters.slice(progress.nextIndex, progress.nextIndex + size);
      const nos = batch.map((m) => m.meterNo);
      const byMeter = mongo ? await fetchRecordsMongo(mongo, nos, tables) : db ? await fetchRecordsMysql(db, nos, tables) : await fetchRecordsApi(api, nos, opts);
      const lines = [];
      let conflicts = 0, anomalies = 0, noRecords = 0, records = 0;
      for (const m of batch) {
        const rows = byMeter.get(m.meterNo) || [];
        const res = analyzeMeter(m, rows, opts);
        records += rows.length;
        if (res.status === 'conflict') conflicts++; if (res.status === 'no_records') noRecords++; if (res.anomalyTotal) anomalies++;
        lines.push(JSON.stringify(res));
      }
      appendFileSync(F.results, lines.join('\n') + '\n', 'utf8');     // results first ...
      progress.nextIndex += batch.length; progress.processed += batch.length; progress.updatedAt = nowIso();
      writeJsonAtomic(F.progress, progress);                           // ... then the checkpoint
      doneThisRun += batch.length;
      const elapsed = (Date.now() - t0) / 1000; const rate = doneThisRun / Math.max(1, elapsed);
      const remaining = meters.length - progress.nextIndex; const eta = rate > 0 ? Math.round(remaining / rate / 60) : null;
      log(`${progress.nextIndex}/${meters.length} (${(100 * progress.nextIndex / meters.length).toFixed(1)}%) · batch: ${records} records, ${conflicts} conflicts, ${anomalies} with anomalies, ${noRecords} without records · ${rate.toFixed(1)} m/s${eta !== null ? ` · ETA ${eta} min` : ''}`);
    }
  } finally { await db?.close(); await mongo?.close(); }

  if (progress.nextIndex >= meters.length) { progress.finishedAt = nowIso(); writeJsonAtomic(F.progress, progress); log('analysis complete'); }
  else log(stopRequested ? 'stopped by request — run again to continue' : `stopped after --limit (${doneThisRun} meters this run)`);
  if (opts.build !== false) await cmdBuild(args);
}

// ───────────────────────────────────────────────────────────── status / reset
async function cmdStatus() {
  const p = readJson(F.progress, null);
  const mp = readJson(F.metersProgress, null);
  if (!mp) console.log('meter list: not fetched yet'); else console.log(`meter list: ${mp.totalCount ?? '?'} meters (${mp.source || 'api'}${mp.done || (mp.totalPages && mp.lastPage >= mp.totalPages) ? ', complete' : `, page ${mp.lastPage}/${mp.totalPages}`})`);
  if (!p) { console.log('analysis: not started'); return; }
  console.log(`analysis: ${p.processed}/${p.metersTotal} meters (${(100 * p.processed / Math.max(1, p.metersTotal)).toFixed(1)}%) · source ${p.source} · schema ${p.schema || '-'} · runs ${p.runs}`);
  console.log(`started ${p.startedAt} · updated ${p.updatedAt}${p.finishedAt ? ` · finished ${p.finishedAt}` : ''}`);
  let n = 0; const counts = {};
  for await (const r of readJsonl(F.results)) { n++; counts[r.status] = (counts[r.status] || 0) + 1; }
  console.log(`results.jsonl: ${n} lines · ` + STATUS.map((s) => `${s}=${counts[s] || 0}`).join(' '));
}

function cmdReset(args) {
  if (!args.yes) { console.log('reset deletes state/, dashboard/ and output/. Run again with --yes (add --keep-meters to keep the meter list).'); return; }
  for (const f of [F.progress, F.results, F.log]) if (existsSync(f)) unlinkSync(f);
  if (!args['keep-meters']) for (const f of [F.meters, F.metersProgress]) if (existsSync(f)) unlinkSync(f);
  for (const d of [DASH_DIR, OUTPUT_DIR]) rmSync(d, { recursive: true, force: true });
  console.log('state cleared');
}

// ───────────────────────────────────────────────────────────── build (dashboard + xlsx)
async function loadResults() {
  const byMeter = new Map();
  for await (const r of readJsonl(F.results)) byMeter.set(r.meterNo, r);  // last line wins (re-analysed meters)
  return [...byMeter.values()];
}

function summarize(results, progress) {
  const s = {
    generatedAt: localStamp(), metersTotal: progress?.metersTotal ?? results.length, analyzed: results.length,
    pending: Math.max(0, (progress?.metersTotal ?? results.length) - results.length), finished: !!progress?.finishedAt,
    source: progress?.source || null, store: progress?.store || null, schema: progress?.schema || null, startedAt: progress?.startedAt || null, updatedAt: progress?.updatedAt || null,
    status: Object.fromEntries(STATUS.map((k) => [k, 0])), anomalies: Object.fromEntries(ANOMALY_TYPES.map((k) => [k, 0])),
    metersWithAnomalies: 0, anomalyRows: 0, recordsTotal: 0, successes: 0, failures: 0,
    conflictsByMonth: {}, anomaliesByMonth: {}, recordsByMonth: {}, conflictRegStatus: {}, months: [],
    recordBins: [[0, 0], [1, 1], [2, 5], [6, 10], [11, 50], [51, 200], [201, null]].map(([min, max]) => ({ min, max, label: max === null ? `> ${min - 1}` : min === max ? String(min) : `${min}–${max}`, count: 0 })),
  };
  const months = new Set();
  for (const r of results) {
    s.status[r.status] = (s.status[r.status] || 0) + 1;
    s.recordsTotal += r.records; s.successes += r.successes || 0; s.failures += r.failures || 0;
    for (const k of ANOMALY_TYPES) s.anomalies[k] += r.anomalies?.[k] || 0;
    if (r.anomalyTotal) { s.metersWithAnomalies++; s.anomalyRows += r.anomalyTotal; }
    for (const [m, n] of Object.entries(r.anomaliesByMonth || {})) { s.anomaliesByMonth[m] = (s.anomaliesByMonth[m] || 0) + n; months.add(m); }
    for (const [m, n] of Object.entries(r.recordsByMonth || {})) { s.recordsByMonth[m] = (s.recordsByMonth[m] || 0) + n; months.add(m); }
    if (r.conflict) { const m = r.conflict.month; if (m) { s.conflictsByMonth[m] = (s.conflictsByMonth[m] || 0) + 1; months.add(m); } const st = r.conflict.regStatus || '?'; s.conflictRegStatus[st] = (s.conflictRegStatus[st] || 0) + 1; }
    const bin = s.recordBins.find((b) => r.records >= b.min && (b.max === null || r.records <= b.max)); if (bin) bin.count++;
  }
  s.months = [...months].sort();
  return s;
}

async function copyVendor() {
  const want = [
    ['chart.umd.js', join(ROOT, 'node_modules', 'chart.js', 'dist', 'chart.umd.js'), 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js'],
    ['xlsx.full.min.js', join(ROOT, 'node_modules', 'xlsx', 'dist', 'xlsx.full.min.js'), 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'],
  ];
  const tags = [];
  for (const [name, src, cdn] of want) {
    const dst = join(VENDOR_DIR, name);
    if (existsSync(src)) { copyFileSync(src, dst); tags.push(`<script src="vendor/${name}"></script>`); }
    else if (existsSync(dst)) tags.push(`<script src="vendor/${name}"></script>`);
    else { log(`vendor ${name} not found in node_modules (run npm install) — dashboard will load it from the CDN`); tags.push(`<script src="${cdn}"></script>`); }
  }
  return tags.join('\n');
}

async function writeXlsx(results, summary) {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: F.xlsx, useStyles: true });
  const bold = { font: { bold: true } };
  const addHeader = (ws, cols) => { ws.columns = cols.map((c) => ({ header: c.h, key: c.k, width: c.w || 16 })); ws.getRow(1).font = bold.font; ws.getRow(1).commit(); };

  const sm = wb.addWorksheet('Resumo');
  addHeader(sm, [{ h: 'Indicador', k: 'k', w: 42 }, { h: 'Valor', k: 'v', w: 24 }]);
  const rowsSummary = [
    ['Gerado em', summary.generatedAt], ['Fonte', summary.source], ['Origem dos registros', summary.store], ['Schema MySQL (lista de medidores)', summary.schema], ['Análise concluída', summary.finished ? 'sim' : 'não'],
    ['Medidores na lista', summary.metersTotal], ['Medidores analisados', summary.analyzed], ['Medidores pendentes', summary.pending],
    ...STATUS.map((k) => [STATUS_LABEL[k], summary.status[k] || 0]),
    ['Medidores com anomalias', summary.metersWithAnomalies], ['Registros com anomalia', summary.anomalyRows],
    ...ANOMALY_TYPES.map((k) => [`Anomalia · ${ANOMALY_LABEL[k]}`, summary.anomalies[k] || 0]),
    ['Registros lidos', summary.recordsTotal], ['Registros com sucesso', summary.successes], ['Registros com falha', summary.failures],
  ];
  for (const [k, v] of rowsSummary) sm.addRow({ k, v }).commit();
  sm.commit();

  const cf = wb.addWorksheet('Conflitos');
  addHeader(cf, [
    { h: 'Medidor', k: 'meterNo' }, { h: 'UC', k: 'pocNo' }, { h: 'DCU (cadastro)', k: 'dcuNo' }, { h: 'Conflito em', k: 'at', w: 20 }, { h: 'moduleId anterior', k: 'prev' }, { h: 'moduleId novo', k: 'next' },
    { h: 'Anterior desde', k: 'prevSince', w: 20 }, { h: 'Registros com o anterior', k: 'prevCount' }, { h: 'Status do registro', k: 'regStatus' }, { h: 'Resultado', k: 'result', w: 40 },
    { h: 'IP', k: 'ip', w: 30 }, { h: 'Porta', k: 'port' }, { h: 'MAC', k: 'mac', w: 24 }, { h: 'Registrado via', k: 'upper' }, { h: 'dateTime no conteúdo', k: 'msgDateTime', w: 20 },
    { h: 'Registros do medidor', k: 'records' }, { h: 'Registros ignorados após conflito', k: 'skipped' }, { h: 'Anomalias antes do conflito', k: 'anomalyTotal' }, { h: 'Primeiro registro', k: 'firstAt', w: 20 }, { h: 'Último registro', k: 'lastAt', w: 20 },
  ]);
  const an = wb.addWorksheet('Anomalias');
  addHeader(an, [
    { h: 'Medidor', k: 'meterNo' }, { h: 'UC', k: 'pocNo' }, { h: 'Status do medidor', k: 'status', w: 28 }, { h: 'Registrado em', k: 'at', w: 20 }, { h: 'Tipo de anomalia', k: 'type', w: 24 },
    { h: 'Status do registro', k: 'regStatus' }, { h: 'Resultado', k: 'result', w: 40 }, { h: 'Conteúdo (início)', k: 'content', w: 60 },
  ]);
  const am = wb.addWorksheet('Anomalias por medidor');
  addHeader(am, [
    { h: 'Medidor', k: 'meterNo' }, { h: 'UC', k: 'pocNo' }, { h: 'Status do medidor', k: 'status', w: 28 }, { h: 'Registros', k: 'records' }, { h: 'Total de anomalias', k: 'anomalyTotal' },
    ...ANOMALY_TYPES.map((k) => ({ h: ANOMALY_LABEL[k], k, w: 22 })), { h: 'Amostras na aba Anomalias', k: 'samples' },
  ]);
  const me = wb.addWorksheet('Medidores');
  addHeader(me, [
    { h: 'Medidor', k: 'meterNo' }, { h: 'UC', k: 'pocNo' }, { h: 'DCU (cadastro)', k: 'dcuNo' }, { h: 'Status', k: 'status', w: 28 }, { h: 'Registros', k: 'records' }, { h: 'Primeiro registro', k: 'firstAt', w: 20 }, { h: 'Último registro', k: 'lastAt', w: 20 },
    { h: 'moduleId', k: 'moduleId' }, { h: 'moduleId desde', k: 'moduleSince', w: 20 }, { h: 'Registros com o moduleId', k: 'moduleCount' }, { h: 'Sucessos', k: 'successes' }, { h: 'Falhas', k: 'failures' },
    { h: 'Anomalias', k: 'anomalyTotal' }, { h: 'Conflito em', k: 'conflictAt', w: 20 }, { h: 'moduleId novo', k: 'conflictNew' }, { h: 'Nº DCUs/gateways distintos', k: 'upperCount' }, { h: 'Nº MACs distintos', k: 'macCount' },
  ]);

  results.sort((a, b) => a.meterNo.localeCompare(b.meterNo));
  for (const r of results) {
    const meta = r.meta || {};
    me.addRow({ meterNo: r.meterNo, pocNo: meta.pocNo ?? '', dcuNo: meta.dcuNo ?? '', status: STATUS_LABEL[r.status] || r.status, records: r.records, firstAt: r.firstAt ?? '', lastAt: r.lastAt ?? '',
      moduleId: r.moduleId ?? '', moduleSince: r.moduleSince ?? '', moduleCount: r.moduleCount, successes: r.successes, failures: r.failures, anomalyTotal: r.anomalyTotal,
      conflictAt: r.conflict?.at ?? '', conflictNew: r.conflict?.newModuleId ?? '', upperCount: r.upperDevices?.length ?? 0, macCount: r.macs?.length ?? 0 }).commit();
    if (r.conflict) {
      const c = r.conflict;
      cf.addRow({ meterNo: r.meterNo, pocNo: meta.pocNo ?? '', dcuNo: meta.dcuNo ?? '', at: c.at, prev: c.previousModuleId, next: c.newModuleId, prevSince: c.previousSince, prevCount: c.previousCount,
        regStatus: `${c.regStatus} · ${REG_STATUS_LABEL[c.regStatus] || '?'}`, result: c.result, ip: c.ip ?? '', port: c.port ?? '', mac: c.macAddress ?? '', upper: c.upperDeviceNo ?? '', msgDateTime: c.msgDateTime ?? '',
        records: r.records, skipped: r.rowsSkippedAfterConflict, anomalyTotal: r.anomalyTotal, firstAt: r.firstAt, lastAt: r.lastAt }).commit();
    }
    if (r.anomalyTotal) {
      am.addRow({ meterNo: r.meterNo, pocNo: meta.pocNo ?? '', status: STATUS_LABEL[r.status] || r.status, records: r.records, anomalyTotal: r.anomalyTotal,
        ...Object.fromEntries(ANOMALY_TYPES.map((k) => [k, r.anomalies?.[k] || 0])), samples: r.anomalySamples?.length || 0 }).commit();
      if (r.status === 'no_records') an.addRow({ meterNo: r.meterNo, pocNo: meta.pocNo ?? '', status: STATUS_LABEL[r.status], at: '', type: ANOMALY_LABEL.no_records, regStatus: '', result: '', content: '' }).commit();
      for (const a of r.anomalySamples || []) an.addRow({ meterNo: r.meterNo, pocNo: meta.pocNo ?? '', status: STATUS_LABEL[r.status] || r.status, at: a.at, type: ANOMALY_LABEL[a.type] || a.type, regStatus: `${a.regStatus} · ${REG_STATUS_LABEL[a.regStatus] || '?'}`, result: a.result, content: a.content }).commit();
    }
  }
  for (const ws of [cf, an, am, me]) ws.commit();
  await wb.commit();
}

function buildDataJs(results, summary, maxOk) {
  // compact rows: [meterNo, statusIdx, records, firstAt, lastAt, moduleId, anomalyTotal, successes, failures, pocNo, dcuNo, moduleCount, upperCount, macCount]
  const meters = []; let okIncluded = 0, okOmitted = 0;
  const conflicts = []; const anomalyMeters = [];
  for (const r of results) {
    const meta = r.meta || {};
    const isPlainOk = r.status === 'ok';
    if (isPlainOk && okIncluded >= maxOk) { okOmitted++; continue; }
    if (isPlainOk) okIncluded++;
    meters.push([r.meterNo, STATUS.indexOf(r.status), r.records, r.firstAt, r.lastAt, r.moduleId, r.anomalyTotal, r.successes, r.failures, meta.pocNo ?? null, meta.dcuNo ?? null, r.moduleCount, r.upperDevices?.length ?? 0, r.macs?.length ?? 0]);
    if (r.conflict) {
      conflicts.push({ meterNo: r.meterNo, ...r.conflict, records: r.records, skipped: r.rowsSkippedAfterConflict, anomalyTotal: r.anomalyTotal, firstAt: r.firstAt, lastAt: r.lastAt,
        timeline: r.timeline, recordsByMonth: r.recordsByMonth, anomalies: r.anomalies, anomaliesByMonth: r.anomaliesByMonth, samples: r.anomalySamples, pocNo: meta.pocNo ?? null, dcuNo: meta.dcuNo ?? null });
    } else if (r.anomalyTotal) {
      anomalyMeters.push({ meterNo: r.meterNo, status: r.status, records: r.records, anomalyTotal: r.anomalyTotal, anomalies: r.anomalies, anomaliesByMonth: r.anomaliesByMonth, samples: r.anomalySamples,
        recordsByMonth: r.recordsByMonth, firstAt: r.firstAt, lastAt: r.lastAt, moduleId: r.moduleId, pocNo: meta.pocNo ?? null });
    }
  }
  summary.dashboardOkOmitted = okOmitted;
  const payload = { summary, STATUS, STATUS_LABEL, ANOMALY_TYPES, ANOMALY_LABEL, REG_STATUS_LABEL, meters, conflicts, anomalyMeters };
  return 'window.AUDIT_DATA = ' + JSON.stringify(payload) + ';\n';
}

async function cmdBuild(args) {
  ensureDirs();
  const maxOk = args?.['dashboard-max-ok'] !== undefined ? Number(args['dashboard-max-ok']) : DEFAULTS.dashboardMaxOk;
  const progress = readJson(F.progress, null);
  const results = await loadResults();
  log(`build: ${results.length} analysed meters`);
  const summary = summarize(results, progress);
  writeJsonAtomic(F.summaryJson, summary);
  const vendorTags = await copyVendor();
  writeFileSync(F.dataJs, buildDataJs(results, summary, maxOk), 'utf8');
  writeFileSync(F.indexHtml, dashboardHtml(vendorTags), 'utf8');
  await writeXlsx(results, summary);
  log(`dashboard: ${F.indexHtml}`);
  log(`xlsx: ${F.xlsx}`);
  log(`summary: ${summary.analyzed}/${summary.metersTotal} analysed · conflicts ${summary.status.conflict} · with anomalies ${summary.metersWithAnomalies} · no records ${summary.status.no_records} · consistent ${summary.status.ok + summary.status.ok_anomalies}`);
}

// ───────────────────────────────────────────────────────────── dashboard template
function dashboardHtml(vendorTags) {
  return `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Auditoria de moduleId</title>
<style>
  :root {
    color-scheme: light;
    --surface-0: #f4f4f2; --surface-1: #fcfcfb; --surface-2: #ffffff;
    --border: #e2e1dc; --grid: #ecebe6;
    --text-primary: #0b0b0b; --text-secondary: #52514e; --text-muted: #7c7b76;
    --c-conflict: #d03b3b; --c-ok: #1baf7a; --c-ok-anom: #eda100; --c-nomod: #4a3aa7; --c-norec: #9a9994;
    --c-empty: #eb6834; --c-json: #4a3aa7; --c-missing: #2a78d6;
    --accent: #2a78d6; --shadow: 0 1px 2px rgba(0,0,0,.05);
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      color-scheme: dark;
      --surface-0: #111110; --surface-1: #1a1a19; --surface-2: #222221; --border: #30302e; --grid: #2a2a28;
      --text-primary: #ffffff; --text-secondary: #c3c2b7; --text-muted: #8b8a83;
      --c-conflict: #e55555; --c-ok: #199e70; --c-ok-anom: #c98500; --c-nomod: #9085e9; --c-norec: #7a7975;
      --c-empty: #d95926; --c-json: #9085e9; --c-missing: #3987e5; --accent: #3987e5; --shadow: none;
    }
  }
  :root[data-theme="dark"] {
    color-scheme: dark;
    --surface-0: #111110; --surface-1: #1a1a19; --surface-2: #222221; --border: #30302e; --grid: #2a2a28;
    --text-primary: #ffffff; --text-secondary: #c3c2b7; --text-muted: #8b8a83;
    --c-conflict: #e55555; --c-ok: #199e70; --c-ok-anom: #c98500; --c-nomod: #9085e9; --c-norec: #7a7975;
    --c-empty: #d95926; --c-json: #9085e9; --c-missing: #3987e5; --accent: #3987e5; --shadow: none;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; }
  body { background: var(--surface-0); color: var(--text-primary); font: 14px/1.45 system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; padding: 16px; }
  .wrap { max-width: 1400px; margin: 0 auto; }
  header { display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end; justify-content: space-between; margin-bottom: 12px; }
  h1 { font-size: 22px; margin: 0 0 4px; font-weight: 600; }
  .sub { color: var(--text-secondary); margin: 0; }
  h2 { font-size: 15px; font-weight: 600; margin: 0 0 10px; display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
  .muted { color: var(--text-muted); font-size: 12px; font-weight: 400; }
  button, select, input { font: inherit; color: var(--text-primary); background: var(--surface-2); border: 1px solid var(--border); border-radius: 6px; padding: 6px 8px; }
  button { cursor: pointer; } button:hover { border-color: var(--accent); }
  button.active { background: var(--accent); color: #fff; border-color: var(--accent); }
  a.btn { display: inline-block; text-decoration: none; color: var(--text-primary); background: var(--surface-2); border: 1px solid var(--border); border-radius: 6px; padding: 6px 8px; }
  .rule { background: var(--surface-1); border: 1px solid var(--border); border-left: 3px solid var(--accent); border-radius: 10px; padding: 10px 14px; margin-bottom: 12px; color: var(--text-secondary); }
  .rule b { color: var(--text-primary); }
  .progress { height: 8px; background: var(--grid); border-radius: 4px; overflow: hidden; margin: 6px 0 12px; }
  .progress > div { height: 100%; background: var(--accent); }
  .filters { display: flex; flex-wrap: wrap; gap: 10px 14px; align-items: flex-end; background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; margin-bottom: 10px; box-shadow: var(--shadow); }
  .filters label { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--text-secondary); }
  .filters label.check { flex-direction: row; align-items: center; gap: 6px; padding-bottom: 7px; }
  .filters input[type="number"] { width: 110px; } .filters input[type="text"] { width: 170px; } .filters .spacer { flex: 1; }
  .presets { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 0 0 14px; }
  .tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 16px; }
  .tile { background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; box-shadow: var(--shadow); }
  .tile .label { color: var(--text-secondary); font-size: 12px; } .tile .value { font-size: 26px; font-weight: 600; line-height: 1.2; margin-top: 2px; }
  .tile.hero .value { font-size: 48px; } .tile .delta { font-size: 12px; color: var(--text-muted); margin-top: 2px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 12px; margin-bottom: 16px; }
  .card { background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px; padding: 14px; box-shadow: var(--shadow); min-width: 0; }
  .card.wide { grid-column: 1 / -1; }
  .chart { position: relative; height: 260px; } .chart.tall { height: 520px; } .chart.short { height: 200px; }
  .legend { display: flex; gap: 14px; flex-wrap: wrap; font-size: 12px; color: var(--text-secondary); margin: 0 0 6px; }
  .legend span::before { content: ""; display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 6px; vertical-align: middle; background: var(--c); }
  .summary p { margin: 0 0 10px; } .summary p:last-child { margin-bottom: 0; } .summary.dynamic { border-left: 3px solid var(--accent); }
  table { width: 100%; border-collapse: collapse; font-size: 13px; font-variant-numeric: tabular-nums; }
  th, td { text-align: left; padding: 6px 8px; border-bottom: 1px solid var(--grid); white-space: nowrap; }
  th { color: var(--text-secondary); font-weight: 600; cursor: pointer; user-select: none; position: sticky; top: 0; background: var(--surface-1); }
  th.sorted::after { content: " ▾"; color: var(--text-muted); } th.sorted.asc::after { content: " ▴"; }
  td.num, th.num { text-align: right; } td.wrap { white-space: normal; max-width: 420px; }
  tbody tr { cursor: pointer; } tbody tr:hover { background: var(--surface-2); } tbody tr.selected { outline: 2px solid var(--accent); outline-offset: -2px; }
  tbody tr.hl { background: color-mix(in srgb, var(--c-conflict) 14%, transparent); }
  .tablewrap { overflow: auto; max-height: 460px; border: 1px solid var(--grid); border-radius: 6px; }
  .pager { display: flex; gap: 8px; align-items: center; justify-content: flex-end; margin-top: 8px; font-size: 12px; color: var(--text-secondary); }
  .badge { display: inline-block; padding: 1px 6px; border-radius: 10px; font-size: 11px; background: var(--surface-2); border: 1px solid var(--border); color: var(--text-secondary); }
  .badge.conflict { border-color: var(--c-conflict); color: var(--c-conflict); font-weight: 600; } .badge.ok { border-color: var(--c-ok); color: var(--c-ok); }
  .badge.ok_anomalies { border-color: var(--c-ok-anom); color: var(--text-primary); } .badge.no_module_id { border-color: var(--c-nomod); color: var(--c-nomod); } .badge.no_records { border-color: var(--c-norec); }
  .tabs { display: flex; gap: 6px; margin-bottom: 10px; flex-wrap: wrap; }
  .meter-head { display: flex; flex-wrap: wrap; gap: 8px 18px; align-items: baseline; margin-bottom: 8px; color: var(--text-secondary); font-size: 13px; }
  .meter-head strong { color: var(--text-primary); font-size: 16px; }
  .empty { color: var(--text-muted); padding: 24px; text-align: center; }
  code { font-size: 12px; background: var(--surface-2); border: 1px solid var(--border); border-radius: 4px; padding: 0 4px; }
  [hidden] { display: none !important; }
  @media (max-width: 600px) { .tile.hero .value { font-size: 36px; } .chart.tall { height: 380px; } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Auditoria de moduleId no registro automático</h1>
      <p class="sub" id="subtitle">Carregando…</p>
    </div>
    <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap">
      <button id="btnXlsx" title="Exporta as linhas filtradas da aba atual em XLSX">Exportar XLSX (filtro atual)</button>
      <a class="btn" id="lnkXlsx" href="../output/auto-register-audit.xlsx" title="Arquivo completo gerado pelo script (todas as abas)">XLSX completo</a>
      <button id="btnTheme" title="Alternar tema claro/escuro">Tema</button>
    </div>
  </header>
  <div class="progress" title="Progresso da análise"><div id="progressBar" style="width:0%"></div></div>

  <div class="rule" id="rule"></div>

  <div class="filters">
    <label>Status do medidor<select id="f-status"><option value="">Todos</option></select></label>
    <label>Mês <span class="muted">(conflito / anomalia / 1º registro)</span><select id="f-month"><option value="">Todos</option></select></label>
    <label>Tipo de anomalia<select id="f-anom"><option value="">Qualquer</option></select></label>
    <label>Status do registro no conflito<select id="f-regstatus"><option value="">Todos</option></select></label>
    <label>Mín. registros<input type="number" id="f-minrec" min="0" step="1" value="0"></label>
    <label>Máx. registros<input type="number" id="f-maxrec" min="0" step="1" placeholder="sem limite"></label>
    <label>Medidor / UC / moduleId<input type="text" id="f-text" placeholder="PMD… ou moduleId"></label>
    <label class="check"><input type="checkbox" id="f-onlyanom"> Só com anomalias</label>
    <label class="check"><input type="checkbox" id="f-multiupper"> Só com mais de um DCU/gateway</label>
    <span class="spacer"></span>
    <button id="f-reset">Limpar filtros</button>
  </div>
  <div class="presets">
    <span class="muted">Atalhos:</span>
    <button data-preset="all">Todos</button>
    <button data-preset="conflict">Conflitos</button>
    <button data-preset="anomalies">Com anomalias</button>
    <button data-preset="no_records">Sem registros</button>
    <button data-preset="no_module_id">Sem moduleId</button>
    <button data-preset="ok">Consistentes</button>
  </div>

  <div class="tiles" id="tiles"></div>
  <div class="grid">
    <div class="card wide summary dynamic"><h2>Resumo do filtro atual</h2><div id="dyn"></div></div>
    <div class="card"><h2>Medidores por status</h2><div class="legend" id="legStatus"></div><div class="chart"><canvas id="cStatus"></canvas></div></div>
    <div class="card"><h2>Conflitos por mês <span class="muted">(mês da mudança de moduleId)</span></h2><div class="chart"><canvas id="cConflictMonth"></canvas></div></div>
    <div class="card"><h2>Anomalias por tipo <span class="muted">(registros)</span></h2><div class="legend" id="legAnom"></div><div class="chart"><canvas id="cAnomType"></canvas></div></div>
    <div class="card"><h2>Anomalias por mês <span class="muted">(registros)</span></h2><div class="chart"><canvas id="cAnomMonth"></canvas></div></div>
    <div class="card"><h2>Registros por medidor <span class="muted">(distribuição)</span></h2><div class="chart"><canvas id="cRecords"></canvas></div></div>
    <div class="card"><h2>Status do registro no momento do conflito</h2><div class="chart"><canvas id="cConflictStatus"></canvas></div></div>
    <div class="card wide"><h2>Registros de auto-registro por mês <span class="muted">(todos os medidores do filtro)</span></h2><div class="chart"><canvas id="cRecMonth"></canvas></div></div>
    <div class="card wide"><h2>Top 25 medidores <span class="muted">(clique para ver o detalhe)</span>
      <select id="topsort" style="margin-left:auto;font-size:12px"><option value="anomalyTotal">por anomalias</option><option value="records">por registros</option><option value="failures">por falhas de registro</option><option value="upperCount">por nº de DCUs/gateways</option></select></h2>
      <div class="chart tall"><canvas id="cTop"></canvas></div></div>
    <div class="card wide"><h2>Detalhe do medidor <span class="muted">(linha do tempo dos registros guardados: janela até o conflito, ou amostras de anomalias)</span></h2>
      <div class="meter-head" id="meterHead"></div>
      <div class="chart short" id="meterChartWrap"><canvas id="cMeter"></canvas></div>
      <div class="tablewrap" style="max-height:320px"><table id="tlTbl"><thead></thead><tbody></tbody></table></div></div>
    <div class="card wide">
      <div class="tabs"><button data-tab="conflicts" class="active">Conflitos</button><button data-tab="anomalies">Anomalias (por medidor)</button><button data-tab="meters">Medidores</button></div>
      <div class="tablewrap"><table id="tbl"><thead></thead><tbody></tbody></table></div><div class="pager" id="pager"></div>
    </div>
  </div>
  <section class="grid">
    <div class="card wide summary"><h2>Texto-resumo da análise completa</h2><div id="staticSummary"></div><p class="muted" id="runInfo"></p></div>
  </section>
</div>

${vendorTags}
<script src="data.js"></script>
<script>
(function () {
  "use strict";
  const D = window.AUDIT_DATA;
  if (!D) { document.getElementById("subtitle").textContent = "data.js não encontrado. Execute node audit.mjs build."; return; }
  const S = D.summary;
  const $ = (id) => document.getElementById(id);
  const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
  const fmt = (n, d = 0) => Number(n).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
  const pct = (a, b) => (b ? (100 * a / b).toFixed(1).replace(".", ",") + "%" : "–");
  const MONTHS_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  const monthLabel = (k) => { if (!k) return "–"; const [y, m] = k.split("-"); return MONTHS_PT[+m - 1] + "/" + y; };
  const STATUS = D.STATUS, SL = D.STATUS_LABEL, AT = D.ANOMALY_TYPES, AL = D.ANOMALY_LABEL, RL = D.REG_STATUS_LABEL;
  const statusColor = () => ({ conflict: css("--c-conflict"), ok: css("--c-ok"), ok_anomalies: css("--c-ok-anom"), no_module_id: css("--c-nomod"), no_records: css("--c-norec") });
  const anomColor = () => ({ empty_content: css("--c-empty"), invalid_json: css("--c-json"), missing_module_id: css("--c-missing"), no_records: css("--c-norec") });

  // ------------------------------------------------------------- data
  const conflictsBy = new Map(D.conflicts.map((c) => [c.meterNo, c]));
  const anomBy = new Map(D.anomalyMeters.map((a) => [a.meterNo, a]));
  const meters = D.meters.map((r) => {
    const st = STATUS[r[1]]; const c = conflictsBy.get(r[0]); const a = anomBy.get(r[0]); const src = c || a || null;
    const anomalies = src ? (src.anomalies || {}) : (st === "no_records" ? { no_records: 1 } : {});
    return { meterNo: r[0], status: st, records: r[2], firstAt: r[3], lastAt: r[4], moduleId: r[5], anomalyTotal: r[6], successes: r[7], failures: r[8], pocNo: r[9], dcuNo: r[10], moduleCount: r[11], upperCount: r[12], macCount: r[13],
      month: c ? c.month : (a && Object.keys(a.anomaliesByMonth || {}).sort()[0]) || (r[3] ? r[3].slice(0, 7) : null), anomalies, conflict: c || null, anom: a || null, src,
      conflictAt: c ? c.at : null, newModuleId: c ? c.newModuleId : null, regStatus: c ? c.regStatus : null };
  });
  const metersById = new Map(meters.map((m) => [m.meterNo, m]));
  delete D.meters;

  // ------------------------------------------------------------ theme
  function applyTheme(t) { if (t) document.documentElement.setAttribute("data-theme", t); else document.documentElement.removeAttribute("data-theme"); try { localStorage.setItem("auditTheme", t || ""); } catch (e) {} }
  try { applyTheme(localStorage.getItem("auditTheme") || ""); } catch (e) {}
  $("btnTheme").addEventListener("click", () => { const dark = matchMedia("(prefers-color-scheme: dark)").matches; const cur = document.documentElement.getAttribute("data-theme") || (dark ? "dark" : "light"); applyTheme(cur === "dark" ? "light" : "dark"); renderAll(); });
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", renderAll);

  // ------------------------------------------------------------ header
  $("subtitle").textContent = \`\${fmt(S.analyzed)} de \${fmt(S.metersTotal)} medidores analisados (\${pct(S.analyzed, S.metersTotal)})\${S.finished ? " · análise concluída" : " · análise em andamento"} · \${fmt(S.recordsTotal)} registros lidos · fonte \${S.store || S.source || "?"} · gerado em \${(S.generatedAt || "").replace("T", " ").slice(0, 19)}\`;
  $("progressBar").style.width = pct(S.analyzed, S.metersTotal).replace(",", ".");
  $("rule").innerHTML = "<b>Regra:</b> para cada medidor da lista, os registros da tela <i>Registro de dispositivo</i> (tabela <code>ops_auto_register</code>) são ordenados por <i>Registrado em</i> (crescente) e o campo <code>moduleId</code> do JSON da coluna <i>Conteúdo</i> é acompanhado. <b>Conflito</b> = o moduleId mudou ao longo do tempo (o medidor deixa de ser analisado a partir daí). <b>Anomalia</b> = registro com conteúdo vazio, que não é JSON ou JSON sem moduleId (contada; o medidor continua sendo analisado). Medidor sem nenhum registro = anomalia <i>Sem registros</i>." + (S.dashboardOkOmitted ? \` <b>Atenção:</b> \${fmt(S.dashboardOkOmitted)} medidores consistentes sem anomalias foram omitidos das tabelas deste dashboard (limite --dashboard-max-ok); os totais e o XLSX completo os incluem.\` : "");
  for (const k of STATUS) { const o = document.createElement("option"); o.value = k; o.textContent = SL[k]; $("f-status").appendChild(o); }
  for (const k of AT) { const o = document.createElement("option"); o.value = k; o.textContent = AL[k]; $("f-anom").appendChild(o); }
  for (const mk of S.months) { const o = document.createElement("option"); o.value = mk; o.textContent = monthLabel(mk); $("f-month").appendChild(o); }
  for (const k of Object.keys(S.conflictRegStatus).sort()) { const o = document.createElement("option"); o.value = k; o.textContent = \`\${k} · \${RL[k] || "?"}\`; $("f-regstatus").appendChild(o); }

  // ----------------------------------------------------------- generic UI
  let selectedMeter = null; let tab = "conflicts";
  const charts = {};
  const colors = () => ({ grid: css("--grid"), text: css("--text-secondary"), textP: css("--text-primary"), accent: css("--accent"), surface2: css("--surface-2") });
  function baseOpts(c, extra) {
    return Object.assign({ responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: "index", intersect: false },
      plugins: { legend: { display: false }, tooltip: { backgroundColor: c.surface2, titleColor: c.textP, bodyColor: c.textP, borderColor: c.grid, borderWidth: 1, padding: 8, boxWidth: 10, boxHeight: 2 } },
      scales: { x: { grid: { color: c.grid, drawTicks: false }, border: { color: c.grid }, ticks: { color: c.text, maxRotation: 0, autoSkip: true } },
        y: { beginAtZero: true, grid: { color: c.grid, drawTicks: false }, border: { display: false }, ticks: { color: c.text, precision: 0, callback: (v) => fmt(v) } } } }, extra || {});
  }
  function mk(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart($(id).getContext("2d"), cfg); return charts[id]; }
  function setLegend(id, items) { const root = $(id); root.textContent = ""; for (const it of items) { const s = document.createElement("span"); s.style.setProperty("--c", it.color); s.textContent = it.label; root.appendChild(s); } }
  const barDs = (label, data, color, extra) => Object.assign({ label, data, backgroundColor: color, borderWidth: 0, borderRadius: 4, borderSkipped: "start", maxBarThickness: 36, categoryPercentage: 0.7, barPercentage: 0.9 }, extra || {});
  const badge = (cls, text) => { const b = document.createElement("span"); b.className = "badge " + cls; b.textContent = text; return b; };
  function tiles(rootId, items) { const root = $(rootId); root.textContent = ""; for (const t of items) { const d = document.createElement("div"); d.className = "tile" + (t.hero ? " hero" : ""); const l = document.createElement("div"); l.className = "label"; l.textContent = t.label; const v = document.createElement("div"); v.className = "value"; v.textContent = t.value; const x = document.createElement("div"); x.className = "delta"; x.textContent = t.delta || ""; d.append(l, v, x); root.appendChild(d); } }
  function paragraphs(rootId, parts) { const root = $(rootId); root.textContent = ""; for (const p of parts) { const el = document.createElement("p"); el.textContent = p; root.appendChild(el); } }

  // ------------------------------------------------------------ filters
  const Fv = {};
  function readF() {
    Fv.status = $("f-status").value; Fv.month = $("f-month").value; Fv.anom = $("f-anom").value; Fv.regStatus = $("f-regstatus").value;
    Fv.minRec = Number($("f-minrec").value) || 0; Fv.maxRec = $("f-maxrec").value === "" ? Infinity : Number($("f-maxrec").value);
    Fv.text = $("f-text").value.trim().toLowerCase(); Fv.onlyAnom = $("f-onlyanom").checked; Fv.multiUpper = $("f-multiupper").checked;
  }
  function passes(m) {
    if (Fv.status && m.status !== Fv.status) return false;
    if (Fv.month && m.month !== Fv.month && !(m.src && m.src.anomaliesByMonth && m.src.anomaliesByMonth[Fv.month]) && !(m.conflict && m.conflict.month === Fv.month)) return false;
    if (Fv.anom && !(m.anomalies && m.anomalies[Fv.anom])) return false;
    if (Fv.regStatus && m.regStatus !== Fv.regStatus) return false;
    if (m.records < Fv.minRec || m.records > Fv.maxRec) return false;
    if (Fv.onlyAnom && !m.anomalyTotal) return false;
    if (Fv.multiUpper && m.upperCount < 2) return false;
    if (Fv.text && !(m.meterNo.toLowerCase().includes(Fv.text) || (m.pocNo && String(m.pocNo).toLowerCase().includes(Fv.text)) || (m.moduleId && String(m.moduleId).toLowerCase().includes(Fv.text)) || (m.newModuleId && String(m.newModuleId).toLowerCase().includes(Fv.text)))) return false;
    return true;
  }
  const PRESETS = { all: {}, conflict: { "f-status": "conflict" }, anomalies: { "f-onlyanom": true }, no_records: { "f-status": "no_records" }, no_module_id: { "f-status": "no_module_id" }, ok: { "f-status": "ok" } };
  function resetF() { $("f-status").value = ""; $("f-month").value = ""; $("f-anom").value = ""; $("f-regstatus").value = ""; $("f-minrec").value = "0"; $("f-maxrec").value = ""; $("f-text").value = ""; $("f-onlyanom").checked = false; $("f-multiupper").checked = false; }
  for (const id of ["f-status", "f-month", "f-anom", "f-regstatus", "f-minrec", "f-maxrec", "f-onlyanom", "f-multiupper"]) $(id).addEventListener("change", () => { tbl.page = 0; renderAll(); });
  $("f-text").addEventListener("input", () => { tbl.page = 0; renderAll(); });
  $("f-reset").addEventListener("click", () => { resetF(); tbl.page = 0; renderAll(); });
  for (const b of document.querySelectorAll("[data-preset]")) b.addEventListener("click", () => { resetF(); const p = PRESETS[b.dataset.preset]; for (const [k, v] of Object.entries(p)) { if (typeof v === "boolean") $(k).checked = v; else $(k).value = v; } tbl.page = 0; renderAll(); });
  for (const b of document.querySelectorAll("[data-tab]")) b.addEventListener("click", () => { tab = b.dataset.tab; for (const x of document.querySelectorAll("[data-tab]")) x.classList.toggle("active", x === b); tbl.page = 0; tbl.sortKey = null; renderTable(); });
  $("topsort").addEventListener("change", renderAll);

  // ------------------------------------------------------------ table
  const tbl = { page: 0, sortKey: null, sortAsc: true };
  const PAGE = 50;
  let filtered = []; let tableRows = [];
  const colStatus = { key: "status", label: "Status", render: (td, v) => td.appendChild(badge(v, SL[v] || v)) };
  const colRegStatus = (key) => ({ key, label: "Status do registro", render: (td, v) => { td.textContent = v === null || v === undefined || v === "" ? "–" : \`\${v} · \${RL[v] || "?"}\`; } });
  const COLS = {
    conflicts: [ { key: "meterNo", label: "Medidor" }, { key: "pocNo", label: "UC" }, { key: "conflictAt", label: "Conflito em" }, { key: "moduleId", label: "moduleId anterior" }, { key: "newModuleId", label: "moduleId novo" },
      { key: "c_previousSince", label: "Anterior desde" }, { key: "c_previousCount", label: "Registros c/ anterior", num: true }, colRegStatus("regStatus"), { key: "c_result", label: "Resultado", wrap: true }, { key: "c_upperDeviceNo", label: "Registrado via" }, { key: "c_macAddress", label: "MAC" },
      { key: "records", label: "Registros", num: true }, { key: "c_skipped", label: "Ignorados após", num: true }, { key: "anomalyTotal", label: "Anomalias antes", num: true } ],
    anomalies: [ { key: "meterNo", label: "Medidor" }, { key: "pocNo", label: "UC" }, colStatus, { key: "records", label: "Registros", num: true }, { key: "anomalyTotal", label: "Anomalias", num: true },
      ...AT.map((k) => ({ key: "a_" + k, label: AL[k], num: true })), { key: "firstAt", label: "Primeiro registro" }, { key: "lastAt", label: "Último registro" }, { key: "moduleId", label: "moduleId" } ],
    meters: [ { key: "meterNo", label: "Medidor" }, { key: "pocNo", label: "UC" }, { key: "dcuNo", label: "DCU (cadastro)" }, colStatus, { key: "records", label: "Registros", num: true }, { key: "firstAt", label: "Primeiro registro" }, { key: "lastAt", label: "Último registro" },
      { key: "moduleId", label: "moduleId" }, { key: "moduleCount", label: "Registros c/ moduleId", num: true }, { key: "successes", label: "Sucessos", num: true }, { key: "failures", label: "Falhas", num: true }, { key: "anomalyTotal", label: "Anomalias", num: true }, { key: "upperCount", label: "DCUs/gateways", num: true }, { key: "macCount", label: "MACs", num: true } ],
  };
  function rowValue(m, key) {
    if (key.startsWith("c_")) return m.conflict ? m.conflict[key.slice(2)] : null;
    if (key.startsWith("a_")) return (m.anomalies || {})[key.slice(2)] || 0;
    return m[key];
  }
  function currentRows() { if (tab === "conflicts") return filtered.filter((m) => m.conflict); if (tab === "anomalies") return filtered.filter((m) => m.anomalyTotal); return filtered; }
  function renderTable() {
    const cols = COLS[tab]; let rows = currentRows();
    if (tbl.sortKey) rows = [...rows].sort((a, b) => { const x = rowValue(a, tbl.sortKey), y = rowValue(b, tbl.sortKey); const r = typeof x === "number" && typeof y === "number" ? x - y : String(x ?? "").localeCompare(String(y ?? "")); return tbl.sortAsc ? r : -r; });
    tableRows = rows;
    const t = $("tbl"); const thead = t.querySelector("thead"); const tbody = t.querySelector("tbody"); thead.textContent = ""; tbody.textContent = "";
    const tr = document.createElement("tr");
    for (const col of cols) { const th = document.createElement("th"); th.textContent = col.label; if (col.num) th.className = "num"; if (tbl.sortKey === col.key) th.classList.add("sorted", tbl.sortAsc ? "asc" : "desc"); th.addEventListener("click", () => { if (tbl.sortKey === col.key) tbl.sortAsc = !tbl.sortAsc; else { tbl.sortKey = col.key; tbl.sortAsc = !col.num; } tbl.page = 0; renderTable(); }); tr.appendChild(th); }
    thead.appendChild(tr);
    const pages = Math.max(1, Math.ceil(rows.length / PAGE)); if (tbl.page >= pages) tbl.page = pages - 1;
    const slice = rows.slice(tbl.page * PAGE, tbl.page * PAGE + PAGE);
    if (!slice.length) { const td = document.createElement("td"); td.colSpan = cols.length; td.className = "empty"; td.textContent = "Nenhum registro."; const r = document.createElement("tr"); r.appendChild(td); tbody.appendChild(r); }
    for (const m of slice) {
      const r = document.createElement("tr"); if (m.meterNo === selectedMeter) r.classList.add("selected"); r.addEventListener("click", () => selectMeter(m.meterNo));
      for (const col of cols) { const td = document.createElement("td"); const v = rowValue(m, col.key); if (col.render) col.render(td, v, m); else if (col.num) { td.className = "num"; td.textContent = v === undefined || v === null ? "–" : fmt(v); } else { td.textContent = v === undefined || v === null || v === "" ? "–" : String(v); if (col.wrap) td.className = "wrap"; } r.appendChild(td); }
      tbody.appendChild(r);
    }
    const pager = $("pager"); pager.textContent = "";
    const info = document.createElement("span"); info.textContent = \`\${fmt(rows.length)} registros · página \${tbl.page + 1} de \${pages}\`;
    const prev = document.createElement("button"); prev.textContent = "‹ anterior"; prev.disabled = tbl.page === 0; prev.addEventListener("click", () => { tbl.page--; renderTable(); });
    const next = document.createElement("button"); next.textContent = "próxima ›"; next.disabled = tbl.page >= pages - 1; next.addEventListener("click", () => { tbl.page++; renderTable(); });
    pager.append(info, prev, next);
  }

  // ------------------------------------------------------------ meter detail
  function selectMeter(no) { selectedMeter = no; renderDetail(); renderTable(); $("meterHead").scrollIntoView({ behavior: "smooth", block: "center" }); }
  function renderDetail() {
    const root = $("meterHead"); root.textContent = ""; const t = $("tlTbl"); t.querySelector("thead").textContent = ""; t.querySelector("tbody").textContent = "";
    const m = selectedMeter && metersById.get(selectedMeter);
    if (!m) { const s = document.createElement("span"); s.className = "muted"; s.textContent = "Selecione um medidor na tabela ou no gráfico Top 25."; root.appendChild(s); if (charts.cMeter) { charts.cMeter.destroy(); delete charts.cMeter; } return; }
    const c = colors();
    const items = [["", m.meterNo, true], ["status", SL[m.status]], ["UC", m.pocNo || "–"], ["DCU (cadastro)", m.dcuNo || "–"], ["registros", fmt(m.records)], ["período", m.firstAt ? \`\${m.firstAt} a \${m.lastAt}\` : "–"], ["moduleId", m.moduleId || "–"], ["anomalias", fmt(m.anomalyTotal)]];
    if (m.conflict) items.push(["conflito em", m.conflict.at], ["moduleId novo", m.conflict.newModuleId], ["registrado via", m.conflict.upperDeviceNo || "–"], ["IP", m.conflict.ip || "–"], ["MAC", m.conflict.macAddress || "–"]);
    for (const [k, v, strong] of items) { const s = document.createElement("span"); if (strong) { const b = document.createElement("strong"); b.textContent = v; s.appendChild(b); s.appendChild(document.createTextNode(" ")); s.appendChild(badge(m.status, SL[m.status])); } else { s.textContent = k + ": "; const b = document.createElement("span"); b.style.color = c.textP; b.textContent = v; s.appendChild(b); } root.appendChild(s); }
    // per-meter monthly records chart (only stored for conflict/anomaly meters)
    const rbm = (m.src && m.src.recordsByMonth) || null;
    if (rbm && Object.keys(rbm).length) { const ms = Object.keys(rbm).sort(); const abm = (m.src && m.src.anomaliesByMonth) || {}; const o = baseOpts(c); o.scales.x.stacked = true; o.scales.y.stacked = true;
      mk("cMeter", { type: "bar", data: { labels: ms.map(monthLabel), datasets: [barDs("registros", ms.map((k) => rbm[k] - (abm[k] || 0)), c.accent, { stack: "s" }), barDs("com anomalia", ms.map((k) => abm[k] || 0), css("--c-empty"), { stack: "s" })] }, options: o }); $("meterChartWrap").hidden = false; }
    else { $("meterChartWrap").hidden = true; if (charts.cMeter) { charts.cMeter.destroy(); delete charts.cMeter; } }
    // timeline table
    const thead = t.querySelector("thead"); const tbody = t.querySelector("tbody");
    const hdr = document.createElement("tr"); for (const h of ["Registrado em", "Status", "moduleId", "Anomalia", "Resultado", "Registrado via", "dateTime no conteúdo"]) { const th = document.createElement("th"); th.textContent = h; hdr.appendChild(th); } thead.appendChild(hdr);
    let rows = [];
    if (m.conflict && m.conflict.timeline) rows = m.conflict.timeline.map((r) => ({ at: r[0], st: r[1], mod: r[2], an: r[3], res: r[4], up: r[5], dt: r[6], hl: r[0] === m.conflict.at && r[2] === m.conflict.newModuleId }));
    else if (m.src && m.src.samples) rows = m.src.samples.map((a) => ({ at: a.at, st: a.regStatus, mod: null, an: a.type, res: a.result, up: null, dt: null, note: a.content }));
    if (!rows.length) { const td = document.createElement("td"); td.colSpan = 7; td.className = "empty"; td.textContent = m.status === "no_records" ? "Medidor sem registros de auto-registro." : "Medidor consistente: a linha do tempo não é guardada no dashboard (consulte o XLSX ou a tela Registro de dispositivo)."; const r = document.createElement("tr"); r.appendChild(td); tbody.appendChild(r); return; }
    for (const r of rows) { const tr = document.createElement("tr"); if (r.hl) tr.classList.add("hl"); tr.style.cursor = "default";
      for (const v of [r.at, r.st === null || r.st === undefined || r.st === "" ? "–" : \`\${r.st} · \${RL[r.st] || "?"}\`, r.mod || "–", r.an ? AL[r.an] || r.an : "", r.res || "–", r.up || "–", r.dt || (r.note ? "conteúdo: " + r.note : "–")]) { const td = document.createElement("td"); td.textContent = v; if (v && String(v).length > 60) td.className = "wrap"; tr.appendChild(td); } tbody.appendChild(tr); }
  }

  // ------------------------------------------------------------ charts
  function renderCharts() {
    const c = colors(); const sc = statusColor(); const ac = anomColor();
    const byStatus = Object.fromEntries(STATUS.map((k) => [k, filtered.filter((m) => m.status === k).length]));
    setLegend("legStatus", STATUS.map((k) => ({ label: SL[k], color: sc[k] })));
    mk("cStatus", { type: "bar", data: { labels: STATUS.map((k) => SL[k].split(" (")[0]), datasets: [barDs("medidores", STATUS.map((k) => byStatus[k]), STATUS.map((k) => sc[k]))] }, options: baseOpts(c) });
    const months = S.months;
    const cm = {}; for (const m of filtered) if (m.conflict && m.conflict.month) cm[m.conflict.month] = (cm[m.conflict.month] || 0) + 1;
    mk("cConflictMonth", { type: "bar", data: { labels: months.map(monthLabel), datasets: [barDs("conflitos", months.map((k) => cm[k] || 0), sc.conflict)] }, options: baseOpts(c) });
    const at = Object.fromEntries(AT.map((k) => [k, 0])); for (const m of filtered) for (const k of AT) at[k] += (m.anomalies || {})[k] || 0;
    setLegend("legAnom", AT.map((k) => ({ label: AL[k], color: ac[k] })));
    mk("cAnomType", { type: "bar", data: { labels: AT.map((k) => AL[k]), datasets: [barDs("registros", AT.map((k) => at[k]), AT.map((k) => ac[k]))] }, options: baseOpts(c) });
    const am = {}; for (const m of filtered) { const src = (m.src && m.src.anomaliesByMonth) || null; if (src) for (const [k, n] of Object.entries(src)) am[k] = (am[k] || 0) + n; }
    mk("cAnomMonth", { type: "bar", data: { labels: months.map(monthLabel), datasets: [barDs("anomalias", months.map((k) => am[k] || 0), ac.empty_content)] }, options: baseOpts(c) });
    const bins = S.recordBins.map((b) => ({ ...b, n: filtered.filter((m) => m.records >= b.min && (b.max === null || m.records <= b.max)).length }));
    mk("cRecords", { type: "bar", data: { labels: bins.map((b) => b.label), datasets: [barDs("medidores", bins.map((b) => b.n), c.accent)] }, options: baseOpts(c) });
    const rs = {}; for (const m of filtered) if (m.conflict) { const k = m.conflict.regStatus || "?"; rs[k] = (rs[k] || 0) + 1; } const rsk = Object.keys(rs).sort();
    mk("cConflictStatus", { type: "bar", data: { labels: rsk.map((k) => \`\${k} · \${RL[k] || "?"}\`), datasets: [barDs("conflitos", rsk.map((k) => rs[k]), sc.conflict)] }, options: baseOpts(c) });
    const rm = {}; for (const m of filtered) { const src = m.src && m.src.recordsByMonth; if (src) for (const [k, n] of Object.entries(src)) rm[k] = (rm[k] || 0) + n; }
    const noFilter = !Fv.status && !Fv.month && !Fv.anom && !Fv.regStatus && !Fv.minRec && Fv.maxRec === Infinity && !Fv.text && !Fv.onlyAnom && !Fv.multiUpper;
    const recSrc = noFilter ? S.recordsByMonth : rm;
    mk("cRecMonth", { type: "bar", data: { labels: months.map(monthLabel), datasets: [barDs(noFilter ? "registros (todos)" : "registros (só medidores com conflito/anomalia do filtro)", months.map((k) => recSrc[k] || 0), c.accent)] }, options: baseOpts(c, { plugins: { legend: { display: true, labels: { color: c.text } } } }) });
    const key = $("topsort").value; const top = [...filtered].filter((m) => m[key] > 0).sort((a, b) => b[key] - a[key]).slice(0, 25);
    const o = baseOpts(c); o.indexAxis = "y"; o.onClick = (evt, els) => { if (els.length) selectMeter(top[els[0].index].meterNo); };
    o.scales = { x: { beginAtZero: true, grid: { color: c.grid, drawTicks: false }, border: { display: false }, ticks: { color: c.text, precision: 0, callback: (v) => fmt(v) } }, y: { grid: { display: false }, border: { color: c.grid }, ticks: { color: c.text, autoSkip: false, font: { size: 11 } } } };
    mk("cTop", { type: "bar", data: { labels: top.map((m) => m.meterNo + (m.pocNo ? " · " + m.pocNo : "")), datasets: [barDs($("topsort").selectedOptions[0].textContent, top.map((m) => m[key]), top.map((m) => sc[m.status]))] }, options: o });
  }

  function renderAll() {
    readF(); filtered = meters.filter(passes);
    const n = filtered.length; const conf = filtered.filter((m) => m.status === "conflict").length; const anom = filtered.filter((m) => m.anomalyTotal).length;
    const anomRows = filtered.reduce((s, m) => s + m.anomalyTotal, 0); const recs = filtered.reduce((s, m) => s + m.records, 0); const norec = filtered.filter((m) => m.status === "no_records").length;
    const okAll = filtered.filter((m) => m.status === "ok" || m.status === "ok_anomalies").length;
    tiles("tiles", [
      { label: "Medidores no filtro", value: fmt(n), delta: \`de \${fmt(S.analyzed)} analisados · \${fmt(S.pending)} pendentes\` },
      { label: "Conflitos de moduleId", value: fmt(conf), delta: pct(conf, n) + " do filtro", hero: true },
      { label: "Com anomalias", value: fmt(anom), delta: \`\${fmt(anomRows)} registros anômalos\` },
      { label: "Sem registros", value: fmt(norec), delta: pct(norec, n) + " do filtro" },
      { label: "Consistentes", value: fmt(okAll), delta: pct(okAll, n) + " do filtro" },
      { label: "Registros lidos", value: fmt(recs), delta: n ? \`média \${fmt(recs / n, 1)} por medidor\` : "" },
    ]);
    paragraphs("dyn", [
      \`\${fmt(n)} medidores no filtro: \${fmt(conf)} com conflito de moduleId (\${pct(conf, n)}), \${fmt(anom)} com pelo menos uma anomalia (\${pct(anom, n)}), \${fmt(norec)} sem nenhum registro de auto-registro e \${fmt(okAll)} consistentes.\`,
      \`Foram lidos \${fmt(recs)} registros nesses medidores; \${fmt(anomRows)} deles são anômalos (\${AT.map((k) => AL[k] + ": " + fmt(filtered.reduce((s, m) => s + ((m.anomalies || {})[k] || 0), 0))).join("; ")}).\`,
    ]);
    renderCharts(); renderTable(); renderDetail();
  }

  // ------------------------------------------------------------ static summary + export
  paragraphs("staticSummary", [
    \`\${fmt(S.analyzed)} de \${fmt(S.metersTotal)} medidores da lista foram analisados (\${pct(S.analyzed, S.metersTotal)})\${S.finished ? "; a análise está concluída." : "; a análise ainda está em andamento e este dashboard reflete o estado parcial."}\`,
    \`\${fmt(S.status.conflict)} medidores apresentaram conflito (o moduleId do conteúdo mudou ao longo do tempo), o que corresponde a \${pct(S.status.conflict, S.analyzed)} dos analisados. \${fmt(S.metersWithAnomalies)} medidores têm pelo menos uma anomalia (\${fmt(S.anomalyRows)} registros): \${AT.map((k) => AL[k] + " " + fmt(S.anomalies[k] || 0)).join(", ")}.\`,
    \`\${fmt(S.status.no_records)} medidores não têm nenhum registro de auto-registro; \${fmt(S.status.no_module_id)} têm registros mas nenhum com moduleId; \${fmt(S.status.ok)} são consistentes sem anomalias e \${fmt(S.status.ok_anomalies)} consistentes com anomalias. No total foram lidos \${fmt(S.recordsTotal)} registros (\${fmt(S.successes)} com sucesso, \${fmt(S.failures)} com falha).\`,
  ]);
  $("runInfo").textContent = \`Fonte: \${S.store || S.source || "?"}\${S.schema ? " · lista de medidores: mysql " + S.schema : ""} · início \${S.startedAt || "?"} · última atualização \${S.updatedAt || "?"} · dashboard gerado em \${S.generatedAt}\`;

  $("btnXlsx").addEventListener("click", () => {
    if (typeof XLSX === "undefined") { alert("Biblioteca XLSX não carregada (sem vendor/xlsx.full.min.js e sem internet)."); return; }
    const cols = COLS[tab]; const rows = tableRows.map((m) => Object.fromEntries(cols.map((col) => { const v = rowValue(m, col.key); return [col.label, col.key === "status" ? SL[v] || v : col.key === "regStatus" ? (v === null || v === undefined ? "" : \`\${v} · \${RL[v] || "?"}\`) : v ?? ""]; })));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), tab === "conflicts" ? "Conflitos" : tab === "anomalies" ? "Anomalias" : "Medidores");
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    XLSX.writeFile(wb, \`auditoria-moduleid-\${tab}-\${stamp}.xlsx\`);
  });

  renderAll();
})();
</script>
</body>
</html>
`;
}

// ───────────────────────────────────────────────────────────── probe (connectivity + document shape)
async function cmdProbe(args) {
  const meterNo = args.meter ? String(args.meter) : null;
  // MySQL (meter list)
  try {
    const schema = await detectAmiSchema(args.db);
    const db = mysqlSession(schema);
    const [{ n }] = await db.query('SELECT COUNT(*) AS n FROM archive_meter WHERE delete_flag = 0', [], 1);
    const [row] = await db.query('SELECT meter_no AS meterNo, dcu_no AS dcuNo FROM archive_meter WHERE delete_flag = 0 LIMIT 1', [], 1);
    await db.close();
    log(`mysql OK: schema ${schema} · archive_meter has ${n} active meters (e.g. ${row?.meterNo})`);
  } catch (err) { log(`mysql FAILED: ${err.message}`); }
  // Mongo (records)
  let mongo = null;
  try {
    mongo = await mongoSession(args['mongo-uri'] || null, args['mongo-db'] || DEFAULT_MONGO_DB);
    const cols = await resolveRegisterCollections(mongo, { explicit: args.collections, from: args.from, to: args.to });
    const latest = cols[cols.length - 1];
    const info = await mongoCollectionInfo(mongo, latest.name);
    const count = await mongo.db.collection(latest.name).estimatedDocumentCount();
    const sample = await mongo.db.collection(latest.name).findOne({}, { projection: MONGO_PROJECTION });
    log(`mongo OK: ${cols.length} collections · ${latest.name} has ~${count} documents · meter field ${info.meterField} · time stored as ${info.timeIsDate ? 'Date' : 'string'}`);
    if (sample) log(`sample document (${latest.name}): ${JSON.stringify(normalizeMongoRow(sample)).slice(0, 600)}`);
    if (meterNo) {
      const byMeter = await fetchRecordsMongo(mongo, [meterNo], cols);
      const rows = byMeter.get(meterNo) || [];
      const res = analyzeMeter({ meterNo }, rows, { maxSamples: 5, maxTimeline: 20 });
      log(`meter ${meterNo}: ${rows.length} records from ${res.firstAt ?? '-'} to ${res.lastAt ?? '-'} · status ${res.status} · moduleId ${res.moduleId ?? '-'}${res.conflict ? ` · conflict at ${res.conflict.at}: ${res.conflict.previousModuleId} → ${res.conflict.newModuleId}` : ''} · anomalies ${res.anomalyTotal}`);
      for (const r of rows.slice(-5)) log(`  ${r.registerTime} · status ${r.registerStatus} · moduleId ${parseContent(r.registerMsg).moduleId ?? '-'} · via ${r.upperDeviceNo ?? '-'}`);
    }
  } catch (err) { log(`mongo FAILED: ${err.message}`); }
  finally { await mongo?.close(); }
}

// ───────────────────────────────────────────────────────────── main
async function main() {
  const args = parseArgs(process.argv.slice(2));
  const cmd = args._[0] || 'run';
  ensureDirs();
  switch (cmd) {
    case 'run': await cmdRun(args); break;
    case 'build': await cmdBuild(args); break;
    case 'status': await cmdStatus(); break;
    case 'reset': cmdReset(args); break;
    case 'probe': await cmdProbe(args); break;
    default: console.log('usage: node audit.mjs run|build|status|reset|probe [options] (see header of this file)'); process.exit(2);
  }
}

export { analyzeMeter, parseContent, summarize, resolveRegisterTables, fetchRecordsMysql, ObjectId, mongoSession, mongoCollectionInfo, mongoTimeBound, MONGO_PROJECTION, resolveRegisterCollections, fetchRecordsMongo, normalizeMongoRow, mongoMeterField, suffixWindow, DEFAULT_MONGO_URI, DEFAULT_MONGO_DB, BASE_TABLE, STATUS, ANOMALY_TYPES, mysqlSession, detectAmiSchema, readJsonl, readJson, writeJsonAtomic, localStamp, normTime, trimStr, nowIso, ensureDirs, STATE_DIR, OUTPUT_DIR, DEFAULT_AMI_SCHEMA };

if (!process.env.AUDIT_NO_MAIN) main().catch((err) => { log(`FATAL: ${err.stack || err.message || err}`); process.exit(1); });
