#!/usr/bin/env node
/**
 * moduleId extraction test
 * ========================
 *
 * Scans every monthly collection ops_auto_register_yyyyMM of MongoDB db 'ami' (the real storage of the
 * "Registro de dispositivo" screen; --source mysql reads ops_auto_register_yyyyMM MySQL tables instead) in
 * batches paginated by _id, parses the "Conteúdo" JSON (register_msg) of every document and extracts its
 * moduleId. Resumable: the checkpoint is collection + last _id; extracted rows go to state/module-ids/rows.jsonl.
 *
 * Commands
 *   node module-ids.mjs run   [--source mongo|mysql] [--batch 5000] [--from 'YYYY-MM-DD'] [--to 'YYYY-MM-DD']
 *                             [--collections a,b] [--mongo-uri ...] [--mongo-db ami] [--category 02] [--limit-batches N] [--no-build]
 *   node module-ids.mjs build        aggregate rows.jsonl -> output/module-ids.txt + output/module-ids.xlsx
 *   node module-ids.mjs status
 *   node module-ids.mjs reset --yes
 *
 * Outputs (output/)
 *   module-ids.txt            distinct moduleId values, one per line
 *   module-ids.xlsx           Resumo · moduleIds (records, meters, first/last seen) · Medidores (moduleIds per meter) · Sem moduleId
 *   module-ids-summary.json
 */

process.env.AUDIT_NO_MAIN = '1';
import { existsSync, mkdirSync, appendFileSync, unlinkSync, writeFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
const A = await import('./audit.mjs');

const DIR = join(A.STATE_DIR, 'module-ids');
const F = {
  rows: join(DIR, 'rows.jsonl'),
  progress: join(DIR, 'progress.json'),
  log: join(DIR, 'run.log'),
  txt: join(A.OUTPUT_DIR, 'module-ids.txt'),
  xlsx: join(A.OUTPUT_DIR, 'module-ids.xlsx'),
  summary: join(A.OUTPUT_DIR, 'module-ids-summary.json'),
};
const DEFAULTS = { batch: 5000 };

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) { args._.push(a); continue; }
    const key = a.slice(2); const next = argv[i + 1];
    if (key.startsWith('no-')) args[key.slice(3)] = false;
    else if (next !== undefined && !next.startsWith('--')) { args[key] = next; i++; }
    else args[key] = true;
  }
  return args;
}
function log(msg) { const line = `[${A.localStamp()}] ${msg}`; console.log(line); try { appendFileSync(F.log, line + '\n', 'utf8'); } catch { /* ignore */ } }

let stopRequested = false;
for (const s of ['SIGINT', 'SIGTERM', 'SIGBREAK']) { try { process.on(s, () => { if (stopRequested) process.exit(130); stopRequested = true; log(`${s} received — finishing the current batch and saving`); }); } catch { /* ignore */ } }

// ───────────────────────────────────────────────────────────── run
async function cmdRun(args) {
  mkdirSync(DIR, { recursive: true });
  const batch = Number(args.batch) || DEFAULTS.batch;
  const limitBatches = args['limit-batches'] !== undefined ? Number(args['limit-batches']) : Infinity;
  const category = args.category ? String(args.category) : null;
  const source = String(args.source || 'mongo');
  if (!['mongo', 'mysql'].includes(source)) throw new Error('--source must be mongo or mysql');
  const schema = source === 'mysql' ? await A.detectAmiSchema(args.db) : null;
  const progress = A.readJson(F.progress, { source, tables: null, tableIndex: 0, lastId: null, lastIdIsOid: false, lastTime: null, rowsRead: 0, rowsWithModuleId: 0, batches: 0, startedAt: A.nowIso(), runs: 0 });
  if (progress.source && progress.source !== source) throw new Error(`state was started with --source ${progress.source}; run reset --yes to switch to ${source}`);
  progress.source = source; progress.from = progress.from || (args.from ? String(args.from) : null); progress.to = progress.to || (args.to ? String(args.to) : null);
  progress.schema = schema; progress.category = progress.category ?? category; progress.runs++; delete progress.finishedAt;

  const db = source === 'mysql' ? A.mysqlSession(schema) : null;
  const mongo = source === 'mongo' ? await A.mongoSession(args['mongo-uri'] || A.DEFAULT_MONGO_URI, args['mongo-db'] || A.DEFAULT_MONGO_DB) : null;
  const t0 = Date.now(); let batchesThisRun = 0; let rowsThisRun = 0;
  try {
    // ops_auto_register is split into monthly physical collections/tables (ops_auto_register_yyyyMM); the list is
    // frozen in progress.json on the first run so a resume walks exactly the same sequence.
    if (!progress.tables) {
      const list = mongo ? await A.resolveRegisterCollections(mongo, { explicit: args.collections, from: progress.from, to: progress.to })
        : await A.resolveRegisterTables(db, { explicit: args.tables, from: progress.from, to: progress.to });
      progress.tables = list.map((t) => t.name);
    }
    A.writeJsonAtomic(F.progress, progress);
    log(`run: source=${source}${schema ? ' schema=' + schema : ''} batch=${batch} window=[${progress.from ?? '-'} .. ${progress.to ?? '-'}]${progress.category ? ` category=${progress.category}` : ''} · ${progress.tables.length} ${mongo ? 'collections' : 'tables'} · resuming at #${progress.tableIndex + 1} (${progress.tables[progress.tableIndex] ?? '-'}) after id ${progress.lastId ?? '-'} · ${progress.rowsRead} rows already read`);

    while (!stopRequested && batchesThisRun < limitBatches) {
      if (progress.tableIndex >= progress.tables.length) { progress.finishedAt = A.nowIso(); A.writeJsonAtomic(F.progress, progress); log('all collections read'); break; }
      const table = progress.tables[progress.tableIndex];
      let rows;
      if (mongo) {
        // keyset on _id (always indexed): sort {_id: 1}, resume with _id > last. Window/category are plain filters.
        const info = await A.mongoCollectionInfo(mongo, table);   // real field names + BSON type of register_time
        const filter = {};
        if (progress.lastId !== null) filter._id = { $gt: progress.lastIdIsOid ? new A.ObjectId(progress.lastId) : progress.lastId };
        if (progress.from || progress.to) filter[info.timeField] = A.mongoTimeBound(info, progress.from, progress.to);
        if (progress.category) filter[info.categoryField] = progress.category;
        const docs = await mongo.db.collection(table).find(filter, { projection: A.MONGO_PROJECTION }).sort({ _id: 1 }).limit(batch).toArray();
        rows = docs.map((d) => ({ ...A.normalizeMongoRow(d), _rawId: d._id }));
      } else {
        // keyset pagination on the PRIMARY KEY (id) inside each monthly table: sequential clustered-index scan,
        // no filesort, no OFFSET cost. The time window and category are plain filters.
        const where = []; const params = [];
        if (progress.from) { where.push('register_time >= ?'); params.push(progress.from); }
        if (progress.to) { where.push('register_time <= ?'); params.push(progress.to); }
        if (progress.lastId !== null) { where.push('id > ?'); params.push(progress.lastId); }
        if (progress.category) { where.push('device_category = ?'); params.push(progress.category); }
        rows = await db.query(
          `SELECT id, device_no AS deviceNo, device_category AS deviceCategory, register_time AS registerTime, register_status AS registerStatus, register_msg AS registerMsg, upper_device_no AS upperDeviceNo
             FROM ${table}${where.length ? ' WHERE ' + where.join(' AND ') : ''} ORDER BY id LIMIT ${batch}`, params
        );
      }
      const lines = []; let withModule = 0;
      for (const r of rows) {
        const c = A.parseContent(r.registerMsg);
        const t = A.normTime(r.registerTime);
        if (c.moduleId) withModule++;
        // [registerTime, deviceNo, moduleId|null, anomaly|null, registerStatus, deviceCategory, upperDeviceNo, meterNoInContent, table]
        lines.push(JSON.stringify([t, r.deviceNo, c.moduleId || null, c.anomaly, A.trimStr(r.registerStatus), r.deviceCategory, r.upperDeviceNo || null, c.json?.meterNo ?? null, table]));
      }
      if (lines.length) appendFileSync(F.rows, lines.join('\n') + '\n', 'utf8');   // rows first ...
      const done = rows.length < batch;
      if (rows.length) {
        const last = rows[rows.length - 1];
        if (mongo) { progress.lastIdIsOid = last._rawId instanceof A.ObjectId; progress.lastId = progress.lastIdIsOid ? last._rawId.toHexString() : last._rawId; }
        else progress.lastId = last.id;
        progress.lastTime = A.normTime(last.registerTime);
      }
      if (done) { progress.tableIndex++; progress.lastId = null; }                 // next collection starts from the beginning
      progress.rowsRead += rows.length; progress.rowsWithModuleId += withModule; progress.batches++; progress.updatedAt = A.nowIso();
      A.writeJsonAtomic(F.progress, progress);                                       // ... then the checkpoint
      batchesThisRun++; rowsThisRun += rows.length;
      const rate = rowsThisRun / Math.max(1, (Date.now() - t0) / 1000);
      log(`${table} batch ${progress.batches}: ${rows.length} rows (${withModule} with moduleId)${done ? ' · finished' : ` · last id ${progress.lastId}`} · total ${progress.rowsRead} rows · ${rate.toFixed(0)} rows/s`);
    }
  } finally { await db?.close(); await mongo?.close(); }
  if (!progress.finishedAt) log(stopRequested ? 'stopped by request — run again to continue' : 'stopped after --limit-batches');
  if (args.build !== false) await cmdBuild();
}

// ───────────────────────────────────────────────────────────── build
async function cmdBuild() {
  A.ensureDirs();
  const progress = A.readJson(F.progress, null);
  const modules = new Map();   // moduleId -> { records, meters:Set, first, last, statuses:{} }
  const meters = new Map();    // deviceNo -> { records, withModule, modules: Map<moduleId, {count, first, last}>, anomalies:{}, category }
  const anomalyTotals = {}; let rows = 0;
  for await (const r of A.readJsonl(F.rows)) {
    rows++;
    const [t, deviceNo, moduleId, anomaly, status, category, , meterNoInContent] = r;
    const m = meters.get(deviceNo) || { records: 0, withModule: 0, modules: new Map(), anomalies: {}, category, first: t, last: t, contentMeterMismatch: 0 };
    m.records++; if (t < m.first) m.first = t; if (t > m.last) m.last = t;
    if (meterNoInContent && String(meterNoInContent) !== String(deviceNo)) m.contentMeterMismatch++;
    if (anomaly) { m.anomalies[anomaly] = (m.anomalies[anomaly] || 0) + 1; anomalyTotals[anomaly] = (anomalyTotals[anomaly] || 0) + 1; }
    if (moduleId) {
      m.withModule++;
      const mm = m.modules.get(moduleId) || { count: 0, first: t, last: t }; mm.count++; if (t < mm.first) mm.first = t; if (t > mm.last) mm.last = t; m.modules.set(moduleId, mm);
      const x = modules.get(moduleId) || { records: 0, meters: new Set(), first: t, last: t, statuses: {} }; x.records++; x.meters.add(deviceNo); if (t < x.first) x.first = t; if (t > x.last) x.last = t; x.statuses[status || '?'] = (x.statuses[status || '?'] || 0) + 1; modules.set(moduleId, x);
    }
    meters.set(deviceNo, m);
  }
  const moduleIds = [...modules.keys()].sort();
  const metersMulti = [...meters.values()].filter((m) => m.modules.size > 1).length;
  const modulesShared = [...modules.values()].filter((x) => x.meters.size > 1).length;
  const summary = {
    generatedAt: A.localStamp(), finished: !!progress?.finishedAt, schema: progress?.schema || null, window: progress ? [progress.from, progress.to] : null, lastCheckpoint: progress ? [progress.lastTime, progress.lastId] : null,
    rows, rowsWithModuleId: [...modules.values()].reduce((s, x) => s + x.records, 0), distinctModuleIds: moduleIds.length, meters: meters.size,
    metersWithoutModuleId: [...meters.values()].filter((m) => m.modules.size === 0).length, metersWithSeveralModuleIds: metersMulti, moduleIdsSharedBySeveralMeters: modulesShared, anomalies: anomalyTotals,
  };
  A.writeJsonAtomic(F.summary, summary);
  writeFileSync(F.txt, moduleIds.join('\n') + (moduleIds.length ? '\n' : ''), 'utf8');

  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: F.xlsx, useStyles: true });
  const header = (ws, cols) => { ws.columns = cols.map(([h, k, w]) => ({ header: h, key: k, width: w || 16 })); ws.getRow(1).font = { bold: true }; ws.getRow(1).commit(); };
  const sm = wb.addWorksheet('Resumo'); header(sm, [['Indicador', 'k', 44], ['Valor', 'v', 30]]);
  for (const [k, v] of [['Gerado em', summary.generatedAt], ['Varredura concluída', summary.finished ? 'sim' : 'não'], ['Schema', summary.schema], ['Janela', summary.window?.join(' .. ')], ['Último checkpoint', summary.lastCheckpoint?.join(' / ')],
    ['Registros lidos', rows], ['Registros com moduleId', summary.rowsWithModuleId], ['moduleIds distintos', summary.distinctModuleIds], ['Medidores vistos', summary.meters], ['Medidores sem nenhum moduleId', summary.metersWithoutModuleId],
    ['Medidores com mais de um moduleId', metersMulti], ['moduleIds usados por mais de um medidor', modulesShared], ...Object.entries(anomalyTotals).map(([k, v]) => [`Anomalia · ${k}`, v])]) sm.addRow({ k, v: v ?? '' }).commit();
  sm.commit();
  const ws1 = wb.addWorksheet('moduleIds'); header(ws1, [['moduleId', 'id', 20], ['Registros', 'records'], ['Medidores', 'meters'], ['Primeira vez', 'first', 20], ['Última vez', 'last', 20], ['Sucessos (1)', 's1'], ['Falhas (2)', 's2'], ['Outros status', 'sx'], ['Medidores (até 20)', 'list', 80]]);
  for (const id of moduleIds) { const x = modules.get(id); const s1 = x.statuses['1'] || 0, s2 = x.statuses['2'] || 0; ws1.addRow({ id, records: x.records, meters: x.meters.size, first: x.first, last: x.last, s1, s2, sx: x.records - s1 - s2, list: [...x.meters].slice(0, 20).join(', ') + (x.meters.size > 20 ? ', …' : '') }).commit(); }
  ws1.commit();
  const ws2 = wb.addWorksheet('Medidores'); header(ws2, [['Medidor', 'meterNo'], ['Tipo', 'category', 8], ['Registros', 'records'], ['Com moduleId', 'withModule'], ['Nº moduleIds', 'n'], ['moduleIds (count, primeira..última)', 'ids', 90], ['Primeiro registro', 'first', 20], ['Último registro', 'last', 20], ['Conteúdo com outro meterNo', 'mismatch'], ['Anomalias', 'anom', 40]]);
  for (const [meterNo, m] of [...meters.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const ids = [...m.modules.entries()].sort((a, b) => a[1].first.localeCompare(b[1].first)).map(([id, v]) => `${id} (${v.count}, ${v.first}..${v.last})`).join(' | ');
    ws2.addRow({ meterNo, category: m.category ?? '', records: m.records, withModule: m.withModule, n: m.modules.size, ids, first: m.first, last: m.last, mismatch: m.contentMeterMismatch, anom: Object.entries(m.anomalies).map(([k, v]) => `${k}=${v}`).join(', ') }).commit();
  }
  ws2.commit();
  const ws3 = wb.addWorksheet('Sem moduleId'); header(ws3, [['Medidor', 'meterNo'], ['Registros', 'records'], ['Conteúdo vazio', 'empty_content'], ['Não é JSON', 'invalid_json'], ['JSON sem moduleId', 'missing_module_id'], ['Primeiro registro', 'first', 20], ['Último registro', 'last', 20]]);
  for (const [meterNo, m] of meters) if (m.modules.size === 0) ws3.addRow({ meterNo, records: m.records, empty_content: m.anomalies.empty_content || 0, invalid_json: m.anomalies.invalid_json || 0, missing_module_id: m.anomalies.missing_module_id || 0, first: m.first, last: m.last }).commit();
  ws3.commit();
  await wb.commit();
  log(`build: ${rows} rows · ${summary.distinctModuleIds} distinct moduleIds · ${summary.meters} meters (${metersMulti} with several moduleIds, ${summary.metersWithoutModuleId} without any) · ${modulesShared} moduleIds shared by several meters`);
  log(`txt: ${F.txt}`); log(`xlsx: ${F.xlsx}`);
}

// ───────────────────────────────────────────────────────────── status / reset
function cmdStatus() {
  const p = A.readJson(F.progress, null);
  if (!p) { console.log('not started'); return; }
  console.log(`${p.rowsRead} rows read (${p.rowsWithModuleId} with moduleId) in ${p.batches} batches · source ${p.source}${p.schema ? ' schema ' + p.schema : ''} · window ${p.from ?? '-'} .. ${p.to ?? '-'} · at #${(p.tableIndex ?? 0) + 1}/${p.tables?.length ?? '?'} (${p.tables?.[p.tableIndex] ?? '-'}) · checkpoint ${p.lastTime ?? '-'} / ${p.lastId ?? '-'}`);
  console.log(`started ${p.startedAt} · updated ${p.updatedAt || '-'}${p.finishedAt ? ` · finished ${p.finishedAt}` : ' · not finished'} · runs ${p.runs}`);
}
function cmdReset(args) {
  if (!args.yes) { console.log('reset deletes state/module-ids/ and output/module-ids.*. Run again with --yes.'); return; }
  rmSync(DIR, { recursive: true, force: true });
  for (const f of [F.txt, F.xlsx, F.summary]) if (existsSync(f)) unlinkSync(f);
  console.log('state cleared');
}

const args = parseArgs(process.argv.slice(2));
mkdirSync(DIR, { recursive: true });
try {
  switch (args._[0] || 'run') {
    case 'run': await cmdRun(args); break;
    case 'build': await cmdBuild(); break;
    case 'status': cmdStatus(); break;
    case 'reset': cmdReset(args); break;
    default: console.log('usage: node module-ids.mjs run|build|status|reset [options] (see header of this file)'); process.exit(2);
  }
} catch (err) { log(`FATAL: ${err.stack || err.message || err}`); process.exit(1); }
