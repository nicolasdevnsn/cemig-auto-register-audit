#!/usr/bin/env node
/**
 * moduleId extraction test
 * ========================
 *
 * Scans the whole `ops_auto_register` table (schema `ami`) in batches, paginated by
 * primary key (clustered scan, no sort), parses the "Conteúdo" JSON (`register_msg`) of every row and
 * extracts its `moduleId`. Resumable: the checkpoint is the last id read; extracted rows are appended to state/module-ids/rows.jsonl.
 *
 * Commands
 *   node module-ids.mjs run   [--batch 5000] [--from 'YYYY-MM-DD HH:mm:ss'] [--to '...'] [--db ami]
 *                             [--category 02] [--limit-batches N] [--no-build]
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
const DEFAULTS = { batch: 5000, from: '2000-01-01 00:00:00', to: '2999-12-31 23:59:59' };

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
  const schema = await A.detectAmiSchema(args.db);
  const progress = A.readJson(F.progress, { lastTime: null, lastId: null, rowsRead: 0, rowsWithModuleId: 0, batches: 0, startedAt: A.nowIso(), runs: 0 });
  progress.from = progress.from || String(args.from || DEFAULTS.from); progress.to = progress.to || String(args.to || DEFAULTS.to);
  progress.schema = schema; progress.category = progress.category ?? category; progress.runs++; delete progress.finishedAt;
  A.writeJsonAtomic(F.progress, progress);
  log(`run: schema=${schema} batch=${batch} window=[${progress.from} .. ${progress.to}]${progress.category ? ` category=${progress.category}` : ''} · resuming after id ${progress.lastId ?? '-'} · ${progress.rowsRead} rows already read`);

  const db = A.mysqlSession(schema);
  const t0 = Date.now(); let batchesThisRun = 0; let rowsThisRun = 0;
  try {
    while (!stopRequested && batchesThisRun < limitBatches) {
      // keyset pagination on the PRIMARY KEY (id): sequential clustered-index scan, no filesort, no OFFSET cost.
      // (register_time, id) would need a composite index the table does not have; the time window is just a filter.
      const where = ['register_time >= ?', 'register_time <= ?'];
      const params = [progress.from, progress.to];
      if (progress.lastId !== null) { where.push('id > ?'); params.push(progress.lastId); }
      if (progress.category) { where.push('device_category = ?'); params.push(progress.category); }
      const rows = await db.query(
        `SELECT id, device_no AS deviceNo, device_category AS deviceCategory, register_time AS registerTime, register_status AS registerStatus, register_msg AS registerMsg, upper_device_no AS upperDeviceNo
           FROM ops_auto_register WHERE ${where.join(' AND ')} ORDER BY id LIMIT ${batch}`, params
      );
      if (!rows.length) { progress.finishedAt = A.nowIso(); A.writeJsonAtomic(F.progress, progress); log('end of table reached'); break; }
      const lines = []; let withModule = 0;
      for (const r of rows) {
        const c = A.parseContent(r.registerMsg);
        const t = A.normTime(r.registerTime);
        if (c.moduleId) withModule++;
        // [registerTime, deviceNo, moduleId|null, anomaly|null, registerStatus, deviceCategory, upperDeviceNo, meterNoInContent]
        lines.push(JSON.stringify([t, r.deviceNo, c.moduleId || null, c.anomaly, A.trimStr(r.registerStatus), r.deviceCategory, r.upperDeviceNo || null, c.json?.meterNo ?? null]));
      }
      appendFileSync(F.rows, lines.join('\n') + '\n', 'utf8');            // rows first ...
      const last = rows[rows.length - 1];
      progress.lastId = last.id; progress.lastTime = A.normTime(last.registerTime);
      progress.rowsRead += rows.length; progress.rowsWithModuleId += withModule; progress.batches++; progress.updatedAt = A.nowIso();
      A.writeJsonAtomic(F.progress, progress);                              // ... then the checkpoint
      batchesThisRun++; rowsThisRun += rows.length;
      const rate = rowsThisRun / Math.max(1, (Date.now() - t0) / 1000);
      log(`batch ${progress.batches}: ${rows.length} rows (${withModule} with moduleId) · last id ${progress.lastId} · total ${progress.rowsRead} rows · ${rate.toFixed(0)} rows/s`);
      if (rows.length < batch) { progress.finishedAt = A.nowIso(); A.writeJsonAtomic(F.progress, progress); log('end of table reached'); break; }
    }
  } finally { await db.close(); }
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
  console.log(`${p.rowsRead} rows read (${p.rowsWithModuleId} with moduleId) in ${p.batches} batches · schema ${p.schema} · window ${p.from} .. ${p.to} · checkpoint ${p.lastTime ?? '-'} / ${p.lastId ?? '-'}`);
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
