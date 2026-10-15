#!/usr/bin/env node
/**
 * Probe: where is the auto-register data and can this machine reach it?
 *
 *   node probe.mjs [--meter PMD268024393] [--pattern regist] [--mongo-uri ...] [--mongo-db ami] [--tunnel [host:port]]
 *                  [--collections a,b] [--db <mysql schema>] [--skip-mysql]
 *   node probe.mjs --ssh-scan [--hosts 10.1.2.3,10.1.2.4]   read-only scan on the cemig-ext-data SSH host: mongo settings in
 *                  SanPlat configs and java processes, live :27017 connections, TCP test of candidate Mongo hosts
 *
 * 1. Connects to the CEMIG Mongo (tries 10.30.104.31 then 10.30.132.30 unless --mongo-uri / MONGO_URI),
 *    lists every database, the collections matching --pattern (default: regist|auto) in each, and a sample
 *    document of any *auto_register* collection it finds.
 * 2. Checks MySQL (archive_meter) and the ops_auto_register_* collections of --mongo-db; with --meter,
 *    fetches that meter's records and applies the moduleId rule.
 * Nothing is written except state/run.log.
 */
process.env.AUDIT_NO_MAIN = '1';
const A = await import('./audit.mjs');

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

const args = parseArgs(process.argv.slice(2));
A.ensureDirs();
if (args['ssh-scan']) {
  console.log('── 0. ssh-scan: asking the cemig-ext-data jump host where Mongo is ──');
  try { await A.cmdSshScan(args); } catch (err) { console.log(`ssh-scan failed: ${err.message}`); }
  if (!args.continue) process.exit(0);
}
console.log('── 1. discovery ──────────────────────────────────────────────');
try { await A.cmdDiscover(args); } catch (err) { console.log(`discovery failed: ${err.message}`); }
console.log('── 2. probe ──────────────────────────────────────────────────');
if (args['skip-mysql']) process.env.AUDIT_SKIP_MYSQL = '1';
await A.cmdProbe(args);
