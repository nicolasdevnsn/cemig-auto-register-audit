# cemig-auto-register-audit

Audits the **moduleId** consistency of every CEMIG meter in the SanPlat auto-registration
records (screen *Gestão de Dispositivos → Registro de dispositivo*, MySQL table
`ops_auto_register`, column `register_msg` = "Conteúdo" JSON).

> **Where the rows really are.** `ops_auto_register` is registered in `db_split_table_def_n` with
> `split_uom = 2` (monthly). Every consumer of that registry reads through `mongoTemplate`, so the
> rows live in **MongoDB**, db `ami`, collections `ops_auto_register_yyyyMM` (snake_case fields:
> `device_no`, `register_time`, `register_status`, `register_msg` …). The MySQL table of the same
> name is an empty template (confirmed empty in prod, 2026-10-15). Both scripts default to Mongo and
> resolve the monthly collections from the server; `--source mysql` keeps the table path for deployments
> that use it. Prod check in Compass/mongosh: `db.getCollectionNames().filter(n => n.startsWith('ops_auto_register'))`.

Rule, per meter, records ordered by *Registrado em* ascending:

| Situation | Result |
|---|---|
| `moduleId` changes over time | **Conflict** – recorded, meter skipped from that point |
| content empty / not JSON / JSON without `moduleId` | **Anomaly** – counted, meter keeps being analysed |
| meter has no records at all | **Anomaly** `no_records` |

Resumable: stop with **Ctrl+C** (finishes the current batch, saves, rebuilds the dashboard);
run again to continue from the next meter.

## Requirements

- Node 18+, `npm install` once.
- `../cemig-ext-data` with its `.env` (credentials, MySQL config, client certificate) – reused as is.
- CEMIG VPN.

## Recommended path: API scan (works wherever the MDC login works)

The screen's endpoint accepts a time window without a meter number, so the whole history is pulled page by page
through the same MDC login cemig-ext-data uses, stored locally, and analysed offline. No MySQL/Mongo access is
needed for the records (the meter list still comes from MySQL `archive_meter`, or from a file with `--meters file`).

```powershell
node audit.mjs fetch-api                            # from 2020-01-01 to yesterday; resumable (Ctrl+C, run again)
node audit.mjs run --source jsonl                   # rule per meter from the local shards -> dashboard + xlsx
node module-ids.mjs run                             # moduleId list from the same shards (default --source jsonl)
node audit.mjs status                               # shows the scan cursor and the analysis progress
```

fetch-api options: `--from` (default 2020-01-01) / `--to` (default yesterday 23:59:59, so no window is still being written while it is paged), `--window` base hours per request (6), `--max-window` (720 h = 30 days), `--size` rows per page (500; a server cap is detected), `--shards` (64), `--limit-windows N`. **Adaptive window**: while windows come back empty the next one doubles (6 h, 12 h, … up to 30 days), so the years before the meters existed cost a few dozen requests instead of thousands; the first window with data drops back to 6 h, and a grown window that would need more than one page is split back down before anything is stored. The log reports "first records found in the window starting …" and, at the end, how many windows were empty. Each request is one PowerShell call (about 1-2 s). Note: the MDC login invalidates other sessions of the shared SanPlat account.

## Usage

```powershell
cd "C:\Users\nsn102249\OneDrive - nansen.com.br\Projects\cemig-auto-register-audit"
npm install
node audit.mjs probe --meter PMD268024393   # FIRST: checks MySQL + Mongo, lists the collections, shows one meter
node audit.mjs run              # meter list from MySQL archive_meter + records from Mongo ops_auto_register_yyyyMM (default, no API login)
node audit.mjs status           # progress
node audit.mjs build            # rebuild dashboard/ + output/*.xlsx from the saved state (can run while `run` is going)
node audit.mjs reset --yes      # start over (--keep-meters keeps the cached meter list)
```

Open `dashboard\index.html` (works on `file://`, Chart.js and SheetJS are in `dashboard\vendor`).
Full workbook: `output\auto-register-audit.xlsx` (Resumo, Conflitos, Anomalias, Anomalias por medidor, Medidores).
The dashboard button *Exportar XLSX (filtro atual)* exports the filtered rows of the current tab.

| Option | Default | Description |
|---|---|---|
| `--source mongo|mysql|api` | `mongo` | `mongo` reads the monthly collections in MongoDB (no credentials, `directConnection`, reads on secondary; by default it tries `10.30.104.31:27017` (primary, the one reachable from the CEMIG-side machines) then `10.30.132.30:27017` (secondary); `--mongo-uri` or env `MONGO_URI` pins one; `--mongo-db`, default `ami`; `--collections a,b` overrides the list). `mysql` reads `ops_auto_register_yyyyMM` tables in schema `ami` (`--db`, `--tables`). `api` uses the screen's endpoint (`/ami/collection/V1/register/result`), one PowerShell request per meter page – much slower. |
| `--meters mysql|api|file` | `mysql` | Meter list: `archive_meter` in the AMI schema (default, no login), MDC `/ami/archives/V1/meters` (same as `cemig-ext-data/fetch-meters-mdc.js`, logs in), or a text file (`--meters-file list.txt`, one meter per line). Cached in `state/meters.jsonl`. |
| `--batch N` | 200 | Meters per batch (one query round per batch) / checkpoint. |
| `--limit N` | – | Analyse at most N meters in this run (useful for a first test: `--limit 1000`). |
| `--from` / `--to` | – | Restrict the monthly collections/tables to this window (dates `YYYY-MM-DD`). |
| `--no-build` | – | Don't rebuild the dashboard at the end of the run. |
| `--dashboard-max-ok N` (build) | all | Cap of consistent meters without anomalies embedded in `data.js` (totals and the XLSX always include everything). |

Note: logging in to the MDC API (meter list or `--source api`) uses the shared SanPlat account and
invalidates any other session of that account, including the `cemig-ext-data` collector.

## Files

- `audit.mjs` – everything (fetch, analysis, state, dashboard template, XLSX).
- `state/` – `meters.jsonl`, `results.jsonl` (one line per analysed meter), `progress.json`, `run.log`.
- `dashboard/` – `index.html`, `data.js`, `vendor/`.
- `output/` – `auto-register-audit.xlsx`, `summary.json`.

## `module-ids.mjs` – moduleId extraction test

Scans every monthly `ops_auto_register_yyyyMM` **Mongo collection** (default; `--source mysql` for tables) in batches
paginated by `_id`, extracts `moduleId` from every Conteúdo JSON and aggregates. Resumable (checkpoint =
collection + last id; the collection list is frozen in progress.json). Options: `--batch`, `--from/--to`, `--collections`, `--category`, `--limit-batches`.

```powershell
node module-ids.mjs run [--batch 5000] [--from "2026-01-01 00:00:00"] [--to "..."] [--category 02] [--limit-batches N]
node module-ids.mjs status | build | reset --yes
```

Outputs: `output\module-ids.txt` (distinct list), `output\module-ids.xlsx` (Resumo · moduleIds · Medidores · Sem moduleId),
`output\module-ids-summary.json`. State and log in `state\module-ids\`.

## Query performance

- Mongo (`audit.mjs`): per batch, one `find({ device_no: { $in: [200 meters] } })` per monthly collection, in parallel; the registry defines an index starting with `device_no` on these collections. Rows are sorted per meter client-side.
- Mongo (`module-ids.mjs`): per collection, `find(filter).sort({ _id: 1 }).limit(n)` with `_id > last` → walks the `_id` index, no in-memory sort.
- MySQL fallback: `UNION ALL` over the monthly tables with `device_no IN (...)` (index `index_device_no`), and `WHERE id > ? ORDER BY id LIMIT n` (primary-key scan) respectively.
- Meter list: `archive_meter` in MySQL, `ORDER BY meter_no LIMIT 50000 OFFSET n` (one-off, ~15 pages).

## `probe.mjs` – find and check the data source

```powershell
node probe.mjs --meter PMD268024393            # discovery (all Mongo dbs + matching collections + samples) then the probe
node probe.mjs --pattern "auto_register|regist" --skip-mysql
```

Prints every database on the CEMIG Mongo, the collections matching the pattern, then samples one document per
collection family in every database and flags the families whose schema has the auto-register markers
(`register_msg`, `register_status`, `device_no` + `register_time`, a string containing `"moduleId"`), with field names
and types; the full inventory goes to `outputdiscover.json`. Then it shows a sample document of any
`*auto_register*` collection, then checks MySQL `archive_meter`, the `ops_auto_register_*` collections of
`--mongo-db` and one meter's records with the rule applied. Equivalent to `node audit.mjs discover` + `node audit.mjs probe`.

## Which Mongo is which (found 2026-10-15)

| Host | Role | Databases | Reachable from nansenu01 |
|---|---|---|---|
| 10.30.104.31:27017 | **MDM** Mongo | `mdm` (11.8 TB), `nms` | yes |
| 10.30.132.30:27017 | **AMI/MDC** Mongo (db `ami`: load profiles, `ops_auto_register_yyyyMM` …) | `ami` | no (different subnet) |

For a machine that only reaches 10.30.104.x, use the SSH jump host that cemig-ext-data already uses for RocketMQ
(`SSH_HOST` 10.30.104.40, credentials and plink path in its `.env`/`config.js`):

```powershell
node probe.mjs --tunnel                 # forwards 127.0.0.1:27117 -> 10.30.132.30:27017 through 10.30.104.40
node audit.mjs run --tunnel
node module-ids.mjs run --tunnel
```

`--tunnel other-host:27017` forwards to a different target. Manual equivalent:
`plink -N -L 127.0.0.1:27117:10.30.132.30:27017 <user>@10.30.104.40`, then `--mongo-uri "mongodb://127.0.0.1:27117/ami?directConnection=true"`.

### Finding the AMI Mongo from the jump host

```powershell
node probe.mjs --ssh-scan [--ssh-user u --ssh-pass p]   # uses cemig-ext-data SSH_HOST/SSH_USER/SSH_PASS + plink, exactly like its RocketMQ job
```

Runs read-only commands on 10.30.104.40 and prints: mongo settings found in SanPlat config files and running java
processes, established connections to port 27017, TCP reachability of candidate hosts (10.30.132.30, 10.30.104.31,
.30, .60, .66, .71 and `--hosts a,b`), and the box's own addresses. Passwords are masked. Then use the host it reveals with
`--mongo-uri` or `--tunnel host:27017`.

### Mongo credentials

If the AMI Mongo requires authentication (SanPlat connects as user `ami`, auth database = the data database), pass them
at run time; nothing is stored in this project:

```powershell
$env:MONGO_USER = "ami"; $env:MONGO_PASS = "<password>"          # or add MONGO_USER/MONGO_PASS to ..\cemig-ext-data\.env
node probe.mjs --tunnel
node audit.mjs run --tunnel
```

Flags `--mongo-user`, `--mongo-pass`, `--auth-source` override the environment. Special characters in the password are
URL-encoded automatically; in a hand-written `--mongo-uri` you must encode them yourself.

### Local overrides (`.env` in this folder)

Copy `.env.example` to `.env` (git-ignored) to override, for this project only, what comes from `..\cemig-ext-data\.env`:
`ACCOUNT` / `SANPLAT_PASSWORD` for the MDC login used by `fetch-api`, `MONGO_*`, `AUDIT_SSH_PASS`. It is loaded first,
so its values win; the collector's own file is never modified.

#### Pagination notes (API scan)

- A window is `[start, start + window - 1 s]`; windows are contiguous, so a row belongs to exactly one window.
- Pages are driven by the server's `pageCount`; without it, by `count` and the page size the server actually returned.
- The checkpoint is written after the rows of a page are appended, so an interruption can re-fetch one page; rows are
  deduplicated by their `id` when the analysis loads the shards.
- The server does not guarantee a stable order inside a window, so only multi-page windows can miss a row at a page
  boundary. The scan counts them and warns at the end; use a smaller `--window` if that number is significant.
