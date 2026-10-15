# cemig-auto-register-audit

Audits the **moduleId** consistency of every CEMIG meter in the SanPlat auto-registration
records (screen *Gestão de Dispositivos → Registro de dispositivo*, MySQL table
`ops_auto_register`, column `register_msg` = "Conteúdo" JSON).

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

## Usage

```powershell
cd "C:\Users\nsn102249\OneDrive - nansen.com.br\Projects\cemig-auto-register-audit"
npm install
node audit.mjs run              # meter list + records straight from MySQL (default, no API login)
node audit.mjs status           # progress
node audit.mjs build            # rebuild dashboard/ + output/*.xlsx from the saved state (can run while `run` is going)
node audit.mjs reset --yes      # start over (--keep-meters keeps the cached meter list)
```

Open `dashboard\index.html` (works on `file://`, Chart.js and SheetJS are in `dashboard\vendor`).
Full workbook: `output\auto-register-audit.xlsx` (Resumo, Conflitos, Anomalias, Anomalias por medidor, Medidores).
The dashboard button *Exportar XLSX (filtro atual)* exports the filtered rows of the current tab.

| Option | Default | Description |
|---|---|---|
| `--source mysql|api` | `mysql` | `mysql` queries `ops_auto_register` directly (schema `ami` on the cemig-ext-data MySQL server, auto-detected if that one has no such table; force with `--db <schema>` or `AMI_MYSQL_DATABASE`). `api` uses the same endpoint as the screen (`/ami/collection/V1/register/result`), one PowerShell request per meter page – much slower. |
| `--meters mysql|api|file` | `mysql` | Meter list: `archive_meter` in the AMI schema (default, no login), MDC `/ami/archives/V1/meters` (same as `cemig-ext-data/fetch-meters-mdc.js`, logs in), or a text file (`--meters-file list.txt`, one meter per line). Cached in `state/meters.jsonl`. |
| `--batch N` | 200 | Meters per MySQL query / checkpoint. |
| `--limit N` | – | Analyse at most N meters in this run (useful for a first test: `--limit 1000`). |
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

Scans the whole `ops_auto_register` table in batches paginated by primary key (clustered scan,
no sort), extracts `moduleId` from every Conteúdo JSON and aggregates. Resumable (checkpoint = last id).

```powershell
node module-ids.mjs run [--batch 5000] [--from "2026-01-01 00:00:00"] [--to "..."] [--category 02] [--limit-batches N]
node module-ids.mjs status | build | reset --yes
```

Outputs: `output\module-ids.txt` (distinct list), `output\module-ids.xlsx` (Resumo · moduleIds · Medidores · Sem moduleId),
`output\module-ids-summary.json`. State and log in `state\module-ids\`.

## Query performance (table indexes: PK `id`, `index_register_time`, `index_device_no`)

- `audit.mjs`: `WHERE device_no IN (200 meters) ORDER BY device_no, register_time, id` → uses `index_device_no`; the sort runs only over the batch's rows.
- `module-ids.mjs`: `WHERE id > ? ORDER BY id LIMIT n` → sequential primary-key scan, no filesort; `register_time`/`device_category` are plain filters.
- `archive_meter` meter list: `ORDER BY meter_no LIMIT 50000 OFFSET n` (one-off, 15 pages).
