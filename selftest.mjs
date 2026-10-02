import { spawn } from 'node:child_process'
import { copyFileSync, existsSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = dirname(fileURLToPath(import.meta.url))
const SERVER = join(HERE, 'server.mjs')
const SRC = process.argv[2]

if (!SRC) {
  console.error('usage: node selftest.mjs <path-to.sta>')
  console.error('example: node selftest.mjs "C:\\Users\\me\\Documents\\LAB3.sta"')
  process.exit(2)
}
if (!existsSync(SRC)) {
  console.error(`file not found: ${SRC}`)
  process.exit(2)
}

const WORK = join(process.env.TEMP, 'sta-selftest-work.sta')
const CSV = join(process.env.TEMP, 'sta-selftest-out.csv')
const SAVED = join(process.env.TEMP, 'sta-selftest-saved.sta')
const IMPORT_CSV = join(process.env.TEMP, 'sta-selftest-import.csv')
for (const f of [WORK, CSV, SAVED, IMPORT_CSV]) rmSync(f, { force: true })
copyFileSync(SRC, WORK)
writeFileSync(IMPORT_CSV, 'x;y;label\n1;2.5;a\n2;3.1;b\n3;4.7;c\n', 'utf8')

const p = spawn(process.execPath, [SERVER], { stdio: ['pipe', 'pipe', 'pipe'] })
let buf = Buffer.alloc(0)
const pending = new Map()
p.stdout.on('data', (c) => {
  buf = Buffer.concat([buf, c])
  let nl
  while ((nl = buf.indexOf(0x0a)) !== -1) {
    const line = buf.subarray(0, nl).toString('utf8').trim()
    buf = buf.subarray(nl + 1)
    if (!line) continue
    let msg
    try {
      msg = JSON.parse(line)
    } catch {
      console.error('unparsable line from server:', line.slice(0, 200))
      continue
    }
    if (msg.id != null && pending.has(msg.id)) {
      pending.get(msg.id)(msg)
      pending.delete(msg.id)
    }
  }
})
p.stderr.on('data', (d) => process.stderr.write(`  [srv] ${d.toString().trim()}\n`))

let seq = 0
function rpc(method, params) {
  const id = ++seq
  return new Promise((res, rej) => {
    pending.set(id, res)
    p.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n')
    setTimeout(() => rej(new Error(`timeout on ${method}`)), 600000)
  })
}

let pass = 0
let fail = 0
const failures = []

function preview(text, maxLines = 18) {
  return text
    .split('\n')
    .slice(0, maxLines)
    .map((l) => `      ${l}`)
    .join('\n')
}

async function call(name, args) {
  const r = await rpc('tools/call', { name, arguments: args })
  return {
    text: r.error ? `RPC ERROR: ${r.error.message}` : r.result?.content?.map((c) => c.text).join('\n') ?? '',
    bad: Boolean(r.error) || Boolean(r.result?.isError),
  }
}

async function check(title, name, args, opts = {}) {
  const { expectError = false, mustInclude } = opts
  const { text, bad } = await call(name, args)
  let ok = expectError ? bad : !bad
  if (ok && mustInclude && !text.includes(mustInclude)) ok = false
  console.log(`\n${ok ? 'PASS' : 'FAIL'}  ${title}`)
  console.log(preview(text))
  if (!ok) {
    fail++
    failures.push(title)
    if (bad) console.log(`      ^ unexpected error: ${text.split('\n')[0]}`)
    if (mustInclude && !text.includes(mustInclude)) console.log(`      ^ missing expected text: ${JSON.stringify(mustInclude)}`)
  } else pass++
  return text
}

console.log(`STATISTICA MCP selftest\nsource: ${SRC}\nwork copy: ${WORK}\n`)

const init = await rpc('initialize', {
  protocolVersion: '2024-11-05',
  capabilities: {},
  clientInfo: { name: 'selftest', version: '2.0.0' },
})
p.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n')

const srv = init.result?.serverInfo
console.log(`server: ${srv?.name} ${srv?.version}  protocol ${init.result?.protocolVersion}`)

const tools = await rpc('tools/list', {})
const toolNames = tools.result.tools.map((t) => t.name)
console.log(`tools:  ${toolNames.length} registered`)
console.log(`        ${toolNames.join(', ')}`)

// --- COM / file inspection ------------------------------------------------
await check('COM availability', 'statistica_info', {}, { mustInclude: 'STATISTICA COM is available' })
const desc = await check('open and describe', 'describe_spreadsheet', { path: WORK }, { mustInclude: 'variables' })
await check('list analysis modules', 'list_analysis_modules', {}, { mustInclude: 'Time Series' })
await check('read first rows', 'read_variables', { path: WORK, variables: [1], limit: 5 }, { mustInclude: 'variable 1' })
await check('error handling: missing file', 'describe_spreadsheet', { path: 'C:\\definitely\\not\\here.sta' }, { expectError: true })
await check('error handling: unknown variable', 'read_variables', { path: WORK, variables: ['no_such_var'] }, { expectError: true })

const nCases = Number(/Size: (\d+) cases/.exec(desc)?.[1] ?? 0)

// --- descriptive statistics ----------------------------------------------
await check('descriptives (JS)', 'descriptives', { path: WORK, variables: [1] }, { mustInclude: 'Descriptive statistics' })
await check('statistica_descriptives (engine)', 'statistica_descriptives', { path: WORK, variables: [1, 2] }, { mustInclude: 'Basic Statistics' })
await check('correlation matrix', 'statistica_correlation', { path: WORK, variables: [2, 3, 4] }, { mustInclude: 'CorrelationMatrix' })
await check('t-test (single)', 'statistica_t_test', { path: WORK, kind: 'single', variables: [2] }, { mustInclude: 't-value' })

// --- analysis engine ------------------------------------------------------
await check('describe_analysis', 'describe_analysis', { path: WORK, module: 1901 }, { mustInclude: 'Time Series' })
await check(
  'run_analysis (generic)',
  'run_analysis',
  {
    path: WORK,
    module: 1301,
    steps: [{ set: { Statistics: 0 } }, { run: true }, { set: { Variables: '1', Mean: true, ValidN: true } }, { result: 'Summary' }],
  },
  { mustInclude: 'Mean' },
)

// --- time series ----------------------------------------------------------
await check('time series: autocorrelation', 'statistica_time_series', { path: WORK, procedure: 'autocorrelation', variables: [2], lags: 6 }, { mustInclude: 'Autocorrelations' })
await check(
  'time series: arima',
  'statistica_time_series',
  { path: WORK, procedure: 'arima', variables: [2], arOrder: 1, maOrder: 0, difference: true, forecasts: 4 },
  { mustInclude: 'ForecastCases' },
)
await check('time series: spectral', 'statistica_time_series', { path: WORK, procedure: 'spectral', variables: [2] }, { mustInclude: 'Frequency' })
await check('time series: smoothing', 'statistica_time_series', { path: WORK, procedure: 'smoothing', variables: [2], window: 3 }, { mustInclude: 'SaveVariables' })
await check('t-test (dependent)', 'statistica_t_test', { path: WORK, kind: 'dependent', variables: [2, 3] }, { mustInclude: 'Confidence' })

// --- spreadsheet editing --------------------------------------------------
await check('add a variable', 'add_variables', { path: WORK, name: 'SELFTEST_VAR', longName: 'self test', after: 0, count: 1, type: 0 })
const col = Array.from({ length: nCases }, (_, i) => i * 1.5)
await check('write the new variable (by index)', 'write_variables', { path: WORK, columns: [{ index: 1, values: col }] })
await check('rename it', 'rename_variables', { path: WORK, renames: { 1: 'SELFTEST_DONE' } })
await check('resize', 'set_size', { path: WORK, variables: 18 })
await check('regression (dependent on SERIES_G)', 'statistica_regression', { path: WORK, dependent: 2, predictors: [1] }, { mustInclude: 'Coefficients' })

// --- io -------------------------------------------------------------------
await check('export csv', 'export_csv', { path: WORK, out: CSV })
console.log(`      csv on disk: ${existsSync(CSV)} (${existsSync(CSV) ? readFileSync(CSV).length : 0} bytes)`)
await check('save as .sta', 'save_spreadsheet', { path: WORK, out: SAVED, overwrite: true })
console.log(`      sta on disk: ${existsSync(SAVED)} (${existsSync(SAVED) ? readFileSync(SAVED).length : 0} bytes)`)
await check('reopen the saved copy', 'describe_spreadsheet', { path: SAVED }, { mustInclude: 'Size:' })
await check('import text file', 'import_data', { source: IMPORT_CSV, format: 'text', separator: ';' }, { mustInclude: 'variables' })
await check('delete the test variable', 'delete_variables', { path: SAVED, from: 1, to: 1 })

console.log(`\n${'='.repeat(60)}\nresult: ${pass} passed, ${fail} failed\n${'='.repeat(60)}`)
if (failures.length) {
  console.log('failed checks:')
  for (const f of failures) console.log(`  - ${f}`)
}

p.stdin.end()
for (const f of [WORK, CSV, SAVED, IMPORT_CSV]) rmSync(f, { force: true })
process.exit(fail > 0 ? 1 : 0)
