import { spawn } from 'node:child_process'
import { copyFileSync, existsSync, rmSync, readFileSync } from 'node:fs'
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
for (const f of [WORK, CSV, SAVED]) rmSync(f, { force: true })
copyFileSync(SRC, WORK)

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
    setTimeout(() => rej(new Error(`timeout on ${method}`)), 300000)
  })
}

let pass = 0
let fail = 0
async function check(title, name, args, expectError = false) {
  const r = await rpc('tools/call', { name, arguments: args })
  const text = r.result?.content?.map((c) => c.text).join('\n') ?? ''
  const bad = r.error || r.result?.isError
  const ok = expectError ? bad : !bad
  console.log(`\n${ok ? 'PASS' : 'FAIL'}  ${title}`)
  console.log(
    text
      .split('\n')
      .map((l) => `      ${l}`)
      .join('\n'),
  )
  if (ok) pass++
  else {
    fail++
    if (bad) console.log(`      ^ unexpected error: ${bad.message ?? '(isError)'}`)
  }
  return text
}

console.log(`STATISTICA MCP selftest\nsource: ${SRC}\nwork copy: ${WORK}\n`)

const init = await rpc('initialize', {
  protocolVersion: '2024-11-05',
  capabilities: {},
  clientInfo: { name: 'selftest', version: '1.0.0' },
})
p.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n')

const srv = init.result?.serverInfo
console.log(`server: ${srv?.name} ${srv?.version}  protocol ${init.result?.protocolVersion}`)

const tools = await rpc('tools/list', {})
console.log(`tools:  ${tools.result.tools.length} registered`)
console.log(`        ${tools.result.tools.map((t) => t.name).join(', ')}`)

await check('COM availability', 'statistica_info', {})
const desc = await check('open and describe', 'describe_spreadsheet', { path: WORK })
await check('descriptive statistics', 'descriptives', { path: WORK })
await check('read first rows', 'read_variables', { path: WORK, variables: [1], limit: 5 })
await check('add a variable', 'add_variables', { path: WORK, name: 'SELFTEST_VAR', longName: 'self test', after: 0, count: 1, type: 0 })

const nCases = Number(/Size: (\d+) cases/.exec(desc)?.[1] ?? 0)
const col = Array.from({ length: nCases }, (_, i) => i * 1.5)
await check('write the new variable (by index)', 'write_variables', { path: WORK, columns: [{ index: 1, values: col }] })
await check('rename it', 'rename_variables', { path: WORK, renames: { 1: 'SELFTEST_DONE' } })
await check('resize', 'set_size', { path: WORK, variables: 11 })
await check('export csv', 'export_csv', { path: WORK, out: CSV })
console.log(`      csv on disk: ${existsSync(CSV)} (${existsSync(CSV) ? readFileSync(CSV).length : 0} bytes)`)
await check('save as .sta', 'save_spreadsheet', { path: WORK, out: SAVED, overwrite: true })
console.log(`      sta on disk: ${existsSync(SAVED)} (${existsSync(SAVED) ? readFileSync(SAVED).length : 0} bytes)`)
await check('reopen the saved copy', 'describe_spreadsheet', { path: SAVED })
await check('delete the test variable', 'delete_variables', { path: SAVED, from: 1, to: 1 })
await check('error handling: missing file', 'describe_spreadsheet', { path: 'C:\\definitely\\not\\here.sta' }, true)
await check('error handling: unknown variable', 'read_variables', { path: WORK, variables: ['no_such_var'] }, true)

console.log(`\n${'='.repeat(60)}\nresult: ${pass} passed, ${fail} failed\n${'='.repeat(60)}`)

p.stdin.end()
for (const f of [WORK, CSV, SAVED]) rmSync(f, { force: true })
process.exit(fail > 0 ? 1 : 0)