#!/usr/bin/env node
// Zero-dependency source checker for this repository.
//
//   node scripts/check.mjs            parse, imports, style
//   node scripts/check.mjs --fix      additionally fix CRLF / trailing spaces / final newline
//   node scripts/check.mjs --strict   warnings (tabs, long lines) fail too
//   node scripts/check.mjs --external also run prettier / eslint / PSScriptAnalyzer when present
//
// Exit code is 1 when errors are found (or warnings with --strict).
import { readFileSync, writeFileSync, existsSync, statSync, readdirSync, mkdtempSync, rmSync } from 'node:fs'
import { join, relative, resolve, dirname } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { spawnSync } from 'node:child_process'
import { tmpdir } from 'node:os'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const SKIP_DIRS = new Set(['.git', 'node_modules'])
const FLAGS = new Set(process.argv.slice(2).filter((a) => a.startsWith('--')))
const OPT = {
  fix: FLAGS.has('--fix'),
  strict: FLAGS.has('--strict'),
  external: FLAGS.has('--external'),
  quiet: FLAGS.has('--quiet'),
}
const KNOWN_FLAGS = new Set(['--fix', '--strict', '--external', '--quiet'])
for (const f of FLAGS) if (!KNOWN_FLAGS.has(f)) {
  console.error(`unknown flag: ${f}`)
  process.exit(2)
}

const MAX_REPEAT = 5
const problems = []
const repeats = new Map()

function rel(file) {
  return relative(ROOT, file).replace(/\\/g, '/')
}
function add(level, rule, file, line, message) {
  const key = `${rule}\u0000${rel(file)}`
  const count = (repeats.get(key) ?? 0) + 1
  repeats.set(key, count)
  if (count > MAX_REPEAT) return
  problems.push({ level, rule, file: rel(file), line, message, more: count === MAX_REPEAT })
}

function walk(dir, out) {
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    if (e.isDirectory()) {
      if (SKIP_DIRS.has(e.name)) continue
      walk(join(dir, e.name), out)
    } else {
      out.push(join(dir, e.name))
    }
  }
  return out
}

function lineAt(src, index) {
  let line = 1
  for (let i = 0; i < index && i < src.length; i++) if (src[i] === '\n') line++
  return line
}

// --- syntax ---------------------------------------------------------------
function checkJsSyntax(files) {
  for (const f of files) {
    const r = spawnSync(process.execPath, ['--check', f], { encoding: 'utf8' })
    if (r.error) {
      add('error', 'js-runner', f, 0, `cannot run node --check: ${r.error.message}`)
      continue
    }
    if (r.status !== 0) {
      const msg = (r.stderr || '').trim().split('\n').slice(0, 3).join(' ').replace(/\s+/g, ' ')
      add('error', 'js-syntax', f, 0, msg)
    }
  }
}

const IMPORT_RE = /(?:^|[^\w$])from\s+['"](\.[^'"]+)['"]|(?:^|[^\w$])import\s+['"](\.[^'"]+)['"]/gm
function resolveRelative(abs) {
  const tries = [abs, `${abs}.mjs`, `${abs}.js`, join(abs, 'index.mjs'), join(abs, 'index.js')]
  for (const p of tries) {
    try {
      if (existsSync(p) && statSync(p).isFile()) return p
    } catch {}
  }
  return null
}
function checkImports(files) {
  for (const f of files) {
    const src = readFileSync(f, 'utf8')
    let m
    IMPORT_RE.lastIndex = 0
    while ((m = IMPORT_RE.exec(src))) {
      const spec = m[1] || m[2]
      if (!resolveRelative(resolve(dirname(f), spec))) {
        add('error', 'import-missing', f, lineAt(src, m.index), `cannot resolve import '${spec}'`)
      }
    }
  }
}

function checkPowerShellSyntax(files) {
  if (files.length === 0) return
  const dir = mkdtempSync(join(tmpdir(), 'sta-check-'))
  try {
    const list = join(dir, 'list.txt')
    writeFileSync(list, files.join('\n'), 'utf8')
    const script = join(dir, 'parse.ps1')
    writeFileSync(
      script,
      `param([string]$ListFile)
$ErrorActionPreference = 'Stop'
$paths = Get-Content -LiteralPath $ListFile -Encoding UTF8
$errs = @()
foreach ($p in $paths) {
  if (-not (Test-Path -LiteralPath $p)) { continue }
  $e = $null; $t = $null
  $null = [System.Management.Automation.Language.Parser]::ParseFile($p, [ref]$t, [ref]$e)
  if ($e) { foreach ($x in $e) { $errs += [pscustomobject]@{ file = $p; line = $x.Extent.StartLineNumber; col = $x.Extent.StartColumnNumber; message = $x.Message } } }
}
ConvertTo-Json -InputObject $errs -Compress -Depth 4
`,
      'utf8',
    )
    const r = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-ListFile', list], { encoding: 'utf8' })
    if (r.error || r.status !== 0) {
      add('error', 'ps-runner', ROOT, 0, `PowerShell parser failed: ${String(r.stderr || r.error?.message || '').trim().slice(0, 300)}`)
      return
    }
    const out = (r.stdout || '').trim()
    if (!out || out === '[]') return
    let parsed
    try {
      parsed = JSON.parse(out)
    } catch {
      add('error', 'ps-parse-output', ROOT, 0, 'cannot parse PowerShell parser output')
      return
    }
    for (const e of Array.isArray(parsed) ? parsed : [parsed]) {
      add('error', 'ps-syntax', e.file, e.line, `${e.message} (col ${e.col})`)
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

const DOTSOURCE_RE = /Join-Path\s+\$PSScriptRoot\s+['"]([^'"]+)['"]/g
function checkDotSources(files) {
  for (const f of files) {
    const src = readFileSync(f, 'utf8')
    let m
    DOTSOURCE_RE.lastIndex = 0
    while ((m = DOTSOURCE_RE.exec(src))) {
      const target = resolve(dirname(f), m[1].replace(/\\/g, '/'))
      if (!existsSync(target)) {
        add('error', 'dot-source-missing', f, lineAt(src, m.index), `cannot resolve dot-source '${m[1]}'`)
      }
    }
  }
}

function checkJson(files) {
  for (const f of files) {
    try {
      JSON.parse(readFileSync(f, 'utf8'))
    } catch (e) {
      add('error', 'json-parse', f, 0, e.message)
    }
  }
}

// --- style ----------------------------------------------------------------
function checkStyle(files) {
  for (const f of files) {
    const raw = readFileSync(f, 'utf8')
    const lines = raw.split('\n')
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i].replace(/\r$/, '')
      if (/[ \t]+$/.test(ln)) add('error', 'trailing-whitespace', f, i + 1, 'trailing whitespace')
      if (/\t/.test(ln)) add('warning', 'tab-indent', f, i + 1, 'tab character in indentation')
      if (ln.length > 300) add('warning', 'line-length', f, i + 1, `line length ${ln.length} > 300`)
    }
    if (/\r\n/.test(raw)) add('error', 'crlf', f, 0, 'CRLF line endings (expected LF)')
    if (raw.length > 0 && !raw.endsWith('\n')) add('error', 'final-newline', f, lines.length, 'missing final newline')
    if (/\n\n+$/.test(raw)) add('warning', 'blank-eof', f, lines.length, 'more than one trailing newline')
  }
}

function applyFixes(files) {
  let changed = 0
  for (const f of files) {
    const src = readFileSync(f, 'utf8')
    let out = src.replace(/\r\n/g, '\n')
    out = out.replace(/[ \t]+\n/g, '\n')
    out = out.replace(/\n+$/, '\n')
    if (out !== src) {
      writeFileSync(f, out, 'utf8')
      changed++
    }
  }
  return changed
}

// --- tool wiring (dynamic import of the MCP server modules) ---------------
async function checkToolWiring() {
  const base = (s) => pathToFileURL(join(ROOT, 'src', 'tools', s)).href
  const groupNames = ['inspect', 'edit', 'io', 'stats', 'engine']
  let groups
  let index
  try {
    groups = await Promise.all(groupNames.map((n) => import(base(`${n}.mjs`))))
    index = await import(base('index.mjs'))
  } catch (e) {
    add('error', 'tool-import', ROOT, 0, `cannot import tool modules: ${e.message}`)
    return
  }
  const handlers = Object.assign({}, ...groups.map((g) => g.default?.handlers ?? {}))
  const names = new Set()
  for (const t of index.TOOLS) {
    if (!t.name) {
      add('error', 'tool-shape', 'src/tools/index.mjs', 0, 'tool without a name')
      continue
    }
    if (names.has(t.name)) add('error', 'tool-duplicate', 'src/tools/index.mjs', 0, `duplicate tool '${t.name}'`)
    names.add(t.name)
    if (typeof handlers[t.name] !== 'function') add('error', 'tool-handler-missing', 'src/tools/index.mjs', 0, `no handler for tool '${t.name}'`)
    if (!t.description || !t.inputSchema) add('warning', 'tool-shape', 'src/tools/index.mjs', 0, `tool '${t.name}' is missing description or inputSchema`)
  }
  for (const h of Object.keys(handlers)) {
    if (!names.has(h)) add('error', 'handler-orphan', 'src/tools/index.mjs', 0, `handler '${h}' has no matching tool`)
  }
  if (index.TOOLS.length === 0) add('error', 'tools-empty', 'src/tools/index.mjs', 0, 'no tools registered')
}

// --- optional external linters --------------------------------------------
function runExternal(cmd, args) {
  return spawnSync(cmd, args, { encoding: 'utf8', shell: process.platform === 'win32' })
}
function checkExternal(jsFiles, psFiles) {
  const prettier = runExternal('prettier', ['--version'])
  if (!prettier.error && prettier.status === 0) {
    const r = runExternal('prettier', ['--check', ...jsFiles.map((f) => rel(f))])
    if (r.status !== 0) add('error', 'prettier', ROOT, 0, 'prettier --check reported formatting issues')
  } else if (!OPT.quiet) {
    console.log('note: prettier not found, skipping (npm i -D prettier)')
  }
  const eslint = runExternal('eslint', ['--version'])
  if (!eslint.error && eslint.status === 0) {
    const r = runExternal('eslint', jsFiles.map((f) => rel(f)))
    if (r.status !== 0) add('error', 'eslint', ROOT, 0, 'eslint reported problems')
  } else if (!OPT.quiet) {
    console.log('note: eslint not found, skipping (npm i -D eslint)')
  }
  const pssa = runExternal('powershell.exe', ['-NoProfile', '-Command', 'if (Get-Module -ListAvailable PSScriptAnalyzer) { exit 0 } else { exit 3 }'])
  if (!pssa.error && pssa.status === 0) {
    const r = runExternal('powershell.exe', ['-NoProfile', '-Command', `$r = Invoke-ScriptAnalyzer -Path '${ROOT}' -Recurse -Severity Error,Warning; if ($r) { $r | ForEach-Object { Write-Output ($_.ScriptName + ':' + $_.Line + ' [' + $_.RuleName + '] ' + $_.Message) }; exit 1 }`])
    if (r.status === 1) {
      for (const line of (r.stdout || '').trim().split('\n')) add('error', 'psscriptanalyzer', ROOT, 0, line.trim())
    }
  } else if (!OPT.quiet) {
    console.log('note: PSScriptAnalyzer not found, skipping (Install-Module PSScriptAnalyzer)')
  }
}

// --- run ------------------------------------------------------------------
const files = walk(ROOT, [])
const jsFiles = files.filter((f) => f.endsWith('.mjs'))
const psFiles = files.filter((f) => f.endsWith('.ps1'))
const jsonFiles = files.filter((f) => f.endsWith('.json'))
const styleFiles = [...jsFiles, ...psFiles]

if (OPT.fix) {
  const changed = applyFixes(styleFiles)
  if (!OPT.quiet) console.log(`fixed ${changed} file(s)`)
}

checkJsSyntax(jsFiles)
checkImports(jsFiles)
checkPowerShellSyntax(psFiles)
checkDotSources(psFiles)
checkJson(jsonFiles)
checkStyle(styleFiles)
await checkToolWiring()
if (OPT.external) checkExternal(jsFiles, psFiles)

// --- report ---------------------------------------------------------------
const order = { error: 0, warning: 1 }
problems.sort((a, b) => a.file.localeCompare(b.file) || a.line - b.line || order[a.level] - order[b.level])
let current = null
for (const p of problems) {
  if (p.file !== current) {
    current = p.file
    console.log(`\n${p.file}`)
  }
  const loc = p.line ? `${p.line}`.padStart(5) : '    -'
  console.log(`  ${loc}  ${p.level === 'error' ? 'ERROR' : 'WARN '}  [${p.rule}] ${p.message}`)
  if (p.more) console.log(`         ... (more ${p.rule} not shown)`)
}

const errors = problems.filter((p) => p.level === 'error').length
const warnings = problems.filter((p) => p.level === 'warning').length
const summary = `checked ${files.length} files (${jsFiles.length} js, ${psFiles.length} ps1, ${jsonFiles.length} json): ${errors} error(s), ${warnings} warning(s)`
console.log(`\n${summary}`)
if (errors === 0 && warnings === 0) console.log('OK')
process.exit(OPT.strict ? (errors + warnings > 0 ? 1 : 0) : errors > 0 ? 1 : 0)
