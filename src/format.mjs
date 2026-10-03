export function fmt(v) {
  if (v === null || v === undefined) return '.'
  if (typeof v !== 'number') {
    const n = Number(v)
    if (v !== '' && !Number.isNaN(n) && String(v).trim() !== '') return fmt(n)
    return String(v)
  }
  if (!Number.isFinite(v)) return String(v)
  if (v !== 0 && Math.abs(v) < 1e-4) return v.toExponential(4)
  return String(Number(v.toPrecision(10)))
}

export function numify(v, isText) {
  if (v === null || v === undefined) return null
  if (isText) return String(v)
  const n = Number(v)
  return Number.isNaN(n) ? null : n
}

export function varSpec(variables) {
  if (variables === undefined || variables === null) return undefined
  if (Array.isArray(variables)) {
    if (variables.length === 0) return undefined
    return variables.join(' ')
  }
  return String(variables)
}

function headerOf(n) {
  const clean = n.cleanName ?? n.name ?? ''
  const long = n.longName ?? ''
  if (long && long !== clean) return `${clean} ${long}`.trim()
  return clean
}

function renderTable(tbl, maxRows = 60, maxCols = 32) {
  if (!tbl || tbl.kind !== 'table') return String(tbl)
  const nv = tbl.variables
  const nc = tbl.cases
  const cols = Math.min(nv, maxCols)
  const rows = Math.min(nc, maxRows)
  const headers = []
  for (let i = 0; i < cols; i++) headers.push(headerOf(tbl.names[i] ?? {}))
  const body = []
  const widths = headers.map((h) => Math.min(24, h.length))
  for (let r = 0; r < rows; r++) {
    const line = []
    for (let c = 0; c < cols; c++) {
      const cell = fmt(tbl.columns?.[c]?.[r])
      line.push(cell)
      if (cell.length > widths[c]) widths[c] = Math.min(24, cell.length)
    }
    body.push(line)
  }
  const out = []
  out.push(headers.map((h, i) => h.padEnd(widths[i])).join('  ').trimEnd())
  out.push(widths.map((w) => '-'.repeat(w)).join('  ').trimEnd())
  for (const line of body) out.push(line.map((x, i) => x.padEnd(widths[i])).join('  ').trimEnd())
  if (nc > rows) out.push(`... ${nc - rows} more case(s)`)
  if (nv > cols) out.push(`... ${nv - cols} more variable(s)`)
  return out.join('\n')
}

export function renderResult(res, indent = '') {
  if (res === null || res === undefined) return `${indent}(no value)`
  if (typeof res !== 'object') return indent + String(res)
  if (res.kind === 'table') return indent + renderTable(res).replace(/\n/g, `\n${indent}`)
  if (res.kind === 'array' || res.kind === 'collection') {
    const items = res.items ?? []
    const out = []
    if (res.kind === 'collection' && res.count > items.length) out.push(`${indent}(showing ${items.length} of ${res.count})`)
    items.forEach((it, i) => {
      if (it && typeof it === 'object' && it.kind === 'document') {
        out.push(`${indent}[${i}] ${it.kind}: ${it.name ?? it.type ?? ''}`)
      } else {
        out.push(`${indent}[${i}]`)
        out.push(renderResult(it, indent + '    '))
      }
    })
    return out.join('\n')
  }
  if (res.kind === 'document') return `${indent}${res.name ? `${res.name} ` : ''}(${res.type ?? 'document'})`
  return `${indent}${JSON.stringify(res)}`
}

export function formatAnalysis(r) {
  const lines = [`Analysis: ${r.name} (module ${r.module})`]
  if (r.steps && typeof r.steps === 'object') {
    const runs = Object.entries(r.steps).filter(([k]) => k.startsWith('run'))
    if (runs.length) lines.push(`Runs: ${runs.map(([, v]) => v).join(', ')}`)
  }
  if (Array.isArray(r.warnings) && r.warnings.length) {
    lines.push('Warnings:')
    for (const w of r.warnings) lines.push(`  ! ${w}`)
  }
  for (const [key, val] of Object.entries(r.results ?? {})) {
    lines.push('', `--- ${key} ---`)
    lines.push(renderResult(val))
  }
  if (r.graphs && Object.keys(r.graphs).length) {
    lines.push('', '--- saved graphs ---')
    for (const [k, v] of Object.entries(r.graphs)) lines.push(`  ${k}: ${v.out} (${v.bytes} bytes)`)
  }
  return lines.join('\n')
}
