import { runWorker } from '../worker.mjs'
import { ANALYSIS_MODULES, ENUMS, resolveModule } from '../modules.mjs'
import { requirePath } from '../util.mjs'
import { fmt, numify } from '../format.mjs'
import { TYPE_NAME } from '../constants.mjs'

const tools = [
  {
    name: 'statistica_info',
    description:
      'Report STATISTICA COM availability, version and executable path. Run this first if STATISTICA operations fail.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'list_analysis_modules',
    description:
      'List every analysis procedure exposed by STATISTICA (id + name) that can be driven with run_analysis.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'describe_spreadsheet',
    description:
      'Open a .sta/.stw file and report case count, variable count and every variable (index, short name, clean name, long name/formula, type, measurement level, missing code).',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to a .sta or .stw file.' },
        sheet: {
          type: ['string', 'integer'],
          description: 'Sheet name or 1-based index inside the file. Omit to use the first sheet.',
        },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'read_variables',
    description:
      'Read variable values. Numeric columns are read vectorized; text columns as strings. Missing values are returned as null.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to a .sta or .stw file.' },
        sheet: { type: ['string', 'integer'], description: 'Sheet name or 1-based index.' },
        variables: {
          type: 'array',
          items: { type: ['string', 'integer'] },
          description: 'Variable names (short or clean) or 1-based indices. Omit to read every variable.',
        },
        offset: { type: 'integer', minimum: 1, description: 'First case to read (1-based). Default 1.' },
        limit: { type: 'integer', minimum: 1, description: 'Maximum number of cases to read.' },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'describe_analysis',
    description:
      'Introspect a STATISTICA analysis dialog before driving it: lists every settable property and callable method, plus known enum constants. Use it to build a run_analysis request.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to a .sta/.stw file used to instantiate the analysis.' },
        sheet: { type: ['string', 'integer'] },
        module: { type: ['string', 'integer'], description: 'Module id or name (see list_analysis_modules).' },
      },
      required: ['path', 'module'],
      additionalProperties: false,
    },
  },
]

const handlers = {
  async statistica_info() {
    const r = await runWorker({ cmd: 'info' })
    return [`STATISTICA COM is available.`, `version: ${r.version} (${r.versionEx ?? ''})`, `exe: ${r.exe}`, `pid: ${r.pid}`].join('\n')
  },

  async list_analysis_modules() {
    const lines = ['STATISTICA analysis modules (use the id or the name with run_analysis):', '']
    for (const [id, nm] of ANALYSIS_MODULES) lines.push(`${String(id).padEnd(6)}${nm}`)
    return lines.join('\n')
  },

  async describe_spreadsheet(a) {
    const r = await runWorker({ cmd: 'describe', path: requirePath(a), sheet: a.sheet, attach: a.attach })
    const lines = [`File: ${a.path}`, `Sheet: ${r.sheetName} (index ${r.sheetIndex})`, `Size: ${r.cases} cases x ${r.variables} variables`]
    if (Array.isArray(r.sheets) && r.sheets.length > 1) lines.push(`Sheets in file: ${r.sheets.join(', ')}`)
    lines.push('', 'idx  name                       type     len  measurement  missing      long name / formula', '---  --------------------------  -------  ---  -----------  -----------  --------------------')
    for (const v of r.varInfo) {
      const md = v.missingValue === null || v.missingValue === undefined ? '' : fmt(v.missingValue)
      lines.push(
        `${String(v.index).padEnd(4)}${(v.cleanName || v.name).slice(0, 26).padEnd(28)}${TYPE_NAME[v.type].padEnd(9)}${String(v.typeLength).padEnd(4)}${String(v.measurementType).padEnd(13)}${String(md).padEnd(13)}${v.longName}`,
      )
    }
    return lines.join('\n')
  },

  async read_variables(a) {
    const r = await runWorker({ cmd: 'read', path: requirePath(a), sheet: a.sheet, variables: a.variables, offset: a.offset, limit: a.limit, attach: a.attach })
    const lines = []
    for (const d of r.data) {
      const isText = d.type === 1
      const vals = d.values.map((v) => numify(v, isText))
      lines.push(`variable ${d.index}: ${d.cleanName || d.name}  [${TYPE_NAME[d.type]}, ${d.count} values from case ${d.offset}]`)
      if (d.longName && d.longName !== d.name) lines.push(`  long name: ${d.longName}`)
      const preview = vals.slice(0, 16).map((v) => (isText ? JSON.stringify(v) : fmt(v)))
      lines.push(`  first: ${preview.join(', ')}`)
      if (vals.length > 16) {
        const tail = vals.slice(-4).map((v) => (isText ? JSON.stringify(v) : fmt(v)))
        lines.push(`  last:  ${tail.join(', ')}`)
      }
      const miss = vals.filter((v) => v === null).length
      lines.push(`  missing: ${miss}`)
    }
    return lines.join('\n')
  },

  async describe_analysis(a) {
    const moduleId = resolveModule(a.module)
    const r = await runWorker({ cmd: 'describe_analysis', path: a.path, sheet: a.sheet, module: moduleId })
    const props = r.members.filter((m) => m.kind === 'Property').map((m) => m.name)
    const methods = r.members.filter((m) => m.kind === 'Method').map((m) => m.name)
    const lines = [`Analysis: ${r.name} (module ${r.module})`, '']
    lines.push(`Settable properties (${props.length}):`)
    lines.push('  ' + props.join(', '))
    lines.push('', `Callable methods (${methods.length}):`)
    lines.push('  ' + methods.join(', '))
    const enums = ENUMS[moduleId]
    if (enums) {
      lines.push('', 'Known enum constants:')
      for (const [enumName, vals] of Object.entries(enums)) {
        lines.push(`  ${enumName}:`)
        for (const [k, v] of Object.entries(vals)) lines.push(`    ${k} = ${v}`)
      }
    }
    return lines.join('\n')
  },
}

export default { tools, handlers }
