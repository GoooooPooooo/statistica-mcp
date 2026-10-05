import { runWorker } from '../worker.mjs'
import { requirePath } from '../util.mjs'
import { fmt } from '../format.mjs'

const tools = [
  {
    name: 'write_variables',
    description:
      'Overwrite whole variables in a spreadsheet. Numeric columns require exactly one value per case; missing values may be null. Text columns accept per-row strings (shorter arrays leave the rest untouched). Changes are only persisted if `save` is given.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file (it is not modified unless `save` is set).' },
        sheet: { type: ['string', 'integer'], description: 'Sheet name or 1-based index.' },
        columns: {
          type: 'array',
          description: 'Columns to overwrite.',
          items: {
            type: 'object',
            properties: {
              index: { type: 'integer', description: '1-based variable index.' },
              name: { type: 'string', description: 'Variable name (alternative to index).' },
              values: { type: 'array', items: {} },
            },
            additionalProperties: false,
          },
        },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'columns'],
      additionalProperties: false,
    },
  },
  {
    name: 'set_formula',
    description:
      'Assign a formula to a variable (STATISTICA keeps it in the variable long name) and recompute it, e.g. formula "=v9*v10". The result values are returned for inspection. Use attach=true to write into the running STATISTICA window.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variable: { type: ['string', 'integer'], description: 'Target variable (name or 1-based index).' },
        formula: { type: 'string', description: 'Formula without or with a leading "=", e.g. "v9*v10".' },
        recalculate: { type: 'boolean', description: 'Recompute the variable after assignment. Default true.' },
        save: { type: 'string', description: 'Optional destination .sta path.' },
      },
      required: ['path', 'variable', 'formula'],
      additionalProperties: false,
    },
  },
  {
    name: 'add_variables',
    description: 'Append new empty variables to a spreadsheet. type: 0=numeric, 1=text, 2=integer, 3=byte.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        name: { type: 'string', description: 'Short name for the new variable(s).' },
        longName: { type: 'string', description: 'Long/description name.' },
        after: { type: 'integer', minimum: 0, description: 'Insert after this 1-based variable index (0 — before the first variable). Omit to append at the end.' },
        count: { type: 'integer', minimum: 1, description: 'How many variables to add. Default 1.' },
        type: { type: 'integer', enum: [0, 1, 2, 3], description: 'Variable type. Default 0 (numeric).' },
        typeLength: { type: 'integer', description: 'Declared length for text variables.' },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'name'],
      additionalProperties: false,
    },
  },
  {
    name: 'rename_variables',
    description: 'Rename variables and/or change their long names, keyed by current variable name or index.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        renames: { type: 'object', description: 'Map of old name or index -> new short name.', additionalProperties: true },
        longNames: { type: 'object', description: 'Map of current name or index -> new long name.', additionalProperties: true },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'renames'],
      additionalProperties: false,
    },
  },
  {
    name: 'delete_variables',
    description: 'Delete an inclusive range of variables by index or name.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        from: { type: ['string', 'integer'], description: 'First variable to delete.' },
        to: { type: ['string', 'integer'], description: 'Last variable to delete.' },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'from', 'to'],
      additionalProperties: false,
    },
  },
  {
    name: 'set_size',
    description: 'Resize a spreadsheet. Growing adds empty cases/variables; shrinking discards the excess.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        cases: { type: 'integer', minimum: 1 },
        variables: { type: 'integer', minimum: 1 },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'case_names',
    description:
      'Read (and optionally set) case names / text labels for the rows of a spreadsheet. When `names` is given they are written (one per case, starting at case 1).',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        names: { type: 'array', items: { type: 'string' }, description: 'Case names to assign (case 1, 2, ...). Omit to only read.' },
        limit: { type: 'integer', minimum: 1, description: 'How many names to return. Default 200, all when reading named cases.' },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'sort_data',
    description:
      'Sort a spreadsheet in place by one or more keys (case names move with their rows). `order` may be a single value or one per key: 0/asc or 1/desc.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] }, description: 'Sort key(s), most significant first.' },
        order: {
          type: ['integer', 'string', 'array'],
          description: 'Ascending (0/"asc") or descending (1/"desc"). A single value or one per key.',
        },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'variables'],
      additionalProperties: false,
    },
  },
  {
    name: 'select_cases',
    description:
      'Keep only the rows matching a condition and drop the rest (all columns are rewritten). Supports numeric and text comparisons; missing values can be matched with op "missing".',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variable: { type: ['string', 'integer'], description: 'Variable the condition is evaluated on.' },
        op: {
          type: 'string',
          enum: ['gt', 'ge', 'lt', 'le', 'eq', 'ne', 'in', 'notin', 'missing', 'notmissing'],
          description: 'Comparison operator. Default eq.',
        },
        value: { type: ['string', 'number', 'boolean'], description: 'Comparison value for gt/ge/lt/le/eq/ne.' },
        values: { type: 'array', items: { type: ['string', 'number'] }, description: 'Value list for in/notin.' },
        keepCaseNames: { type: 'boolean', description: 'Carry case names to the surviving rows. Default true.' },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'variable', 'op'],
      additionalProperties: false,
    },
  },
  {
    name: 'recode',
    description:
      'Recode the values of one variable using a mapping (old value -> new value), optionally sending every unlisted value to `default`. Missing values are preserved unless `missing` is given.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variable: { type: ['string', 'integer'], description: 'Variable to recode.' },
        map: { type: 'object', description: 'Mapping old value -> new value (keys are compared as strings).', additionalProperties: true },
        default: { type: ['string', 'number'], description: 'Value assigned to every unlisted, non-missing cell.' },
        missing: { type: ['string', 'number'], description: 'Replacement for missing cells (default: leave missing).' },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'variable', 'map'],
      additionalProperties: false,
    },
  },
  {
    name: 'set_measurement',
    description:
      'Set the measurement level of a variable (auto/continuous/categorical/ordinal). STATISTICA uses it to treat the variable as a factor or a covariate in analysis.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variable: { type: ['string', 'integer'], description: 'Variable name or 1-based index.' },
        type: {
          type: 'string',
          enum: ['unspecified', 'auto', 'continuous', 'categorical', 'ordinal'],
          description: 'Measurement level.',
        },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'variable', 'type'],
      additionalProperties: false,
    },
  },
  {
    name: 'value_labels',
    description:
      'Attach or clear text labels for the numeric values of a variable. `labels` maps a numeric value to a label string or an object {label, description}; `clear` removes existing labels first.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variable: { type: ['string', 'integer'], description: 'Variable name or 1-based index.' },
        labels: { type: 'object', additionalProperties: true, description: 'Map of numeric value -> label string or {label, description}.' },
        clear: { type: 'boolean', description: 'Remove existing value labels before applying. Default false.' },
        save: { type: 'string', description: 'Optional destination path to persist the result as .sta.' },
      },
      required: ['path', 'variable'],
      additionalProperties: false,
    },
  },
]

const handlers = {
  async write_variables(a) {
    if (!Array.isArray(a.columns) || a.columns.length === 0) throw new Error('`columns` must be a non-empty array')
    const r = await runWorker({ cmd: 'write', path: requirePath(a), sheet: a.sheet, columns: a.columns, save: a.save, attach: a.attach })
    const lines = [`Wrote ${r.written.length} variable(s); spreadsheet now ${r.cases} cases`]
    for (const w of r.written) lines.push(`  #${w.index} ${w.name} (${w.type}): ${w.written} value(s)`)
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },

  async set_formula(a) {
    const r = await runWorker({
      cmd: 'formula',
      path: requirePath(a),
      sheet: a.sheet,
      variable: a.variable,
      formula: a.formula,
      recalculate: a.recalculate,
      save: a.save,
      attach: a.attach,
    })
    const vals = (r.values ?? []).map((v) => fmt(v)).join(', ')
    return [`Set formula on #${r.index} ${r.name}: ${r.longName}`, `first values: ${vals}`].join('\n')
  },

  async add_variables(a) {
    const r = await runWorker({
      cmd: 'add_variables',
      path: requirePath(a),
      sheet: a.sheet,
      name: a.name,
      longName: a.longName,
      after: a.after,
      count: a.count,
      type: a.type,
      typeLength: a.typeLength,
      save: a.save,
      attach: a.attach,
    })
    const lines = [`Now ${r.variables} variables`]
    for (const v of r.added) lines.push(`  added #${v.index} "${v.name}"`)
    if (a.save) lines.push(`Saved to ${a.save}`)
    return lines.join('\n')
  },

  async rename_variables(a) {
    const r = await runWorker({ cmd: 'rename', path: requirePath(a), sheet: a.sheet, renames: a.renames, longNames: a.longNames, save: a.save, attach: a.attach })
    const lines = [`Renamed ${r.renamed.length} variable(s)`]
    for (const d of r.renamed) lines.push(`  #${d.index}: "${d.from}" -> "${d.to}"`)
    if (a.save) lines.push(`Saved to ${a.save}`)
    return lines.join('\n')
  },

  async delete_variables(a) {
    const r = await runWorker({ cmd: 'delete_variables', path: requirePath(a), sheet: a.sheet, from: a.from, to: a.to, save: a.save, attach: a.attach })
    const lines = [`Deleted variables ${r.deletedFrom}..${r.deletedTo}; now ${r.variables} variables`]
    if (a.save) lines.push(`Saved to ${a.save}`)
    return lines.join('\n')
  },

  async set_size(a) {
    const r = await runWorker({ cmd: 'set_size', path: requirePath(a), sheet: a.sheet, cases: a.cases, variables: a.variables, save: a.save, attach: a.attach })
    const lines = [`Resized to ${r.cases} cases x ${r.variables} variables`]
    if (a.save) lines.push(`Saved to ${a.save}`)
    return lines.join('\n')
  },

  async case_names(a) {
    const r = await runWorker({ cmd: 'case_names', path: requirePath(a), sheet: a.sheet, names: a.names, limit: a.limit, save: a.save, attach: a.attach })
    const lines = [`Cases: ${r.cases}, named: ${r.named}`]
    const preview = (r.names ?? []).filter((x) => x !== null && x !== '').slice(0, 20)
    if (preview.length) lines.push(`names: ${preview.map((x) => JSON.stringify(x)).join(', ')}`)
    if (a.names) lines.push(`assigned ${Array.isArray(a.names) ? a.names.length : 0} case name(s)`)
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },

  async sort_data(a) {
    if (!Array.isArray(a.variables) || a.variables.length === 0) throw new Error('`variables` must be a non-empty array')
    const r = await runWorker({ cmd: 'sort', path: requirePath(a), sheet: a.sheet, variables: a.variables, order: a.order, save: a.save, attach: a.attach })
    const lines = [`Sorted ${r.cases} cases x ${r.variables} variables by ${r.sortedBy.map((x) => JSON.stringify(x)).join(', ')}`]
    if (r.firstCaseNames?.length) lines.push(`first cases: ${r.firstCaseNames.map((x) => JSON.stringify(x)).join(', ')}`)
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },

  async select_cases(a) {
    const read = await runWorker({ cmd: 'read', path: requirePath(a), sheet: a.sheet, variables: [a.variable], attach: a.attach })
    const d = read.data?.[0]
    if (!d) throw new Error(`could not read variable ${a.variable}`)
    const isText = d.type === 1
    const op = a.op ?? 'eq'
    const values = d.values.map((v) => (isText ? (v === null ? null : String(v)) : v === null ? null : Number(v)))
    const num = (x) => {
      if (x === null || x === undefined) return null
      const n = Number(x)
      return Number.isNaN(n) ? null : n
    }
    const target = num(a.value)
    const list = Array.isArray(a.values) ? a.values.map((x) => (isText ? String(x) : num(x))) : []
    const keep = []
    for (let i = 0; i < values.length; i++) {
      const v = values[i]
      let ok = false
      switch (op) {
        case 'missing':
          ok = v === null
          break
        case 'notmissing':
          ok = v !== null
          break
        case 'in':
          ok = v !== null && list.some((t) => t === v)
          break
        case 'notin':
          ok = v !== null && !list.some((t) => t === v)
          break
        default: {
          if (v === null) break
          const cv = isText ? String(a.value) : target
          if (cv === null) break
          switch (op) {
            case 'gt':
              ok = v > cv
              break
            case 'ge':
              ok = v >= cv
              break
            case 'lt':
              ok = v < cv
              break
            case 'le':
              ok = v <= cv
              break
            case 'ne':
              ok = v !== cv
              break
            case 'eq':
            default:
              ok = v === cv
          }
        }
      }
      if (ok) keep.push(i + 1)
    }
    if (keep.length === 0) throw new Error('select_cases matched 0 cases')
    const r = await runWorker({ cmd: 'select', path: requirePath(a), sheet: a.sheet, cases: keep, caseNames: a.keepCaseNames, save: a.save, attach: a.attach })
    const lines = [`Kept ${r.kept} of ${r.cases + r.removed} cases (${r.removed} removed); now ${r.cases} cases x ${r.variables} variables`]
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },

  async recode(a) {
    const read = await runWorker({ cmd: 'read', path: requirePath(a), sheet: a.sheet, variables: [a.variable], attach: a.attach })
    const d = read.data?.[0]
    if (!d) throw new Error(`could not read variable ${a.variable}`)
    const isText = d.type === 1
    const map = a.map ?? {}
    const hasDefault = a.default !== undefined && a.default !== null
    const hasMissing = a.missing !== undefined && a.missing !== null
    const keyOf = (v) => (isText ? String(v) : String(Number(v)))
    const out = d.values.map((v) => {
      if (v === null || v === undefined) return hasMissing ? a.missing : null
      const k = keyOf(v)
      if (Object.prototype.hasOwnProperty.call(map, k)) return map[k]
      if (hasDefault) return a.default
      return v
    })
    const r = await runWorker({ cmd: 'write', path: requirePath(a), sheet: a.sheet, columns: [{ index: d.index, values: out }], save: a.save, attach: a.attach })
    const changed = out.filter((v, i) => String(v) !== String(d.values[i])).length
    const lines = [`Recoded #${d.index} ${d.cleanName || d.name}: ${changed} value(s) changed`]
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },

  async set_measurement(a) {
    const r = await runWorker({ cmd: 'set_measurement', path: requirePath(a), sheet: a.sheet, variable: a.variable, type: a.type, save: a.save, attach: a.attach })
    const names = ['unspecified', 'auto', 'continuous', 'categorical', 'ordinal']
    const lines = [`#${r.index} ${r.name}: measurement = ${names[r.measurementType] ?? r.measurementType}`]
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },

  async value_labels(a) {
    const r = await runWorker({ cmd: 'labels', path: requirePath(a), sheet: a.sheet, variable: a.variable, labels: a.labels, clear: a.clear, save: a.save, attach: a.attach })
    const lines = [`#${r.index} ${r.name}: ${r.applied.length} label(s) applied, ${r.count ?? '?'} total`]
    for (const l of r.applied) lines.push(`  ${fmt(l.value)} = ${JSON.stringify(l.label)}`)
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },
}

export default { tools, handlers }
