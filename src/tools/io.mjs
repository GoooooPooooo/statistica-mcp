import { runWorker } from '../worker.mjs'
import { requirePath } from '../util.mjs'

const tools = [
  {
    name: 'export_csv',
    description: 'Export a sheet to CSV using the STATISTICA built-in CSV writer.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        out: { type: 'string', description: 'Destination CSV path. Parent folders are created.' },
      },
      required: ['path', 'out'],
      additionalProperties: false,
    },
  },
  {
    name: 'save_spreadsheet',
    description: 'Save a sheet to a new file. Extension selects the format (.sta, .stw, .csv, .xlsx).',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        out: { type: 'string', description: 'Destination path.' },
        overwrite: { type: 'boolean', description: 'Allow replacing an existing file. Default false.' },
        copy: { type: 'boolean', description: 'true = leave the source untouched (default), false = save in place.' },
      },
      required: ['path', 'out'],
      additionalProperties: false,
    },
  },
  {
    name: 'import_data',
    description:
      'Import a delimited text file or an Excel workbook into STATISTICA and optionally save it as .sta.',
    inputSchema: {
      type: 'object',
      properties: {
        source: { type: 'string', description: 'Absolute path to the .csv/.txt/.xls/.xlsx file.' },
        format: { type: 'string', enum: ['text', 'xls'], description: 'File format. Default text.' },
        separator: { type: 'string', description: 'Field separator for text files (e.g. "," or ";"). Omit for auto-detection.' },
        sheetNumber: { type: 'integer', minimum: 1, description: 'Excel sheet number. Default 1.' },
        variableNamesFromFirstRow: { type: 'boolean', description: 'Text format: use the first row as variable names. Default true.' },
        caseNamesFromFirstColumn: { type: 'boolean', description: 'Use the first column as case names. Default false.' },
        save: { type: 'string', description: 'Optional destination .sta path.' },
      },
      required: ['source'],
      additionalProperties: false,
    },
  },
]

const handlers = {
  async export_csv(a) {
    const r = await runWorker({ cmd: 'export_csv', path: requirePath(a), sheet: a.sheet, out: a.out })
    return `Exported to ${r.out} (${r.bytes} bytes)`
  },

  async save_spreadsheet(a) {
    const r = await runWorker({ cmd: 'save_as', path: requirePath(a), sheet: a.sheet, out: a.out, overwrite: a.overwrite, copy: a.copy })
    return `Saved ${r.copy ? 'a copy' : 'in place'} to ${r.out} (${r.bytes} bytes)`
  },

  async import_data(a) {
    const r = await runWorker({
      cmd: 'import',
      source: a.source,
      format: a.format,
      separator: a.separator,
      sheetNumber: a.sheetNumber,
      variableNamesFromFirstRow: a.variableNamesFromFirstRow,
      caseNamesFromFirstColumn: a.caseNamesFromFirstColumn,
      save: a.save,
    })
    const lines = [`Imported ${a.source}`, `Result: ${r.cases} cases x ${r.variables} variables ("${r.name}")`]
    if (r.saved) lines.push(`Saved to ${r.saved} (${r.bytes} bytes)`)
    return lines.join('\n')
  },
}

export default { tools, handlers }
