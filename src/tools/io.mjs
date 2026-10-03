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
        separator: { type: 'string', description: 'Field separator character (e.g. "," or ";"). Default comma.' },
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
  {
    name: 'statistica_screenshot',
    description:
      'Launch STATISTICA visibly, optionally build a graph/analysis result, then capture the main window to a PNG/JPG image for a report. Content is prepared while hidden, then the window is shown and captured.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'File to open. Omit only in attach mode.' },
        sheet: { type: ['string', 'integer'] },
        out: { type: 'string', description: 'Image path (.png or .jpg). Required.' },
        module: { type: ['string', 'integer'], description: 'Optional analysis/graph module to build before capturing.' },
        variables: { type: 'string', description: 'Variable list for the module, e.g. "2 | 11".' },
        properties: { type: 'object', additionalProperties: true, description: 'Extra dialog properties for the module.' },
        result: { type: 'string', description: 'Result property to read/build (default "Graphs").' },
        run: { type: 'boolean', description: 'Call Run on the module before reading the result.' },
        mode: {
          type: 'string',
          enum: ['screen', 'window', 'document'],
          description: 'screen = whole display (default, nothing clipped), window = STATISTICA frame, document = active data/graph window only.',
        },
        waitMs: { type: 'integer', minimum: 0, description: 'Delay after showing the window before capture. Default 1500 ms.' },
      },
      required: ['out'],
      additionalProperties: false,
    },
  },
]

const handlers = {
  async export_csv(a) {
    const r = await runWorker({ cmd: 'export_csv', path: requirePath(a), sheet: a.sheet, out: a.out, separator: a.separator })
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

  async statistica_screenshot(a) {
    const r = await runWorker({
      cmd: 'screenshot',
      path: a.path,
      sheet: a.sheet,
      out: a.out,
      module: a.module,
      variables: a.variables,
      properties: a.properties,
      result: a.result,
      run: a.run,
      mode: a.mode,
      waitMs: a.waitMs,
      attach: a.attach,
    })
    return `Screenshot saved to ${r.out} (${r.bytes} bytes, ${r.mode ?? 'screen'} mode)`
  },
}

export default { tools, handlers }
