#!/usr/bin/env node
import { spawn } from 'node:child_process'
import { mkdtempSync, writeFileSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = dirname(fileURLToPath(import.meta.url))
const WORKER = join(HERE, 'sta.ps1')
const SERVER_NAME = 'statistica-mcp'
const SERVER_VERSION = '2.0.0'
const DEFAULT_PROTOCOL = '2024-11-05'
const TIMEOUT_MS = 600000

function log(msg) {
  process.stderr.write(`[${SERVER_NAME}] ${msg}\n`)
}

let buf = Buffer.alloc(0)
process.stdin.on('data', (chunk) => {
  buf = Buffer.concat([buf, chunk])
  let nl
  while ((nl = buf.indexOf(0x0a)) !== -1) {
    const line = buf.subarray(0, nl).toString('utf8').trim()
    buf = buf.subarray(nl + 1)
    if (line) handleLine(line)
  }
})
process.stdin.on('end', () => process.exit(0))

function send(msg) {
  process.stdout.write(JSON.stringify(msg) + '\n')
}
function reply(id, result) {
  send({ jsonrpc: '2.0', id, result })
}
function fail(id, code, message, data) {
  const err = { code, message }
  if (data !== undefined) err.data = data
  send({ jsonrpc: '2.0', id, error: err })
}

function handleLine(line) {
  let msg
  try {
    msg = JSON.parse(line)
  } catch (e) {
    log(`bad JSON: ${e.message}`)
    return
  }
  const { id, method, params } = msg
  if (id === undefined || id === null) {
    if (method === 'notifications/initialized') log('client initialized')
    return
  }
  try {
    switch (method) {
      case 'initialize':
        reply(id, {
          protocolVersion: pickProtocol(params?.protocolVersion),
          capabilities: { tools: { listChanged: false } },
          serverInfo: { name: SERVER_NAME, version: SERVER_VERSION },
        })
        break
      case 'ping':
        reply(id, {})
        break
      case 'tools/list':
        reply(id, { tools: TOOLS })
        break
      case 'tools/call':
        callTool(params?.name ?? '', params?.arguments ?? {}).then(
          (text) => reply(id, text === null ? { content: [], isError: true } : { content: [{ type: 'text', text }] }),
          (e) => {
            log(`tool error: ${e.message}`)
            reply(id, { content: [{ type: 'text', text: `ERROR: ${e.message}` }], isError: true })
          },
        )
        break
      default:
        fail(id, -32601, `method not found: ${method}`)
    }
  } catch (e) {
    fail(id, -32603, e.message)
  }
}

function pickProtocol(clientVersion) {
  return typeof clientVersion === 'string' && clientVersion.length > 0 ? clientVersion : DEFAULT_PROTOCOL
}

function runWorker(payload) {
  return new Promise((resolve, reject) => {
    const dir = mkdtempSync(join(tmpdir(), 'sta-mcp-'))
    const req = join(dir, 'req.json')
    const res = join(dir, 'res.json')
    writeFileSync(req, JSON.stringify(payload), 'utf8')
    const ps = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', WORKER, '-RequestFile', req, '-ResponseFile', res],
      { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true },
    )
    let stderr = ''
    ps.stdout.on('data', (d) => log(`worker: ${d.toString().trim()}`))
    ps.stderr.on('data', (d) => {
      stderr += d.toString()
    })
    const timer = setTimeout(() => {
      try {
        ps.kill()
      } catch {}
      reject(new Error(`worker timed out after ${TIMEOUT_MS} ms`))
    }, TIMEOUT_MS)
    ps.on('error', (e) => {
      clearTimeout(timer)
      cleanup()
      reject(new Error(`failed to start worker: ${e.message}`))
    })
    ps.on('close', (code) => {
      clearTimeout(timer)
      if (code !== 0) {
        cleanup()
        reject(new Error(`worker exited with code ${code}. ${stderr.trim().split('\n').slice(-4).join(' | ')}`))
        return
      }
      let raw
      try {
        raw = readFileSync(res, 'utf8')
      } catch (e) {
        cleanup()
        reject(new Error(`worker produced no response: ${e.message}. ${stderr.trim().split('\n').slice(-3).join(' | ')}`))
        return
      }
      cleanup()
      let parsed
      try {
        parsed = JSON.parse(raw)
      } catch (e) {
        reject(new Error(`unreadable worker response: ${e.message}`))
        return
      }
      if (parsed.ok) resolve(parsed.result)
      else reject(new Error(`${parsed.error}${parsed.hresult ? ` (${parsed.hresult})` : ''}`))
    })
    function cleanup() {
      try {
        rmSync(dir, { recursive: true, force: true })
      } catch {}
    }
  })
}

const TYPE_NAME = { 0: 'numeric', 1: 'text', 2: 'integer', 3: 'byte' }
const MISSING = -999999998

// AnalysisIdentifier from the STATISTICA type library.
const ANALYSIS_MODULES = [
  [1301, 'Basic Statistics/Tables'],
  [1601, 'Nonparametrics'],
  [1645, 'Distributions (Fit)'],
  [1701, 'Multiple Regression'],
  [1748, 'Fixed Nonlinear Regression'],
  [1791, 'Nonlinear Estimation'],
  [1901, 'Time Series/Forecasting'],
  [2101, 'Factor Analysis'],
  [2144, 'Principal Components & Classification'],
  [2201, 'Cluster Analysis'],
  [2281, 'Canonical Analysis'],
  [2321, 'Multidimensional Scaling'],
  [2401, 'Reliability/Item Analysis'],
  [2530, 'Neural Networks'],
  [2601, 'ANOVA/MANOVA'],
  [2801, 'Discriminant Analysis'],
  [2901, 'Log-Linear Analysis'],
  [3001, 'Survival Analysis'],
  [3101, 'Process Analysis'],
  [3251, 'Quality Control Charts'],
  [3401, 'Design of Experiments'],
  [3711, 'MARSplines'],
  [3721, 'Rapid Deployment'],
  [3731, 'Machine Learning'],
  [3801, 'SEPATH (Structural Equations)'],
  [3907, 'Correspondence Analysis'],
  [3951, 'Variance Components'],
  [4100, 'General Linear Models (GLM)'],
  [4400, 'Association Rules'],
  [4501, 'Generalized Linear/Nonlinear (GLZ)'],
  [4601, 'General Regression Models (GRM)'],
  [4651, 'Partial Least Squares (PLS)'],
  [4751, 'General Discriminant Analysis (GDA)'],
  [4801, 'Goodness of Fit'],
  [4830, 'General Cluster Analysis'],
  [4851, 'General CHAID'],
  [4870, 'Feature Selection'],
  [4900, 'Power Analysis'],
  [4911, 'Boosting Trees'],
  [4950, 'Generalized Additive Models'],
  [6690, 'Random Forest'],
  [6784, 'Cox Proportional Hazards'],
]

const MODULE_BY_NAME = new Map()
for (const [id, name] of ANALYSIS_MODULES) MODULE_BY_NAME.set(name.toLowerCase(), id)

function resolveModule(m) {
  if (typeof m === 'number' && Number.isFinite(m)) return m
  if (typeof m === 'string') {
    const s = m.trim()
    if (/^\d+$/.test(s)) return Number(s)
    const key = s.toLowerCase()
    if (MODULE_BY_NAME.has(key)) return MODULE_BY_NAME.get(key)
    for (const [id, name] of ANALYSIS_MODULES) {
      if (name.toLowerCase().includes(key)) return id
    }
  }
  throw new Error(`unknown analysis module: ${m}. Use list_analysis_modules.`)
}

const ENUMS = {
  1301: {
    BasicStatisticsConstants: {
      scBasDescriptives: 0,
      scBasCorrelationMatrices: 1,
      scBasIndependetTTestByGroups: 3,
      scBasIndependetTTestByVariables: 4,
      scBasDependentTTest: 5,
      scBasSingleTTest: 6,
      scBasBreakdowns: 8,
      scBasFrequencies: 9,
      scBasTablesandBanners: 10,
      scBasMultipleResponseTables: 11,
      scBasProbabilityCalculator: 13,
      scBasOtherSignificanceTests: 14,
      scBasUniqueCombinationBreakdown: 16,
    },
  },
  1901: {
    // NB: the Time Series dialog rejects the flagged constants (0x40000000+n) with
    // an Access Violation and expects the plain index (constant - 0x40000000).
    'TypeOfTransformation (settable index)': {
      SimpleOneSeries: 0,
      Smoothing: 1,
      TwoSeries: 2,
      Shifting: 3,
      DifferencingIntegrate: 4,
      Fourier: 5,
      ReviewAndPlot: 6,
      Autocorrelation: 7,
      Descriptive: 8,
    },
    InterruptedTimeSeries: {
      scTimAbruptPermanent: 1073741824,
      scTimGradualPermanent: 1073741825,
      scTimAbruptTemporary: 1073741826,
    },
  },
  1701: {
    RegressionInput: { scRegRawData: 1073741824, scRegCorrelationMatrix: 1073741825 },
    RegressionModelBuilding: { scRegAlleffects: 1073741824, scRegStandard: 1073741825, scRegForwardStepwise: 1073741826, scRegBackwardStepwise: 1073741827 },
  },
  4601: {
    RegressionInput: { scRegRawData: 1073741824, scRegCorrelationMatrix: 1073741825 },
    RegressionModelBuilding: { scRegAlleffects: 1073741824, scRegStandard: 1073741825, scRegForwardStepwise: 1073741826, scRegBackwardStepwise: 1073741827 },
  },
}

const TS_PROCEDURES = [
  'descriptives',
  'autocorrelation',
  'partial_autocorrelation',
  'cross_correlation',
  'arima',
  'spectral',
  'smoothing',
  'exponential_smoothing',
  'differencing',
  'seasonal_decomposition',
]

const TOOLS = [
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
        after: { type: 'integer', minimum: 0, description: 'Insert after this 1-based variable index. 0 appends at the end.' },
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
  {
    name: 'descriptives',
    description:
      'Fast descriptive statistics computed in Node from data read through COM (N, missing, mean, sd, se, min, q1, median, q3, max, sum). Missing values are honoured using each variable\'s declared missing code.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] }, description: 'Subset to report on.' },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_descriptives',
    description:
      'Descriptive statistics computed by the STATISTICA engine itself (Basic Statistics module), including ones the JS shortcut does not provide.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] }, description: 'Subset to describe. Omit for all variables.' },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_correlation',
    description: 'Pearson correlation matrix computed by the STATISTICA Basic Statistics module.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] } },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_frequencies',
    description: 'Frequency tables and histograms computed by the STATISTICA Basic Statistics module.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] } },
      },
      required: ['path'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_regression',
    description:
      'Multiple linear regression via STATISTICA General Regression Models. The first predictor list entry is the dependent variable? No: `dependent` is the outcome and `predictors` are the regressors.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        dependent: { type: ['string', 'integer'], description: 'Dependent (outcome) variable.' },
        predictors: {
          type: 'array',
          items: { type: ['string', 'integer'] },
          description: 'Independent (predictor) variables.',
        },
        method: {
          type: 'string',
          enum: ['standard', 'forward', 'backward', 'all_effects'],
          description: 'Model-building method. Default standard.',
        },
      },
      required: ['path', 'dependent', 'predictors'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_t_test',
    description:
      'Student t-tests via the STATISTICA Basic Statistics module. kind=single tests means against a constant; kind=dependent runs paired comparisons over the listed variables (pairs).',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        kind: { type: 'string', enum: ['single', 'dependent'], description: 'Test kind. Default single.' },
        variables: {
          type: 'array',
          items: { type: ['string', 'integer'] },
          description: 'single: variables to test. dependent: variables forming pairs (2, 4, ... entries).',
        },
        constant: { type: 'number', description: 'single: reference constant to test the mean against. Default 0.' },
        summary: { type: 'boolean', description: 'dependent only: also return the per-variable summary. Default true.' },
      },
      required: ['path', 'kind', 'variables'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_graph',
    description:
      'Build a STATISTICA graph and optionally export it to an image file. `variables` uses the module syntax, typically "x | y". ' +
      '`properties` sets additional dialog options (e.g. GraphType, FitType, ShowRawDataPoints). ' +
      'Common modules: 11003 2D Scatterplots, 11012 2D Line Plots, 11002 2D Histograms, 11010 2D Box Plots, 11021 3D Sequential, 11032 3D Surface.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        module: { type: ['string', 'integer'], description: 'Graph module id (see list_analysis_modules / AnalysisIdentifier).' },
        variables: { type: 'string', description: 'Variable list for the graph, e.g. "2 | 11".' },
        properties: { type: 'object', description: 'Extra dialog properties to set before building the graph.', additionalProperties: true },
        out: { type: 'string', description: 'Optional image path (.png/.jpg/.emf); parent folders are created.' },
      },
      required: ['path', 'module', 'variables'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_time_series',
    description:
      'Time Series / Forecasting module. procedure is one of: ' +
      TS_PROCEDURES.join(', ') +
      '. autocorrelation/partial_autocorrelation/spectral/smoothing/differencing/descriptives operate on a single series.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        procedure: { type: 'string', enum: TS_PROCEDURES, description: 'Which time-series analysis to run.' },
        variables: {
          type: 'array',
          items: { type: ['string', 'integer'] },
          description: 'Series variable(s). For arima the first is modelled (or use `focus`).',
        },
        focus: { type: 'integer', minimum: 1, description: '1-based position within `variables` of the series to analyse. Default 1.' },
        lags: { type: 'integer', minimum: 1, description: 'autocorrelation: number of lags. Default 20.' },
        window: { type: 'integer', minimum: 2, description: 'smoothing: moving-average window size. Default 3.' },
        arOrder: { type: 'integer', minimum: 0, description: 'arima: autoregressive order p. Default 1.' },
        maOrder: { type: 'integer', minimum: 0, description: 'arima: moving-average order q. Default 0.' },
        difference: { type: 'boolean', description: 'arima: difference the series. Default false.' },
        differenceLag: { type: 'integer', minimum: 1, description: 'arima: differencing lag. Default 1.' },
        differencePasses: { type: 'integer', minimum: 1, description: 'arima: number of differencing passes. Default 1.' },
        seasonalLag: { type: 'integer', minimum: 1, description: 'arima: seasonal lag (0 disables seasonality).' },
        sarOrder: { type: 'integer', minimum: 0, description: 'arima: seasonal AR order.' },
        smaOrder: { type: 'integer', minimum: 0, description: 'arima: seasonal MA order.' },
        forecasts: { type: 'integer', minimum: 1, description: 'arima: cases to forecast. Default 12.' },
        confidenceLevel: { type: 'number', description: 'arima: forecast confidence level. Default 0.95.' },
      },
      required: ['path', 'procedure', 'variables'],
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
  {
    name: 'run_analysis',
    description:
      'Run ANY STATISTICA analysis by executing an ordered list of steps against its dialog. Each step is one of: ' +
      '{"set": {PropertyName: value}} (set dialog properties), ' +
      '{"call": "MethodName", "args": [...]} (invoke a dialog method, e.g. ARIMAAndAutocorrelationFunctions), ' +
      '{"run": true} (execute the analysis), ' +
      '{"result": "Summary"} (read a result document/table after the run), ' +
      '{"saveGraph": "C:\\\\out\\\\plot.png", "result": "Graphs"} (export a graph document to an image; .png/.jpg/.emf). ' +
      'Results are returned as tables, arrays or document handles. Use describe_analysis to discover property and method names.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the source file.' },
        sheet: { type: ['string', 'integer'] },
        module: { type: ['string', 'integer'], description: 'Module id or name.' },
        steps: {
          type: 'array',
          description: 'Ordered steps, e.g. [{"set":{"Variables":"2 1"}},{"run":true},{"result":"Summary"}].',
          items: { type: 'object', additionalProperties: true },
        },
        save: { type: 'string', description: 'Optional destination .sta path to save the (possibly modified) input spreadsheet.' },
      },
      required: ['path', 'module', 'steps'],
      additionalProperties: false,
    },
  },
]

function fmt(v) {
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

function requirePath(a) {
  if (typeof a.path !== 'string' || a.path.trim() === '') throw new Error('`path` is required')
  return a.path
}

function numify(v, isText) {
  if (v === null || v === undefined) return null
  if (isText) return String(v)
  const n = Number(v)
  return Number.isNaN(n) ? null : n
}

function varSpec(variables) {
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

function renderResult(res, indent = '') {
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

function formatAnalysis(r) {
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

async function analysis(a, module, steps) {
  const r = await runWorker({
    cmd: 'analysis',
    path: requirePath(a),
    sheet: a.sheet,
    module: resolveModule(module),
    steps,
    save: a.save,
    attach: a.attach,
  })
  return formatAnalysis(r)
}

// Add the optional `attach` switch to every data/analysis tool (not to info/list/import/export).
for (const t of TOOLS) {
  if (['statistica_info', 'list_analysis_modules', 'import_data', 'export_csv'].includes(t.name)) continue
  if (t.inputSchema && t.inputSchema.properties) {
    t.inputSchema.properties.attach = {
      type: 'boolean',
      description: 'Attach to the already-running STATISTICA instance and edit it live (no new process, the app is not closed).',
    }
  }
}

async function callTool(name, a) {
  switch (name) {
    case 'statistica_info': {
      const r = await runWorker({ cmd: 'info' })
      return [`STATISTICA COM is available.`, `version: ${r.version} (${r.versionEx ?? ''})`, `exe: ${r.exe}`, `pid: ${r.pid}`].join('\n')
    }

    case 'list_analysis_modules': {
      const lines = ['STATISTICA analysis modules (use the id or the name with run_analysis):', '']
      for (const [id, nm] of ANALYSIS_MODULES) lines.push(`${String(id).padEnd(6)}${nm}`)
      return lines.join('\n')
    }

    case 'describe_spreadsheet': {
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
    }

    case 'read_variables': {
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
    }

    case 'write_variables': {
      if (!Array.isArray(a.columns) || a.columns.length === 0) throw new Error('`columns` must be a non-empty array')
      const r = await runWorker({ cmd: 'write', path: requirePath(a), sheet: a.sheet, columns: a.columns, save: a.save, attach: a.attach })
      const lines = [`Wrote ${r.written.length} variable(s); spreadsheet now ${r.cases} cases`]
      for (const w of r.written) lines.push(`  #${w.index} ${w.name} (${w.type}): ${w.written} value(s)`)
      if (r.saved) lines.push(`Saved to ${r.saved}`)
      return lines.join('\n')
    }

    case 'set_formula': {
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
    }

    case 'add_variables': {
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
    }

    case 'rename_variables': {
      const r = await runWorker({ cmd: 'rename', path: requirePath(a), sheet: a.sheet, renames: a.renames, longNames: a.longNames, save: a.save, attach: a.attach })
      const lines = [`Renamed ${r.renamed.length} variable(s)`]
      for (const d of r.renamed) lines.push(`  #${d.index}: "${d.from}" -> "${d.to}"`)
      if (a.save) lines.push(`Saved to ${a.save}`)
      return lines.join('\n')
    }

    case 'delete_variables': {
      const r = await runWorker({ cmd: 'delete_variables', path: requirePath(a), sheet: a.sheet, from: a.from, to: a.to, save: a.save, attach: a.attach })
      const lines = [`Deleted variables ${r.deletedFrom}..${r.deletedTo}; now ${r.variables} variables`]
      if (a.save) lines.push(`Saved to ${a.save}`)
      return lines.join('\n')
    }

    case 'set_size': {
      const r = await runWorker({ cmd: 'set_size', path: requirePath(a), sheet: a.sheet, cases: a.cases, variables: a.variables, save: a.save, attach: a.attach })
      const lines = [`Resized to ${r.cases} cases x ${r.variables} variables`]
      if (a.save) lines.push(`Saved to ${a.save}`)
      return lines.join('\n')
    }

    case 'export_csv': {
      const r = await runWorker({ cmd: 'export_csv', path: requirePath(a), sheet: a.sheet, out: a.out })
      return `Exported to ${r.out} (${r.bytes} bytes)`
    }

    case 'save_spreadsheet': {
      const r = await runWorker({ cmd: 'save_as', path: requirePath(a), sheet: a.sheet, out: a.out, overwrite: a.overwrite, copy: a.copy })
      return `Saved ${r.copy ? 'a copy' : 'in place'} to ${r.out} (${r.bytes} bytes)`
    }

    case 'import_data': {
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
    }

    case 'descriptives': {
      const r = await runWorker({ cmd: 'read', path: requirePath(a), sheet: a.sheet, variables: a.variables, attach: a.attach })
      const lines = [`Descriptive statistics for ${a.path}`, '']
      lines.push('variable                       N   missing      mean         sd         se       min         q1     median         q3       max        sum')
      lines.push('---------------------------  ---  -------  ----------  ----------  ---------  ----------  ----------  ----------  ----------  ----------  ----------')
      for (const d of r.data) {
        const label = (d.cleanName || d.name).slice(0, 27).padEnd(27)
        if (d.type === 1) {
          const uniq = new Set(d.values.map((v) => String(v ?? '')))
          lines.push(`${label} text, ${uniq.size} distinct value(s)`)
          continue
        }
        const vals = d.values.map((v) => numify(v, false)).filter((v) => v !== null)
        const missing = d.count - vals.length
        if (vals.length === 0) {
          lines.push(`${label}    0  ${String(missing).padStart(7)}  (all missing)`)
          continue
        }
        const sorted = [...vals].sort((x, y) => x - y)
        const n = vals.length
        const sum = vals.reduce((s, v) => s + v, 0)
        const mean = sum / n
        const ss = vals.reduce((s, v) => s + (v - mean) * (v - mean), 0)
        const sd = n > 1 ? Math.sqrt(ss / (n - 1)) : 0
        const se = n > 1 ? sd / Math.sqrt(n) : 0
        const q = (p) => {
          const h = (n - 1) * p
          const lo = Math.floor(h)
          const hi = Math.ceil(h)
          return sorted[lo] + (sorted[hi] - sorted[lo]) * (h - lo)
        }
        lines.push(
          [
            label,
            String(n).padStart(3),
            String(missing).padStart(9),
            fmt(mean).padStart(10),
            fmt(sd).padStart(10),
            fmt(se).padStart(9),
            fmt(sorted[0]).padStart(10),
            fmt(q(0.25)).padStart(10),
            fmt(q(0.5)).padStart(10),
            fmt(q(0.75)).padStart(10),
            fmt(sorted[n - 1]).padStart(10),
            fmt(sum).padStart(11),
          ].join('  '),
        )
      }
      return lines.join('\n')
    }

    case 'statistica_descriptives': {
      const opts = {
        ValidN: true,
        Mean: true,
        StandardDeviation: true,
        StandardErrorOfMean: true,
        MinimumMaximum: true,
        Median: true,
        Sum: true,
        Variance: true,
        Skewness: true,
        Kurtosis: true,
        LowerUpperQuartiles: true,
      }
      const vs = varSpec(a.variables)
      if (vs) opts.Variables = vs
      const steps = [{ set: { Statistics: 0 } }, { run: true }, { set: opts }, { result: 'Summary' }]
      return analysis(a, 1301, steps)
    }

    case 'statistica_correlation': {
      const opts = { DisplayCorrelationMatrix: true, MeansAndStandardDeviations: true, DisplayPAndN: true }
      const vs = varSpec(a.variables)
      if (vs) opts.VariableList = vs
      const steps = [{ set: { Statistics: 1 } }, { run: true }, { set: opts }, { run: true }, { result: 'CorrelationMatrix' }]
      return analysis(a, 1301, steps)
    }

    case 'statistica_frequencies': {
      const opts = {}
      const vs = varSpec(a.variables)
      if (vs) opts.Variables = vs
      const steps = [{ set: { Statistics: 9 } }, { run: true }, { set: opts }, { run: true }, { result: 'Summary' }, { result: 'Histograms' }]
      return analysis(a, 1301, steps)
    }

    case 'statistica_regression': {
      const dep = varSpec([a.dependent])
      const pred = varSpec(a.predictors)
      if (!dep || !pred) throw new Error('`dependent` and `predictors` are required')
      const methodMap = {
        all_effects: 'RegressionAllEffectsIncluded',
        forward: 'ForwardStepwiseRegression',
        backward: 'BackwardStepwiseRegression',
      }
      const opts = { Variables: `${dep} | ${pred}` }
      const methodProp = methodMap[a.method ?? 'standard']
      if (methodProp) opts[methodProp] = true
      const steps = [
        { set: { GRMAnalysisItem: 1 } },
        { run: true },
        { set: opts },
        { run: true },
        { result: 'Coefficients' },
        { result: 'UnivariateResults' },
      ]
      return analysis(a, 4601, steps)
    }

    case 'statistica_t_test': {
      const vs = varSpec(a.variables)
      if (!vs) throw new Error('`variables` is required')
      const kind = a.kind ?? 'single'
      let steps
      if (kind === 'dependent') {
        if (Array.isArray(a.variables) && a.variables.length % 2 !== 0) {
          throw new Error('dependent t-test needs an even number of variables (pairs)')
        }
        steps = [{ set: { Statistics: 5 } }, { run: true }, { set: { Variables: vs } }, { run: true }, { result: 'Summary' }]
      } else {
        const opts = { Variables: vs, TestMeansAgainstConstant: true }
        if (a.constant !== undefined) opts.Constant = a.constant
        steps = [{ set: { Statistics: 6 } }, { run: true }, { set: opts }, { run: true }, { result: 'TTests' }]
      }
      return analysis(a, 1301, steps)
    }

    case 'statistica_graph': {
      const steps = []
      if (a.properties && typeof a.properties === 'object') steps.push({ set: a.properties })
      steps.push({ set: { Variables: String(a.variables) } })
      if (a.out) steps.push({ saveGraph: a.out, result: 'Graphs' })
      else steps.push({ result: 'Graphs' })
      return analysis(a, a.module, steps)
    }

    case 'statistica_time_series': {
      const vs = varSpec(a.variables)
      if (!vs) throw new Error('`variables` is required')
      const focus = a.focus ?? 1
      const proc = a.procedure
      let steps
      switch (proc) {
        case 'descriptives':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { run: true },
            { set: { TypeOfTransformation: 8 } },
            { result: 'DescriptiveStatistics' },
          ]
          break
        case 'autocorrelation':
        case 'partial_autocorrelation':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { run: true },
            { set: { TypeOfTransformation: 7, NumberOfLags: a.lags ?? 20, WhiteNoiseStandardErrors: true, PLevelForHighlighting: 0.05 } },
            { result: 'Autocorrelations' },
            { result: 'PartialAutocorrelations' },
          ]
          break
        case 'cross_correlation':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { run: true },
            { set: { TypeOfTransformation: 7, NumberOfLags: a.lags ?? 20, WhiteNoiseStandardErrors: true } },
            { set: { ResultsSelection: 1 } },
            { result: 'CrossCorrelations' },
          ]
          break
        case 'smoothing':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { run: true },
            {
              set: {
                TypeOfTransformation: 1,
                NPointsMovingAverage: true,
                NPointsWindowForMovingAverage: a.window ?? 3,
              },
            },
            { run: true },
            { result: 'SaveVariables' },
            { result: 'DescriptiveStatistics' },
          ]
          break
        case 'spectral':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { call: 'SpectralFourierAnalysis' },
            { run: true },
            { result: 'Summary' },
          ]
          break
        case 'differencing':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { run: true },
            {
              set: {
                TypeOfTransformation: 4,
                DifferenceSeries: true,
                LagForSimpleDifference: a.differenceLag ?? 1,
              },
            },
            { run: true },
            { result: 'SaveVariables' },
          ]
          break
        case 'exponential_smoothing':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { call: 'ExponentialSmoothingAndForecasting' },
            { set: { NoTrendNoSeasonalCompononent: true, ForecastNCases: a.forecasts ?? 12 } },
            { run: true },
            { result: 'Summary' },
          ]
          break
        case 'seasonal_decomposition':
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { call: 'SeasonalDecompositionCensus1' },
            { run: true },
            { result: 'Summary' },
          ]
          break
        case 'arima':
        default: {
          const set2 = {
            NumberOfAutoregressiveParameters: a.arOrder ?? 1,
            NumberOfMovingAverageParameters: a.maOrder ?? 0,
            DifferenceSeries: a.difference ?? false,
            DifferenceSeriesLag1: a.differenceLag ?? 1,
            NumberOfPassesDifferencing1: a.differencePasses ?? 1,
          }
          if (a.seasonalLag) {
            set2.SeasonalLag = a.seasonalLag
            set2.NumberOfSeasonalAutoregressiveParameters = a.sarOrder ?? 0
            set2.NumberOfSeasonalMovingAverageParameters = a.smaOrder ?? 0
          }
          steps = [
            { set: { Variables: vs, FocusTimeSeriesVariable: focus } },
            { call: 'ARIMAAndAutocorrelationFunctions' },
            { set: set2 },
            { run: true },
            { set: { NumberOfCasesToForecast: a.forecasts ?? 12, ConfidenceLevelsForForecasts: a.confidenceLevel ?? 0.95 } },
            { result: 'Summary' },
            { result: 'ForecastCases' },
          ]
          break
        }
      }
      return analysis(a, 1901, steps)
    }

    case 'describe_analysis': {
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
    }

    case 'run_analysis': {
      if (!Array.isArray(a.steps) || a.steps.length === 0) throw new Error('`steps` must be a non-empty array')
      return analysis(a, a.module, a.steps)
    }

    default:
      throw new Error(`unknown tool: ${name}`)
  }
}

log(`ready (worker: ${WORKER})`)
