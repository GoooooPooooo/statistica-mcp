import { runWorker } from '../worker.mjs'
import { analysis } from '../analysis.mjs'
import { requirePath } from '../util.mjs'
import { fmt, numify, varSpec } from '../format.mjs'
import { TS_PROCEDURES } from '../modules.mjs'

const tools = [
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
    name: 'statistica_anova',
    description:
      'Analysis of variance via STATISTICA General Linear Models (module 4100 / ANOVA). `dependent` is the outcome; `between` lists factor/covariate effects. Returns the ANOVA table (UnivariateResults) and parameter estimates.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        dependent: { type: ['string', 'integer'], description: 'Dependent (outcome) variable.' },
        between: {
          type: 'array',
          items: { type: ['string', 'integer'] },
          description: 'Factors / effects entered in the model.',
        },
        method: {
          type: 'string',
          enum: ['standard', 'all_effects', 'forward', 'backward'],
          description: 'Model-building method. Default standard.',
        },
      },
      required: ['path', 'dependent', 'between'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_cluster',
    description:
      'Hierarchical cluster analysis (module 2201). `variables` are the variables to cluster; returns cluster membership, the amalgamation schedule and descriptive statistics.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] }, description: 'Variables to cluster.' },
      },
      required: ['path', 'variables'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_factor',
    description:
      'Factor analysis / principal components (module 2101). Extraction method defaults to PrincipalComponents; `factors` sets the requested number of factors. Returns eigenvalues, loadings and communalities.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        variables: { type: 'array', items: { type: ['string', 'integer'] } },
        method: {
          type: 'string',
          enum: ['principal_components', 'principal_axis', 'centroid', 'maximum_likelihood', 'minres'],
          description: 'Extraction method. Default principal_components.',
        },
        factors: { type: 'integer', minimum: 1, description: 'Number of factors to extract. Omit for the engine default.' },
      },
      required: ['path', 'variables'],
      additionalProperties: false,
    },
  },
  {
    name: 'statistica_correlation_matrix',
    description:
      'Build lagged series products for a time series: for lags 1..lags it creates variables Lag1..LagK holding x(t)*x(t-lag) (correlation products), optionally smoothed with a moving average, and returns their preview. Use mode "shift" for plain lagged series.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string' },
        sheet: { type: ['string', 'integer'] },
        variable: { type: ['string', 'integer'], description: 'Source series.' },
        lags: { type: 'integer', minimum: 1, description: 'Number of lag columns to build. Default 12.' },
        mode: { type: 'string', enum: ['product', 'shift'], description: 'product = x(t)*x(t-lag) (default), shift = x(t-lag).' },
        smooth: { type: 'integer', minimum: 2, description: 'Moving-average window applied to each new column.' },
        prefix: { type: 'string', description: 'Name prefix for the new variables. Default "Lag".' },
      },
      required: ['path', 'variable'],
      additionalProperties: false,
    },
  },
]

const handlers = {
  async descriptives(a) {
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
  },

  async statistica_descriptives(a) {
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
  },

  async statistica_correlation(a) {
    const opts = { DisplayCorrelationMatrix: true, MeansAndStandardDeviations: true, DisplayPAndN: true }
    const vs = varSpec(a.variables)
    if (vs) opts.VariableList = vs
    const steps = [{ set: { Statistics: 1 } }, { run: true }, { set: opts }, { run: true }, { result: 'CorrelationMatrix' }]
    return analysis(a, 1301, steps)
  },

  async statistica_frequencies(a) {
    const opts = {}
    const vs = varSpec(a.variables)
    if (vs) opts.Variables = vs
    const steps = [{ set: { Statistics: 9 } }, { run: true }, { set: opts }, { run: true }, { result: 'Summary' }, { result: 'Histograms' }]
    return analysis(a, 1301, steps)
  },

  async statistica_regression(a) {
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
  },

  async statistica_t_test(a) {
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
  },

  async statistica_graph(a) {
    const steps = []
    if (a.properties && typeof a.properties === 'object') steps.push({ set: a.properties })
    steps.push({ set: { Variables: String(a.variables) } })
    if (a.out) steps.push({ saveGraph: a.out, result: 'Graphs' })
    else steps.push({ result: 'Graphs' })
    return analysis(a, a.module, steps)
  },

  async statistica_time_series(a) {
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
  },

  async statistica_anova(a) {
    const dep = varSpec([a.dependent])
    const between = varSpec(a.between)
    if (!dep || !between) throw new Error('`dependent` and `between` are required')
    const methodMap = {
      all_effects: 'RegressionAllEffectsIncluded',
      forward: 'ForwardStepwiseRegression',
      backward: 'BackwardStepwiseRegression',
    }
    const opts = { Variables: `${dep} | ${between}` }
    const methodProp = methodMap[a.method ?? 'standard']
    if (methodProp) opts[methodProp] = true
    const steps = [
      { set: { GLMAnalysisItem: 1 } },
      { run: true },
      { set: opts },
      { run: true },
      { result: 'UnivariateResults' },
      { result: 'Coefficients' },
    ]
    return analysis(a, 4100, steps)
  },

  async statistica_cluster(a) {
    const vs = varSpec(a.variables)
    if (!vs) throw new Error('`variables` is required')
    const steps = [
      { run: true },
      { set: { Variables: vs } },
      { run: true },
      { result: 'AmalgamationSchedule' },
      { result: 'DistanceMatrix' },
      { result: 'DescriptiveStatistics' },
    ]
    return analysis(a, 2201, steps)
  },

  async statistica_factor(a) {
    const vs = varSpec(a.variables)
    if (!vs) throw new Error('`variables` is required')
    const methodMap = {
      principal_components: 'PrincipalComponents',
      principal_axis: 'PrincipalAxisMethod',
      centroid: 'CentroidMethod',
      maximum_likelihood: 'MaximumLikelihoodFactors',
      minres: 'IteratedCommunalitiesMINRES',
    }
    const methodProp = methodMap[a.method ?? 'principal_components']
    const extract = { [methodProp]: true }
    if (a.factors !== undefined) extract.NumberOfFactors = a.factors
    const steps = [
      { set: { Variables: vs } },
      { run: true },
      { set: extract },
      { run: true },
      { result: 'Eigenvalues' },
      { result: 'FactorLoadings' },
      { result: 'Communalities' },
    ]
    return analysis(a, 2101, steps)
  },

  async statistica_correlation_matrix(a) {
    const lags = a.lags ?? 12
    const prefix = a.prefix ?? 'Lag'
    const mode = a.mode ?? 'product'
    const read = await runWorker({ cmd: 'read', path: requirePath(a), sheet: a.sheet, variables: [a.variable], attach: a.attach })
    const d = read.data?.[0]
    if (!d) throw new Error(`could not read variable ${a.variable}`)
    if (d.type === 1) throw new Error('statistica_correlation_matrix needs a numeric series')
    const n = d.values.length
    const x = d.values.map((v) => (v === null || v === undefined ? null : Number(v)))
    const cols = []
    for (let L = 1; L <= lags; L++) {
      const vals = new Array(n).fill(null)
      for (let i = L; i < n; i++) {
        if (x[i] === null || x[i - L] === null) continue
        vals[i] = mode === 'shift' ? x[i - L] : x[i] * x[i - L]
      }
      if (a.smooth && a.smooth >= 2) {
        const w = a.smooth
        const sm = new Array(n).fill(null)
        for (let i = 0; i < n; i++) {
          let sum = 0
          let cnt = 0
          for (let k = 0; k < w; k++) {
            const j = i - k
            if (j < 0) continue
            if (vals[j] === null) { cnt = 0; sum = 0; break }
            sum += vals[j]
            cnt++
          }
          if (cnt === w) sm[i] = sum / w
        }
        cols.push({ name: `${prefix}${L}`, type: 0, values: sm })
      } else {
        cols.push({ name: `${prefix}${L}`, type: 0, values: vals })
      }
    }
    const r = await runWorker({ cmd: 'addwrite', path: requirePath(a), sheet: a.sheet, columns: cols, save: a.save, attach: a.attach })
    const lines = [`Added ${r.added.length} lag column(s) (${mode}${a.smooth ? `, SMA(${a.smooth})` : ''}); spreadsheet now ${r.variables} variables`]
    for (const c of r.added) lines.push(`  #${c.index} ${c.name}: ${c.written} value(s)`)
    const preview = cols.slice(0, 3).map((c) => {
      const vals = c.values.filter((v) => v !== null).slice(0, 6)
      return `  ${c.name}: ${vals.map((v) => fmt(v)).join(', ')}`
    })
    if (preview.length) lines.push('preview:', ...preview)
    if (r.saved) lines.push(`Saved to ${r.saved}`)
    return lines.join('\n')
  },
}

export default { tools, handlers }
