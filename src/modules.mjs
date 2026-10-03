// AnalysisIdentifier from the STATISTICA type library.
export const ANALYSIS_MODULES = [
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

export function resolveModule(m) {
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

export const ENUMS = {
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
  4100: {
    'GLMAnalysisItem (first Run opens the specification dialog)': { GeneralLinearModels: 1 },
  },
  2101: {
    'Extraction method (set the matching property to true on the second dialog)': {
      PrincipalComponents: 1,
      PrincipalAxisMethod: 1,
      CentroidMethod: 1,
      MaximumLikelihoodFactors: 1,
      IteratedCommunalitiesMINRES: 1,
    },
  },
}

export const TS_PROCEDURES = [
  'descriptives',
  'autocorrelation',
  'partial_autocorrelation',
  'cross_correlation',
  'arima',
  'spectral',
  'smoothing',
  'shift',
  'exponential_smoothing',
  'differencing',
  'seasonal_decomposition',
]
