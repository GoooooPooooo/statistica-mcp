import { analysis } from '../analysis.mjs'

const tools = [
  {
    name: 'run_analysis',
    description:
      'Run ANY STATISTICA analysis by executing an ordered list of steps against its dialog. Each step is one of: ' +
      '{"set": {PropertyName: value}} (set dialog properties), ' +
      '{"call": "MethodName", "args": [...]} (invoke a dialog method, e.g. ARIMAAndAutocorrelationFunctions), ' +
      '{"run": true} (execute the analysis), ' +
      '{"result": "Summary"} (read a result document/table after the run), ' +
      '{"saveGraph": "C:\\\\out\\\\plot.png", "result": "Graphs"} (export a graph document to an image or a native STATISTICA graph; .png/.jpg/.emf/.stg). ' +
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

const handlers = {
  async run_analysis(a) {
    if (!Array.isArray(a.steps) || a.steps.length === 0) throw new Error('`steps` must be a non-empty array')
    return analysis(a, a.module, a.steps)
  },
}

export default { tools, handlers }
