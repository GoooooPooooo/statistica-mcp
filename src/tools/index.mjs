import inspect from './inspect.mjs'
import edit from './edit.mjs'
import io from './io.mjs'
import stats from './stats.mjs'
import engine from './engine.mjs'

const GROUPS = [inspect, edit, io, stats, engine]

export const TOOLS = GROUPS.flatMap((g) => g.tools)

const HANDLERS = Object.assign({}, ...GROUPS.map((g) => g.handlers))

// Add the optional `attach` switch to every data/analysis tool (not to info/list/import/export).
for (const t of TOOLS) {
  if (['statistica_info', 'list_analysis_modules', 'import_data', 'export_csv', 'statistica_open'].includes(t.name)) continue
  if (t.inputSchema && t.inputSchema.properties) {
    t.inputSchema.properties.attach = {
      type: 'boolean',
      description: 'Attach to the already-running STATISTICA instance and edit it live (no new process, the app is not closed).',
    }
  }
}

export async function callTool(name, a) {
  const handler = HANDLERS[name]
  if (!handler) throw new Error(`unknown tool: ${name}`)
  return handler(a)
}
