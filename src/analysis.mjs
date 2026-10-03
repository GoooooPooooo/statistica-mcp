import { runWorker } from './worker.mjs'
import { resolveModule } from './modules.mjs'
import { requirePath } from './util.mjs'
import { formatAnalysis } from './format.mjs'

export async function analysis(a, module, steps) {
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
