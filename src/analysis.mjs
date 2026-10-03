import { readFileSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, extname } from 'node:path'
import { runWorker } from './worker.mjs'
import { resolveModule } from './modules.mjs'
import { requirePath } from './util.mjs'
import { formatAnalysis } from './format.mjs'
import { pngToPdf } from './pdf.mjs'

export async function analysis(a, module, steps) {
  // STATISTICA cannot write PDF itself, so a `.pdf` saveGraph target is exported
  // to a temporary PNG by the worker and converted here.
  const conversions = []
  const rewritten = steps.map((step) => {
    if (step && typeof step === 'object' && step.saveGraph && /\.pdf$/i.test(String(step.saveGraph))) {
      const pdfPath = String(step.saveGraph)
      const tmpPng = join(tmpdir(), `sta-graph-${process.pid}-${conversions.length}-${Date.now()}.png`)
      conversions.push({ pdfPath, tmpPng })
      return { ...step, saveGraph: tmpPng }
    }
    return step
  })

  const r = await runWorker({
    cmd: 'analysis',
    path: requirePath(a),
    sheet: a.sheet,
    module: resolveModule(module),
    steps: rewritten,
    save: a.save,
    attach: a.attach,
  })

  for (const { pdfPath, tmpPng } of conversions) {
    const ext = extname(tmpPng)
    const base = tmpPng.slice(0, -ext.length)
    const pdfBase = pdfPath.slice(0, -extname(pdfPath).length)
    const matches = Object.values(r.graphs ?? {}).filter((v) => typeof v.out === 'string' && v.out.startsWith(base) && /\.png$/i.test(v.out))
    if (matches.length === 0) {
      r.warnings = r.warnings ?? []
      r.warnings.push(`pdf export failed for ${pdfPath}: no graph image was produced`)
      continue
    }
    for (const v of matches) {
      const suffix = String(v.out).slice(base.length, -extname(String(v.out)).length)
      const target = matches.length > 1 ? `${pdfBase}${suffix}.pdf` : pdfPath
      try {
        writeFileSync(target, pngToPdf(readFileSync(v.out)))
        rmSync(v.out, { force: true })
        v.out = target
        v.bytes = readFileSync(target).length
      } catch (e) {
        r.warnings = r.warnings ?? []
        r.warnings.push(`pdf export failed for ${target}: ${e.message}`)
      }
    }
  }

  for (const { tmpPng } of conversions) {
    try {
      rmSync(tmpPng, { force: true })
    } catch {}
  }

  return formatAnalysis(r)
}
