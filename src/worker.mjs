import { spawn } from 'node:child_process'
import { mkdtempSync, writeFileSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { log } from './log.mjs'
import { TIMEOUT_MS } from './constants.mjs'

const HERE = dirname(fileURLToPath(import.meta.url))
export const WORKER = resolve(HERE, '..', 'sta.ps1')

export function runWorker(payload) {
  return new Promise((resolvePromise, reject) => {
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
      if (parsed.ok) resolvePromise(parsed.result)
      else reject(new Error(`${parsed.error}${parsed.hresult ? ` (${parsed.hresult})` : ''}`))
    })
    function cleanup() {
      try {
        rmSync(dir, { recursive: true, force: true })
      } catch {}
    }
  })
}
