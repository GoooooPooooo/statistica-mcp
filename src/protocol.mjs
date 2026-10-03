import { log } from './log.mjs'
import { SERVER_NAME, SERVER_VERSION, DEFAULT_PROTOCOL } from './constants.mjs'

export function send(msg) {
  process.stdout.write(JSON.stringify(msg) + '\n')
}

export function reply(id, result) {
  send({ jsonrpc: '2.0', id, result })
}

export function fail(id, code, message, data) {
  const err = { code, message }
  if (data !== undefined) err.data = data
  send({ jsonrpc: '2.0', id, error: err })
}

function pickProtocol(clientVersion) {
  return typeof clientVersion === 'string' && clientVersion.length > 0 ? clientVersion : DEFAULT_PROTOCOL
}

export function startServer({ tools, callTool }) {
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
          reply(id, { tools })
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
}
