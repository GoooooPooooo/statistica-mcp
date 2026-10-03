import { SERVER_NAME } from './constants.mjs'

export function log(msg) {
  process.stderr.write(`[${SERVER_NAME}] ${msg}\n`)
}
