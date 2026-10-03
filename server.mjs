#!/usr/bin/env node
import { startServer } from './src/protocol.mjs'
import { TOOLS, callTool } from './src/tools/index.mjs'
import { WORKER } from './src/worker.mjs'
import { log } from './src/log.mjs'

startServer({ tools: TOOLS, callTool })
log(`ready (worker: ${WORKER})`)
