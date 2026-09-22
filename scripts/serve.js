import { createReadStream, existsSync, statSync } from 'node:fs'
import { createServer } from 'node:http'
import { dirname, extname, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const root = process.argv.includes('--dist') ? resolve(projectRoot, 'dist') : projectRoot
const port = Number(process.env.PORT || 4173)
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8' }

createServer((request, response) => {
  const pathname = decodeURIComponent(new URL(request.url, `http://${request.headers.host}`).pathname)
  const candidate = resolve(root, `.${pathname === '/' ? '/index.html' : pathname}`)
  const relativePath = relative(root, candidate)
  if (relativePath.startsWith('..') || !existsSync(candidate) || statSync(candidate).isDirectory()) {
    response.writeHead(404).end('Not found')
    return
  }
  response.writeHead(200, { 'Content-Type': types[extname(candidate)] || 'application/octet-stream' })
  createReadStream(candidate).pipe(response)
}).listen(port, '0.0.0.0', () => console.log(`Ozon P&L: http://0.0.0.0:${port}`))
