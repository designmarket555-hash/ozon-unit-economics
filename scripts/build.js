import { cp, mkdir, rm } from 'node:fs/promises'

await rm('dist', { recursive: true, force: true })
await mkdir('dist/src/domain', { recursive: true })
await cp('index.html', 'dist/index.html')
await cp('src/app.js', 'dist/src/app.js')
await cp('src/styles.css', 'dist/src/styles.css')
await cp('src/domain/calculator.js', 'dist/src/domain/calculator.js')
await cp('src/domain/importSchema.js', 'dist/src/domain/importSchema.js')
console.log('Production files written to dist/')
