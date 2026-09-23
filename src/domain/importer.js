import { COST_PRICE_SCHEMA, OZON_ACCRUALS_SCHEMA } from './importSchema.js'

export class ImportError extends Error {
  constructor(message, code = 'INVALID_FILE') {
    super(message)
    this.name = 'ImportError'
    this.code = code
  }
}

export async function readExcelFile(file) {
  if (!file || !/\.xlsx$/i.test(file.name || '')) {
    throw new ImportError('Выберите файл Excel в формате .xlsx.', 'INVALID_FILE')
  }
  if (!file.size) throw new ImportError('Файл пустой.', 'EMPTY_FILE')
  try {
    return await parseXlsx(await file.arrayBuffer())
  } catch (error) {
    if (error instanceof ImportError) throw error
    throw new ImportError('Не удалось прочитать файл. Убедитесь, что это корректный .xlsx файл.', 'INVALID_FILE')
  }
}

export function rowsFromSheet(workbook, schema) {
  const sheet = workbook[schema.sheet]
  if (!sheet) throw new ImportError(`В файле нет листа «${schema.sheet}».`, 'MISSING_SHEET')
  if (sheet.length < schema.headerRow) throw new ImportError(`Лист «${schema.sheet}» пустой.`, 'EMPTY_FILE')
  const headers = sheet[schema.headerRow - 1].map(value => String(value ?? '').trim())
  const required = schema.fields || [schema.matchKey, schema.costField]
  const missing = required.filter(column => !headers.includes(column))
  if (missing.length) throw new ImportError(`Не найдены обязательные колонки: ${missing.join(', ')}.`, 'MISSING_COLUMNS')
  const rows = sheet.slice(schema.headerRow).filter(row => row.some(value => value !== '' && value != null))
  if (!rows.length) throw new ImportError(`На листе «${schema.sheet}» нет данных.`, 'EMPTY_FILE')
  return rows.map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])))
}

export function buildArticleReport(accrualWorkbook, costWorkbook) {
  const accruals = rowsFromSheet(accrualWorkbook, OZON_ACCRUALS_SCHEMA)
  const costs = rowsFromSheet(costWorkbook, COST_PRICE_SCHEMA)
  const costByArticle = new Map()
  costs.forEach(row => {
    const article = normalizeArticle(row[COST_PRICE_SCHEMA.matchKey])
    if (article) costByArticle.set(article, number(row[COST_PRICE_SCHEMA.costField]))
  })
  const result = new Map()
  const serviceGroups = new Map()
  accruals.forEach(row => {
    const article = normalizeArticle(row['Артикул'])
    const amount = number(row['Сумма итого, руб.'])
    const group = String(row['Группа услуг'] ?? '').trim()
    serviceGroups.set(group, (serviceGroups.get(group) || 0) + amount)
    const key = article || '0'
    const item = result.get(key) || { article: key, sku: '', name: '', quantity: 0, revenue: 0, ozonReward: 0, otherOzonExpenses: 0, accruals: 0 }
    item.sku ||= String(row.SKU ?? '').trim()
    item.name ||= String(row['Название товара'] ?? '').trim()
    if (group === 'Продажи' && String(row['Тип начисления'] ?? '').trim() === 'Выручка') item.quantity += number(row['Количество'])
    if (group === 'Продажи') item.revenue += amount
    else if (group === 'Вознаграждение Ozon') item.ozonReward += amount
    else item.otherOzonExpenses += amount
    item.accruals += amount
    result.set(key, item)
  })
  const articles = [...result.values()].map(item => {
    if (item.article === '0') return { ...item, name: 'Общие расходы без артикула', cost: null, profit: null, profitability: null, missingCost: false, unassigned: true }
    const hasCost = costByArticle.has(item.article)
    const cost = hasCost ? costByArticle.get(item.article) * item.quantity : null
    const profit = hasCost ? item.accruals - cost : null
    return { ...item, cost, profit, profitability: hasCost && item.revenue ? profit / item.revenue * 100 : null, missingCost: !hasCost, unassigned: false }
  }).sort((a, b) => a.article.localeCompare(b.article, 'ru', { numeric: true }))
  const revenue = serviceGroups.get('Продажи') || 0
  const accrualsTotal = [...serviceGroups.values()].reduce((sum, value) => sum + value, 0)
  const cost = articles.reduce((sum, item) => sum + (item.cost || 0), 0)
  const profit = accrualsTotal - cost
  return {
    period: extractPeriod(accrualWorkbook[OZON_ACCRUALS_SCHEMA.sheet][0]),
    serviceGroups: Object.fromEntries(serviceGroups),
    totals: { quantity: articles.reduce((sum, item) => sum + item.quantity, 0), revenue, accruals: accrualsTotal, cost, profit, profitability: revenue ? profit / revenue * 100 : 0, revenueWithoutVat: revenue / 1.22, profitWithoutVat: profit / 1.22 },
    articles,
  }
}

function extractPeriod(firstRow) {
  return (firstRow || []).map(value => String(value ?? '').trim()).filter(Boolean).join(' · ') || 'Период не указан'
}

function normalizeArticle(value) {
  return String(value ?? '').trim()
}

function number(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0
  const parsed = Number(String(value ?? '').replace(/\s/g, '').replace(',', '.'))
  return Number.isFinite(parsed) ? parsed : 0
}

async function parseXlsx(buffer) {
  if (buffer.byteLength < 4) throw new ImportError('Файл пустой.', 'EMPTY_FILE')
  const files = await unzip(buffer)
  const workbookXml = textFile(files, 'xl/workbook.xml')
  const relationshipsXml = textFile(files, 'xl/_rels/workbook.xml.rels')
  const relationships = new Map([...relationshipsXml.matchAll(/<Relationship\b([^>]+?)\/?\s*>/g)].map(match => {
    const attrs = attributes(match[1]); return [attrs.Id, attrs.Target]
  }))
  const shared = files.has('xl/sharedStrings.xml') ? parseSharedStrings(textFile(files, 'xl/sharedStrings.xml')) : []
  const workbook = {}
  for (const match of workbookXml.matchAll(/<sheet\b([^>]+?)\/?\s*>/g)) {
    const attrs = attributes(match[1])
    const target = relationships.get(attrs['r:id'])
    if (!target) continue
    const path = normalizePath(`xl/${target}`)
    workbook[decodeXml(attrs.name)] = parseSheet(textFile(files, path), shared)
  }
  if (!Object.keys(workbook).length) throw new ImportError('В книге нет листов.', 'EMPTY_FILE')
  return workbook
}

async function unzip(buffer) {
  const view = new DataView(buffer)
  const bytes = new Uint8Array(buffer)
  const files = new Map()
  let end = -1
  for (let index = Math.max(0, bytes.length - 65557); index <= bytes.length - 22; index++) {
    if (view.getUint32(index, true) === 0x06054b50) end = index
  }
  if (end < 0) throw new Error('ZIP directory not found')
  const entries = view.getUint16(end + 10, true)
  let offset = view.getUint32(end + 16, true)
  for (let i = 0; i < entries; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('Invalid ZIP entry')
    const method = view.getUint16(offset + 10, true)
    const compressedSize = view.getUint32(offset + 20, true)
    const nameLength = view.getUint16(offset + 28, true)
    const extraLength = view.getUint16(offset + 30, true)
    const commentLength = view.getUint16(offset + 32, true)
    const localOffset = view.getUint32(offset + 42, true)
    const name = new TextDecoder().decode(bytes.slice(offset + 46, offset + 46 + nameLength))
    const localNameLength = view.getUint16(localOffset + 26, true)
    const localExtraLength = view.getUint16(localOffset + 28, true)
    const start = localOffset + 30 + localNameLength + localExtraLength
    const compressed = bytes.slice(start, start + compressedSize)
    let data
    if (method === 0) data = compressed
    else if (method === 8) data = new Uint8Array(await new Response(new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())
    else throw new Error('Unsupported ZIP compression')
    files.set(normalizePath(name), data)
    offset += 46 + nameLength + extraLength + commentLength
  }
  return files
}

function parseSharedStrings(xml) {
  return [...xml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)].map(match =>
    [...match[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(part => decodeXml(part[1])).join(''))
}

function parseSheet(xml, shared) {
  const rows = []
  for (const rowMatch of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowAttrs = attributes(rowMatch[1])
    const rowIndex = Number(rowAttrs.r || rows.length + 1) - 1
    const row = []
    for (const cellMatch of rowMatch[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const cellAttrs = attributes(cellMatch[1])
      const column = columnIndex(cellAttrs.r || '')
      const inline = cellMatch[2].match(/<is(?:\s[^>]*)?>([\s\S]*?)<\/is>/)
      const raw = cellMatch[2].match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/)?.[1] ?? ''
      let value
      if (inline) value = [...inline[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(part => decodeXml(part[1])).join('')
      else if (cellAttrs.t === 's') value = shared[Number(raw)] ?? ''
      else if (cellAttrs.t === 'str') value = decodeXml(raw)
      else value = raw === '' ? '' : Number(raw)
      row[column] = value
    }
    rows[rowIndex] = row
  }
  return Array.from({ length: rows.length }, (_, index) => rows[index] || [])
}

function attributes(source) {
  return Object.fromEntries([...source.matchAll(/([\w:-]+)="([^"]*)"/g)].map(match => [match[1], decodeXml(match[2])]))
}

function columnIndex(reference) {
  return [...(reference.match(/^[A-Z]+/i)?.[0] || 'A').toUpperCase()].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1
}

function decodeXml(value) {
  return String(value).replace(/&(?:#(\d+)|#x([\da-f]+)|amp|lt|gt|quot|apos);/gi, (entity, decimal, hex) => {
    if (decimal) return String.fromCodePoint(Number(decimal))
    if (hex) return String.fromCodePoint(parseInt(hex, 16))
    return { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" }[entity.toLowerCase()]
  })
}

function textFile(files, path) {
  const value = files.get(normalizePath(path))
  if (!value) throw new Error(`Missing ${path}`)
  return new TextDecoder().decode(value)
}

function normalizePath(path) {
  const parts = []
  path.replace(/\\/g, '/').split('/').forEach(part => part === '..' ? parts.pop() : part !== '.' && part && parts.push(part))
  return parts.join('/')
}
