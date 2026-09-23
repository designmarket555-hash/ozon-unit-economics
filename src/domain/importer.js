import { COST_PRICE_SCHEMA, OZON_ACCRUALS_SCHEMA } from './importSchema.js'

export class ImportError extends Error {
  constructor(message, code = 'INVALID_FILE') { super(message); this.name = 'ImportError'; this.code = code }
}

const KNOWN_GROUPS = new Set(['Возвраты', 'Вознаграждение Ozon', 'Другие услуги и штрафы', 'Продажи', 'Продвижение и реклама', 'Прочие начисления', 'Услуги FBO', 'Услуги агентов', 'Услуги доставки'])
const KNOWN_TYPES = new Set(['Выручка', 'Баллы за скидки', 'Комиссия', 'Доставка'])

export async function readExcelFile(file) {
  if (!file) throw new ImportError('Выберите файл Excel в формате .xlsx.', 'INVALID_FILE')
  if (!/\.xlsx$/i.test(file.name || '')) throw new ImportError('Выберите файл Excel в формате .xlsx.', 'INVALID_FILE')
  if (file.size === 0) throw new ImportError('Файл пустой.', 'EMPTY_FILE')
  let buffer
  try { buffer = await file.arrayBuffer() } catch { throw new ImportError('Не удалось прочитать выбранный файл.', 'READ_ERROR') }
  try { return await parseXlsx(buffer) } catch (error) {
    if (error instanceof ImportError) throw error
    throw new ImportError(`Повреждён ZIP/XLSX: ${error.message}`, 'CORRUPT_XLSX')
  }
}

export function rowsFromSheet(workbook, schema) {
  const found = locateTable(workbook, schema)
  const { sheet, sheetName, headerIndex, headers } = found
  const rows = sheet.slice(headerIndex + 1).filter(row => row?.some(value => value !== '' && value != null))
  if (!rows.length) throw new ImportError(`На листе «${sheetName}» нет данных.`, 'EMPTY_FILE')
  return rows.map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])))
}

export function buildArticleReport(accrualWorkbook, costWorkbook) {
  const accrualTable = locateTable(accrualWorkbook, OZON_ACCRUALS_SCHEMA)
  const accruals = rowsFromSheet(accrualWorkbook, OZON_ACCRUALS_SCHEMA)
  const costs = rowsFromSheet(costWorkbook, COST_PRICE_SCHEMA)
  const costByArticle = new Map(costs.map(row => [normalizeArticle(value(row, COST_PRICE_SCHEMA.matchKey)), number(value(row, COST_PRICE_SCHEMA.costField))]).filter(([article]) => article))
  const result = new Map(), serviceGroups = new Map(), unknownGroups = new Set(), unknownTypes = new Set(), categoryAmounts = new Map()
  for (const row of accruals) {
    const article = normalizeArticle(value(row, 'Артикул'))
    const amount = number(value(row, 'Сумма итого, руб.'))
    const group = cleanText(value(row, 'Группа услуг'))
    const type = cleanText(value(row, 'Тип начисления'))
    serviceGroups.set(group, (serviceGroups.get(group) || 0) + amount)
    if (group && !KNOWN_GROUPS.has(group)) unknownGroups.add(group)
    if (type && !KNOWN_TYPES.has(type)) unknownTypes.add(type)
    categoryAmounts.set(`${group} — ${type}`, (categoryAmounts.get(`${group} — ${type}`) || 0) + amount)
    const key = article || '0'
    const item = result.get(key) || { article: key, sku: '', name: '', quantity: 0, revenue: 0, ozonReward: 0, otherOzonExpenses: 0, accruals: 0 }
    item.sku ||= cleanText(value(row, 'SKU')); item.name ||= cleanText(value(row, 'Название товара'))
    if (equalText(group, 'Продажи') && equalText(type, 'Выручка')) item.quantity += number(value(row, 'Количество'))
    if (equalText(group, 'Продажи')) item.revenue += amount
    else if (equalText(group, 'Вознаграждение Ozon')) item.ozonReward += amount
    else item.otherOzonExpenses += amount
    item.accruals += amount; result.set(key, item)
  }
  const articles = [...result.values()].map(item => {
    if (item.article === '0') return { ...item, name: 'Общие расходы без артикула', cost: null, profit: null, profitability: null, missingCost: false, unassigned: true }
    const hasCost = costByArticle.has(item.article), cost = hasCost ? costByArticle.get(item.article) * item.quantity : null
    const profit = hasCost ? item.accruals - cost : null
    return { ...item, cost, profit, profitability: hasCost && item.revenue ? profit / item.revenue * 100 : null, missingCost: !hasCost, unassigned: false }
  }).sort((a, b) => a.article.localeCompare(b.article, 'ru', { numeric: true }))
  const revenue = [...serviceGroups].filter(([group]) => equalText(group, 'Продажи')).reduce((sum, [, amount]) => sum + amount, 0)
  const accrualsTotal = [...serviceGroups.values()].reduce((sum, amount) => sum + amount, 0)
  const cost = articles.reduce((sum, item) => sum + (item.cost || 0), 0), profit = accrualsTotal - cost
  const warnings = [...unknownGroups].map(group => `Обнаружена новая группа услуг Ozon: «${group}». Сумма включена в общий итог.`)
  warnings.push(...[...unknownTypes].map(type => `Обнаружен новый тип начисления Ozon: «${type}». Сумма включена в общий итог.`))
  const newCategoryDetails = [...categoryAmounts].filter(([label]) => [...unknownGroups].some(group => label.startsWith(`${group} —`)) || [...unknownTypes].some(type => label.endsWith(`— ${type}`))).map(([label, amount]) => ({ label, amount }))
  return { period: extractPeriod(accrualTable.sheet, accrualTable.headerIndex), serviceGroups: Object.fromEntries(serviceGroups), unknownGroups: [...unknownGroups], unknownTypes: [...unknownTypes], newCategoryDetails, warnings, totals: { quantity: articles.reduce((sum, item) => sum + item.quantity, 0), revenue, accruals: accrualsTotal, cost, profit, profitability: revenue ? profit / revenue * 100 : 0, revenueWithoutVat: revenue / 1.22, profitWithoutVat: profit / 1.22 }, articles }
}

function locateTable(workbook, schema) {
  if (!workbook || typeof workbook !== 'object') throw new ImportError('Книга Excel не содержит листов.', 'EMPTY_FILE')
  const entries = Object.entries(workbook)
  const preferred = entries.filter(([name]) => equalText(name, schema.sheet))
  const ordered = [...preferred, ...entries.filter(entry => !preferred.includes(entry))]
  let best = null
  for (const [sheetName, sheet] of ordered) for (let index = 0; index < Math.min(20, sheet.length); index++) {
    const headers = (sheet[index] || []).map(cleanText), normalized = new Map(headers.map((header, column) => [normalizeText(header), column]))
    const required = schema.fields || [schema.matchKey, schema.costField]
    const matched = required.filter(header => normalized.has(normalizeText(header))).length
    if (!best || matched > best.matched) best = { sheetName, index, headers, matched, required }
    if (matched === required.length) return { sheet, sheetName, headerIndex: index, headers }
  }
  if (!entries.length) throw new ImportError('В книге нет листов.', 'EMPTY_FILE')
  const missing = (best?.required || schema.fields || []).filter(header => !best?.headers.some(found => normalizeText(found) === normalizeText(header)))
  const sheets = entries.map(([name]) => `«${name}»`).join(', ')
  const similar = best ? `Лучшая похожая строка: лист «${best.sheetName}», строка ${best.index + 1} (${best.headers.filter(Boolean).join(' | ') || 'пусто'}).` : ''
  throw new ImportError(`Не найдена таблица с нужными заголовками. Найденные листы: ${sheets}. ${similar} Отсутствуют обязательные колонки: ${missing.join(', ')}.`, 'MISSING_COLUMNS')
}

function value(row, expected) { const key = Object.keys(row).find(header => normalizeText(header) === normalizeText(expected)); return key == null ? '' : row[key] }
function extractPeriod(sheet, headerIndex) { return sheet.slice(0, headerIndex).flat().map(cleanText).filter(Boolean).join(' · ') || 'Период не указан' }
function normalizeArticle(value) { const text = cleanText(value); return text === '0' ? '' : text }
function number(value) { if (typeof value === 'number') return Number.isFinite(value) ? value : 0; const parsed = Number(String(value ?? '').replace(/\s/g, '').replace(',', '.')); return Number.isFinite(parsed) ? parsed : 0 }
function cleanText(value) { return decodeXml(String(value ?? '')).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim() }
function normalizeText(value) { return cleanText(value).toLocaleLowerCase('ru').replace(/ё/g, 'е').replace(/[.]+$/g, '').trim() }
function equalText(a, b) { return normalizeText(a) === normalizeText(b) }

async function parseXlsx(buffer) {
  if (buffer.byteLength < 4) throw new ImportError('Файл пустой.', 'EMPTY_FILE')
  const files = await unzip(buffer)
  let workbookRoot
  try { workbookRoot = parseXml(textFile(files, 'xl/workbook.xml')) } catch (error) { throw new ImportError(`Не удалось прочитать workbook.xml: ${error.message}`, 'INVALID_WORKBOOK') }
  let relRoot
  try { relRoot = parseXml(textFile(files, 'xl/_rels/workbook.xml.rels')) } catch (error) { throw new ImportError(`Не удалось прочитать связи workbook.xml: ${error.message}`, 'INVALID_WORKBOOK') }
  const relationships = new Map(descendants(relRoot, 'Relationship').map(node => [node.attrs.Id, node.attrs.Target]))
  const shared = files.has('xl/sharedStrings.xml') ? parseSharedStrings(parseXml(textFile(files, 'xl/sharedStrings.xml'))) : []
  const workbook = {}, unresolved = []
  for (const node of descendants(workbookRoot, 'sheet')) {
    const target = relationships.get(node.attrs['r:id'])
    if (!target) { unresolved.push(node.attrs.name || '?'); continue }
    const path = normalizePath(target.startsWith('/') ? target : `xl/${target}`)
    if (!files.has(path)) { unresolved.push(`${node.attrs.name || '?'} → ${path}`); continue }
    workbook[cleanText(node.attrs.name)] = parseSheet(parseXml(textFile(files, path)), shared)
  }
  if (unresolved.length) throw new ImportError(`Не удалось разрешить путь к листу: ${unresolved.join(', ')}.`, 'UNRESOLVED_SHEET')
  if (!Object.keys(workbook).length) throw new ImportError('В книге нет листов.', 'EMPTY_FILE')
  return workbook
}

function parseSharedStrings(root) { return descendants(root, 'si').map(node => textContent(node)) }
function parseSheet(root, shared) {
  const rows = []
  for (const rowNode of descendants(root, 'row')) {
    const rowIndex = positiveIndex(rowNode.attrs.r, rows.length + 1) - 1, row = [], cells = rowNode.children.filter(child => child.name === 'c')
    let nextColumn = 0
    for (const cell of cells) {
      const column = cell.attrs.r ? columnIndex(cell.attrs.r) : nextColumn
      nextColumn = column + 1
      const type = cell.attrs.t || 'n', valueNode = cell.children.find(child => child.name === 'v'), inlineNode = cell.children.find(child => child.name === 'is')
      const raw = valueNode ? textContent(valueNode) : ''
      if (type === 'inlineStr') row[column] = inlineNode ? textContent(inlineNode) : ''
      else if (type === 's') row[column] = shared[Number(raw)] ?? ''
      else if (type === 'str') row[column] = cleanText(raw)
      else if (type === 'n' || !cell.attrs.t) row[column] = raw === '' ? '' : Number(raw)
      else if (type === 'b') row[column] = raw === '1'
      else throw new ImportError(`Неподдерживаемый тип ячейки «${type}».`, 'UNSUPPORTED_CELL_TYPE')
    }
    rows[rowIndex] = row
  }
  return Array.from({ length: rows.length }, (_, index) => rows[index] || [])
}

// A small namespace-aware SAX-style XML reader. Unlike the old importer it tokenizes
// XML and resolves entities before workbook structure is inspected.
function parseXml(xml) {
  const root = { name: '#document', attrs: {}, children: [] }, stack = [root]
  let index = 0
  while (index < xml.length) {
    const open = xml.indexOf('<', index)
    if (open < 0) { appendText(stack.at(-1), xml.slice(index)); break }
    appendText(stack.at(-1), xml.slice(index, open))
    if (xml.startsWith('<!--', open)) { const end = xml.indexOf('-->', open + 4); if (end < 0) throw new Error('незакрытый комментарий XML'); index = end + 3; continue }
    if (xml.startsWith('<?', open)) { const end = xml.indexOf('?>', open + 2); if (end < 0) throw new Error('незакрытая XML-инструкция'); index = end + 2; continue }
    if (xml.startsWith('<![CDATA[', open)) { const end = xml.indexOf(']]>', open + 9); if (end < 0) throw new Error('незакрытый CDATA'); appendText(stack.at(-1), xml.slice(open + 9, end), false); index = end + 3; continue }
    const end = findTagEnd(xml, open + 1); if (end < 0) throw new Error('незакрытый XML-тег')
    const source = xml.slice(open + 1, end).trim()
    if (source.startsWith('!')) { index = end + 1; continue }
    if (source.startsWith('/')) { if (stack.length === 1) throw new Error('лишний закрывающий XML-тег'); stack.pop() }
    else {
      const selfClosing = source.endsWith('/'), body = selfClosing ? source.slice(0, -1).trim() : source
      const split = body.search(/\s/), qualifiedName = split < 0 ? body : body.slice(0, split)
      const node = { name: localName(qualifiedName), attrs: parseAttributes(split < 0 ? '' : body.slice(split)), children: [] }
      stack.at(-1).children.push(node); if (!selfClosing) stack.push(node)
    }
    index = end + 1
  }
  if (stack.length !== 1) throw new Error('незакрытый XML-тег')
  return root
}
function findTagEnd(xml, start) { let quote = ''; for (let i = start; i < xml.length; i++) { const char = xml[i]; if (quote) { if (char === quote) quote = '' } else if (char === '"' || char === "'") quote = char; else if (char === '>') return i } return -1 }
function parseAttributes(source) { const attrs = {}; let i = 0; while (i < source.length) { while (/\s/.test(source[i] || '')) i++; if (i >= source.length) break; const start = i; while (i < source.length && !/[\s=]/.test(source[i])) i++; const name = source.slice(start, i); while (/\s/.test(source[i] || '')) i++; if (source[i++] !== '=') throw new Error(`некорректный атрибут ${name}`); while (/\s/.test(source[i] || '')) i++; const quote = source[i++]; if (quote !== '"' && quote !== "'") throw new Error(`некорректный атрибут ${name}`); const end = source.indexOf(quote, i); if (end < 0) throw new Error(`незакрытый атрибут ${name}`); attrs[name] = decodeXml(source.slice(i, end)); i = end + 1 } return attrs }
function appendText(node, text, decode = true) { if (text) node.children.push(decode ? decodeXml(text) : text) }
function descendants(node, name, result = []) { for (const child of node.children || []) if (typeof child !== 'string') { if (child.name === name) result.push(child); descendants(child, name, result) } return result }
function textContent(node) { return cleanText((node.children || []).map(child => typeof child === 'string' ? child : textContent(child)).join('')) }
function localName(name) { return name.includes(':') ? name.slice(name.indexOf(':') + 1) : name }

async function unzip(buffer) {
  const view = new DataView(buffer), bytes = new Uint8Array(buffer), files = new Map(); let end = -1
  for (let index = Math.max(0, bytes.length - 65557); index <= bytes.length - 22; index++) if (view.getUint32(index, true) === 0x06054b50) end = index
  if (end < 0) throw new Error('не найдена центральная директория ZIP')
  const entries = view.getUint16(end + 10, true); let offset = view.getUint32(end + 16, true)
  for (let i = 0; i < entries; i++) {
    if (offset + 46 > bytes.length || view.getUint32(offset, true) !== 0x02014b50) throw new Error('повреждена запись ZIP')
    const method = view.getUint16(offset + 10, true), compressedSize = view.getUint32(offset + 20, true), nameLength = view.getUint16(offset + 28, true), extraLength = view.getUint16(offset + 30, true), commentLength = view.getUint16(offset + 32, true), localOffset = view.getUint32(offset + 42, true)
    const name = new TextDecoder().decode(bytes.slice(offset + 46, offset + 46 + nameLength)); const localNameLength = view.getUint16(localOffset + 26, true), localExtraLength = view.getUint16(localOffset + 28, true), start = localOffset + 30 + localNameLength + localExtraLength, compressed = bytes.slice(start, start + compressedSize)
    let data
    if (method === 0) data = compressed
    else if (method === 8) { try { data = new Uint8Array(await new Response(new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer()) } catch { throw new Error(`не удалось распаковать ${name}`) } }
    else throw new Error(`неподдерживаемый метод сжатия ZIP: ${method}`)
    files.set(normalizePath(name), data); offset += 46 + nameLength + extraLength + commentLength
  }
  return files
}
function positiveIndex(value, fallback) { const parsed = Number(value); return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback }
function columnIndex(reference) { const letters = reference.match(/^[A-Z]+/i)?.[0]; if (!letters) return 0; return [...letters.toUpperCase()].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1 }
function decodeXml(value) { return String(value).replace(/&(?:#(\d+)|#x([\da-f]+)|amp|lt|gt|quot|apos);/gi, (entity, decimal, hex) => decimal ? String.fromCodePoint(Number(decimal)) : hex ? String.fromCodePoint(parseInt(hex, 16)) : ({ '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" })[entity.toLowerCase()]) }
function textFile(files, path) { const value = files.get(normalizePath(path)); if (!value) throw new Error(`отсутствует ${path}`); return new TextDecoder().decode(value) }
function normalizePath(path) { const parts = []; path.replace(/\\/g, '/').split('/').forEach(part => part === '..' ? parts.pop() : part !== '.' && part && parts.push(part)); return parts.join('/') }
