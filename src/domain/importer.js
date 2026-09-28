import { COST_PRICE_SCHEMA, OZON_ACCRUALS_SCHEMA } from './importSchema.js'
import { EXPENSE_CATEGORIES, FINANCE_CONFIG } from './financeConfig.js'

export class ImportError extends Error {
  constructor(message, code = 'INVALID_FILE') { super(message); this.name = 'ImportError'; this.code = code }
}

const SALE = 'продажи', RETURN = 'возвраты'

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

export function readCostReference(workbook) {
  const table = locateTable(workbook, COST_PRICE_SCHEMA)
  const articleColumn = table.columns.get('article')
  const costColumn = table.columns.get('cost')
  const costs = new Map(), issues = [], conflicts = [], articles = new Set()
  let skippedEmptyRows = 0
  for (let index = table.headerIndex + 1; index < table.sheet.length; index++) {
    const row = table.sheet[index] || []
    const article = normalizeArticle(row[articleColumn])
    if (!article) { if (row.some(cell => cell !== '' && cell != null)) skippedEmptyRows++; continue }
    articles.add(article)
    const cell = row[costColumn]
    const parsed = strictNumber(cell)
    if (parsed == null) {
      issues.push({ article, address: cell?.address || `${columnName(costColumn)}${index + 1}`, error: cell?.xlsxError || (cleanText(cell) || 'пустая ячейка') })
      continue
    }
    if (costs.has(article) && costs.get(article) !== parsed) {
      conflicts.push({ article, values: [costs.get(article), parsed], address: `${columnName(costColumn)}${index + 1}` })
      continue
    }
    costs.set(article, parsed)
  }
  return {
    costs, issues, conflicts,
    summary: { sheetName: table.sheetName, headerRow: table.headerIndex + 1, articleCount: articles.size, skippedEmptyRows, invalidCostCount: issues.length, conflictCount: conflicts.length },
  }
}

export function buildArticleReport(accrualWorkbook, costWorkbook) {
  const accrualTable = locateTable(accrualWorkbook, OZON_ACCRUALS_SCHEMA)
  const accruals = rowsFromSheet(accrualWorkbook, OZON_ACCRUALS_SCHEMA)
  const costReference = readCostReference(costWorkbook)
  const costByArticle = costReference.costs
  const result = new Map(), serviceGroups = new Map(), categoryAmounts = new Map(), unknownCategories = new Map()
  const expenseCategories = Object.fromEntries(EXPENSE_CATEGORIES.map(([key]) => [key, 0]))
  const seen = new Set(), duplicateIds = []
  let financialTotal = 0
  for (const [rowIndex, row] of accruals.entries()) {
    const id = cleanText(value(row, 'ID начисления'))
    const identity = id || `row:${rowIndex}`
    if (seen.has(identity)) { duplicateIds.push(identity); continue }
    seen.add(identity)
    const article = normalizeArticle(value(row, 'Артикул'))
    const amount = number(value(row, 'Сумма итого, руб.'))
    const group = cleanText(value(row, 'Группа услуг'))
    const type = cleanText(value(row, 'Тип начисления'))
    financialTotal += amount
    serviceGroups.set(group, (serviceGroups.get(group) || 0) + amount)
    categoryAmounts.set(`${group} — ${type}`, (categoryAmounts.get(`${group} — ${type}`) || 0) + amount)
    const key = article || '0'
    const item = result.get(key) || emptyArticle(key)
    item.sku ||= cleanText(value(row, 'SKU')); item.name ||= cleanText(value(row, 'Название товара'))
    const normalizedGroup = normalizeText(group), normalizedType = normalizeText(type)
    const quantity = Math.abs(number(value(row, 'Количество')))
    if (normalizedGroup === SALE && normalizedType === 'выручка') {
      item.soldQuantity += quantity; item.salesGross += Math.abs(amount); item.statementAmount += amount
    } else if (normalizedGroup === RETURN && (normalizedType === 'выручка' || /return|возврат/.test(normalizedType))) {
      item.returnedQuantity += quantity; item.returnsGross += Math.abs(amount); item.statementAmount += amount
    } else {
      const kind = operationKind(normalizedGroup, normalizedType)
      if (kind === 'income') item.compensations += amount
      else {
        const category = expenseCategory(normalizedGroup, normalizedType)
        const expense = -amount
        item.expenseCategories[category] += expense; item.ozonExpenses += expense
        expenseCategories[category] += expense
        if (category === 'advertising' || category === 'marketing') item.advertising += expense
        if (category === 'unknown') unknownCategories.set(`${group} — ${type}`, (unknownCategories.get(`${group} — ${type}`) || 0) + amount)
      }
      item.statementAmount += amount
    }
    result.set(key, item)
  }
  const articles = [...result.values()].map(item => {
    const netQuantity = item.soldQuantity - item.returnedQuantity
    const netRevenueGross = item.salesGross - item.returnsGross
    const vat = netRevenueGross * FINANCE_CONFIG.vatRatePercent / (100 + FINANCE_CONFIG.vatRatePercent)
    const netRevenue = netRevenueGross - vat
    if (item.article === '0') return { ...item, name: 'Расходы без артикула', netQuantity, netRevenueGross, vat, netRevenue, cost: null, grossProfit: null, marginalProfit: item.compensations - item.ozonExpenses, marginalProfitability: null, missingCost: false, unassigned: true }
    const hasSales = item.soldQuantity > 0 || item.returnedQuantity > 0
    const hasCost = !hasSales || costByArticle.has(item.article)
    const cost = hasCost ? (costByArticle.get(item.article) || 0) * netQuantity : null
    const grossProfit = hasCost ? netRevenue - cost : null
    const marginalProfit = hasCost ? grossProfit - item.ozonExpenses + item.compensations : null
    return { ...item, netQuantity, netRevenueGross, vat, netRevenue, cost, grossProfit, marginalProfit, marginalProfitability: hasCost && netRevenue ? marginalProfit / netRevenue * 100 : null, missingCost: !hasCost, unassigned: false }
  }).sort((a, b) => a.article.localeCompare(b.article, 'ru', { numeric: true }))
  const sum = field => articles.reduce((total, item) => total + (item[field] || 0), 0)
  const salesGross = sum('salesGross'), returnsGross = sum('returnsGross'), netRevenueGross = salesGross - returnsGross
  const vat = netRevenueGross * FINANCE_CONFIG.vatRatePercent / (100 + FINANCE_CONFIG.vatRatePercent), netRevenue = netRevenueGross - vat
  const knownCost = sum('cost'), ozonExpenses = sum('ozonExpenses'), compensations = sum('compensations')
  const complete = articles.every(item => !item.missingCost), grossProfit = complete ? netRevenue - knownCost : null
  const marginalProfit = complete ? grossProfit - ozonExpenses + compensations : null
  const warnings = [...unknownCategories].map(([label]) => `Обнаружена новая категория Ozon: «${label}». Сумма сохранена и включена один раз.`)
  if (duplicateIds.length) warnings.push(`Исключены дублирующиеся строки: ${duplicateIds.join(', ')}.`)
  warnings.push(...costReference.conflicts.map(item => `Конфликт себестоимости для артикула «${item.article}»: ${item.values.join(' и ')} (${item.address}).`))
  const salesWithCost = articles.filter(item => !item.unassigned && !item.missingCost).reduce((total, item) => total + item.salesGross, 0)
  const unassigned = articles.find(item => item.unassigned)?.marginalProfit || 0
  const skuMargin = articles.filter(item => !item.unassigned).reduce((total, item) => total + (item.marginalProfit || 0), 0)
  const checks = {
    categoriesEqualStatement: close([...serviceGroups.values()].reduce((a, b) => a + b, 0), financialTotal),
    salesMinusReturns: close(salesGross - returnsGross, netRevenueGross),
    expensesEqualCategories: close(Object.values(expenseCategories).reduce((a, b) => a + b, 0), ozonExpenses),
    skuPlusUnassigned: !complete || close(skuMargin + unassigned, marginalProfit),
    noDuplicateRows: duplicateIds.length === 0,
    missingCostNotZero: articles.filter(item => item.missingCost).every(item => item.cost === null),
    unknownCategoriesKept: close(expenseCategories.unknown, -[...unknownCategories.values()].reduce((a, b) => a + b, 0)),
  }
  const newCategoryDetails = [...unknownCategories].map(([label, amount]) => ({ label, amount }))
  return { period: extractPeriod(accrualTable.sheet, accrualTable.headerIndex), vatRatePercent: FINANCE_CONFIG.vatRatePercent, serviceGroups: Object.fromEntries(serviceGroups), expenseCategories, unknownCategories: newCategoryDetails, newCategoryDetails, warnings, checks, costImport: { ...costReference.summary, issues: costReference.issues, conflicts: costReference.conflicts }, totals: { soldQuantity: sum('soldQuantity'), returnedQuantity: sum('returnedQuantity'), netQuantity: sum('netQuantity'), salesGross, returnsGross, netRevenueGross, vat, netRevenue, cost: complete ? knownCost : null, knownCost, grossProfit, grossMargin: grossProfit != null && netRevenue ? grossProfit / netRevenue * 100 : null, ozonExpenses, advertising: articles.reduce((total, item) => total + item.advertising, 0), compensations, marginalProfit, marginalMargin: marginalProfit != null && netRevenue ? marginalProfit / netRevenue * 100 : null, accruals: financialTotal, complete, costCoveragePercent: salesGross ? salesWithCost / salesGross * 100 : 100, unassignedImpact: unassigned }, articles }
}

function emptyArticle(article) { return { article, sku: '', name: '', soldQuantity: 0, returnedQuantity: 0, salesGross: 0, returnsGross: 0, ozonExpenses: 0, advertising: 0, compensations: 0, statementAmount: 0, expenseCategories: Object.fromEntries(EXPENSE_CATEGORIES.map(([key]) => [key, 0])) } }
function operationKind(group, type) { return group === SALE || /compens|income|компенсац|прочие начислен|доход|балл.*скидк/.test(`${group} ${type}`) ? 'income' : 'expense' }
function expenseCategory(group, type) {
  const text = `${group} ${type}`
  if (/комисс|вознагражден/.test(text)) return 'commission'
  if (/эквайр/.test(text)) return 'acquiring'
  if (/обратн.*логист/.test(text)) return 'reverseLogistics'
  if (/обработ.*возврат/.test(text)) return 'returnProcessing'
  if (/последн.*мил/.test(text)) return 'lastMile'
  if (/приемк|приёмк/.test(text)) return 'acceptance'
  if (/хранен|размещен/.test(text)) return 'storageFbo'
  if (/fbs/.test(text)) return 'fbs'
  if (/агент|партнер|партнёр/.test(text)) return 'partners'
  if (/продвиж|реклам/.test(text)) return 'advertising'
  if (/premium|отзыв|маркетинг/.test(text)) return 'marketing'
  if (/достав|логист/.test(text)) return 'delivery'
  if (/штраф|другие услуги|прочие услуги|fbo/.test(text)) return 'penalties'
  return 'unknown'
}
function close(a, b) { return Math.abs(a - b) < 0.005 }

function locateTable(workbook, schema) {
  if (!workbook || typeof workbook !== 'object') throw new ImportError('Книга Excel не содержит листов.', 'EMPTY_FILE')
  const entries = Object.entries(workbook)
  const preferred = entries.filter(([name]) => equalText(name, schema.sheet))
  const ordered = [...preferred, ...entries.filter(entry => !preferred.includes(entry))]
  let best = null
  for (const [sheetName, sheet] of ordered) for (let index = 0; index < Math.min(schema === COST_PRICE_SCHEMA ? 30 : 20, sheet.length); index++) {
    const headers = Array.from(sheet[index] || [], cleanText), normalized = new Map(headers.map((header, column) => [normalizeText(header), column]))
    const required = schema.fields || [schema.matchKey, schema.costField]
    const columns = schema === COST_PRICE_SCHEMA ? costColumns(headers) : null
    const matched = schema === COST_PRICE_SCHEMA ? columns.size : required.filter(header => normalized.has(normalizeText(header))).length
    if (!best || matched > best.matched) best = { sheetName, index, headers, matched, required }
    if (matched === required.length) return { sheet, sheetName, headerIndex: index, headers, columns }
  }
  if (!entries.length) throw new ImportError('В книге нет листов.', 'EMPTY_FILE')
  const missing = (best?.required || schema.fields || []).filter(header => !best?.headers.some(found => normalizeText(found) === normalizeText(header)))
  const sheets = entries.map(([name]) => `«${name}»`).join(', ')
  const similar = best ? `Лучшая похожая строка: лист «${best.sheetName}», строка ${best.index + 1} (${best.headers.filter(Boolean).join(' | ') || 'пусто'}).` : ''
  throw new ImportError(`Не найдена таблица с нужными заголовками. Найденные листы: ${sheets}. ${similar} Отсутствуют обязательные колонки: ${missing.join(', ')}.`, 'MISSING_COLUMNS')
}

function costColumns(headers) {
  const columns = new Map()
  headers.forEach((header, index) => {
    const normalized = normalizeText(header).replace(/\s*,\s*/g, ', ')
    if (normalized === 'артикул') columns.set('article', index)
    if (['себес (руб)', 'себес', 'себестоимость', 'себестоимость, руб'].includes(normalized)) columns.set('cost', index)
  })
  return columns
}

function value(row, expected) { const key = Object.keys(row).find(header => normalizeText(header) === normalizeText(expected)); return key == null ? '' : row[key] }
function extractPeriod(sheet, headerIndex) { return sheet.slice(0, headerIndex).flat().map(cleanText).filter(Boolean).join(' · ') || 'Период не указан' }
function normalizeArticle(value) { const text = cleanText(value); return text === '0' ? '' : text }
function strictNumber(value) { if (typeof value === 'number') return Number.isFinite(value) ? value : null; if (value && typeof value === 'object') return null; const text = String(value ?? '').replace(/\s/g, '').replace(',', '.'); if (!text) return null; const parsed = Number(text); return Number.isFinite(parsed) ? parsed : null }
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
      else if (type === 'e') row[column] = { xlsxError: cleanText(raw) || 'ошибка Excel', address: cell.attrs.r || `${columnName(column)}${rowIndex + 1}` }
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
function columnName(index) { let result = ''; for (let value = index + 1; value; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result; return result }
function decodeXml(value) { return String(value).replace(/&(?:#(\d+)|#x([\da-f]+)|amp|lt|gt|quot|apos);/gi, (entity, decimal, hex) => decimal ? String.fromCodePoint(Number(decimal)) : hex ? String.fromCodePoint(parseInt(hex, 16)) : ({ '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" })[entity.toLowerCase()]) }
function textFile(files, path) { const value = files.get(normalizePath(path)); if (!value) throw new Error(`отсутствует ${path}`); return new TextDecoder().decode(value) }
function normalizePath(path) { const parts = []; path.replace(/\\/g, '/').split('/').forEach(part => part === '..' ? parts.pop() : part !== '.' && part && parts.push(part)); return parts.join('/') }
