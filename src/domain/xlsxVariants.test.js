import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { buildArticleReport, readCostReference, readExcelFile, rowsFromSheet } from './importer.js'
import { COST_PRICE_SCHEMA, OZON_ACCRUALS_SCHEMA } from './importSchema.js'

const headers = ['ID начисления', 'Дата начисления', 'Группа услуг', 'Тип начисления', 'Артикул', 'SKU', 'Название товара', 'Количество', 'Цена продавца', 'Сумма итого, руб.']
const encoder = new TextEncoder()

function xmlEscape(value) { return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') }
function numericEntities(value) { return [...value].map(char => `&#x${char.codePointAt(0).toString(16).toUpperCase()};`).join('') }
function sheetXml({ headerRow = 2, noRefs = false, strings = 'str', reordered = false, rows = 1 } = {}) {
  const ordered = reordered ? [...headers].reverse() : headers
  const prefix = Array.from({ length: headerRow - 1 }, (_, index) => `<row${noRefs ? '' : ` r="${index + 1}"`}><c t="str"><v>${index ? 'Примечание' : 'Период: 01.08.2026-31.08.2026'}</v></c></row>`).join('')
  const shared = [...headers, 'A', 'Продажи', 'Выручка', 'ART', 'SKU', 'Товар']
  const cell = (value, column) => {
    const ref = noRefs ? '' : ` r="${columnName(column)}${headerRow}"`
    if (strings === 'shared') return `<c${ref} t="s"><v>${shared.indexOf(value)}</v></c>`
    if (strings === 'inline') return `<c${ref} t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`
    return `<c${ref} t="str"><v>${xmlEscape(value)}</v></c>`
  }
  const header = `<row${noRefs ? '' : ` r="${headerRow}"`}>${ordered.map(cell).join('')}</row>`
  const values = { 'ID начисления': 'A', 'Дата начисления': '01.08.2026', 'Группа услуг': 'Продажи', 'Тип начисления': 'Выручка', Артикул: 'ART', SKU: 'SKU', 'Название товара': 'Товар', Количество: 1, 'Цена продавца': 100, 'Сумма итого, руб.': 100 }
  const data = Array.from({ length: rows }, (_, rowIndex) => `<row${noRefs ? '' : ` r="${headerRow + 1 + rowIndex}"`}>${ordered.map((heading, column) => {
    const value = values[heading], ref = noRefs ? '' : ` r="${columnName(column)}${headerRow + 1 + rowIndex}"`
    if (typeof value === 'number') return `<c${ref}><v>${value}</v></c>`
    if (strings === 'shared') return `<c${ref} t="s"><v>${shared.indexOf(value)}</v></c>`
    if (strings === 'inline') return `<c${ref} t="inlineStr"><is><t>${value}</t></is></c>`
    return `<c${ref} t="str"><v>${value}</v></c>`
  }).join('')}</row>`).join('')
  return `<worksheet><sheetData>${prefix}${header}${data}</sheetData></worksheet>`
}
function columnName(index) { let result = ''; for (let value = index + 1; value; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result; return result }

function makeXlsx(options = {}) {
  const relId = options.relId || 'rId1', target = options.target || 'worksheets/sheet1.xml', name = options.entityName ? numericEntities(options.sheetName || 'Начисления') : (options.sheetName || 'Начисления')
  const workbook = `<workbook xmlns:r="relationships"><sheets>${options.extraSheet ? '<sheet name="Справка" sheetId="2" r:id="extra"/>' : ''}<sheet name="${name}" sheetId="1" r:id="${relId}"/></sheets></workbook>`
  const rels = `<Relationships>${options.extraSheet ? '<Relationship Id="extra" Target="worksheets/extra.xml"/>' : ''}<Relationship Id="${relId}" Target="${target}"/></Relationships>`
  const files = { 'xl/workbook.xml': workbook, 'xl/_rels/workbook.xml.rels': rels, 'xl/worksheets/sheet1.xml': sheetXml(options), ...(options.extraSheet ? { 'xl/worksheets/extra.xml': '<worksheet><sheetData><row><c t="str"><v>не отчёт</v></c></row></sheetData></worksheet>' } : {}) }
  if (options.strings === 'shared') files['xl/sharedStrings.xml'] = `<sst>${[...headers, 'A', 'Продажи', 'Выручка', 'ART', 'SKU', 'Товар'].map(value => `<si><t>${xmlEscape(value)}</t></si>`).join('')}</sst>`
  return zip(files)
}
function zip(files) {
  const locals = [], centrals = []; let offset = 0
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name), data = encoder.encode(content)
    const local = new Uint8Array(30 + nameBytes.length + data.length), lv = new DataView(local.buffer)
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint32(18, data.length, true); lv.setUint32(22, data.length, true); lv.setUint16(26, nameBytes.length, true); local.set(nameBytes, 30); local.set(data, 30 + nameBytes.length); locals.push(local)
    const central = new Uint8Array(46 + nameBytes.length), cv = new DataView(central.buffer)
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint32(20, data.length, true); cv.setUint32(24, data.length, true); cv.setUint16(28, nameBytes.length, true); cv.setUint32(42, offset, true); central.set(nameBytes, 46); centrals.push(central); offset += local.length
  }
  const centralSize = centrals.reduce((sum, value) => sum + value.length, 0), end = new Uint8Array(22), ev = new DataView(end.buffer)
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, centrals.length, true); ev.setUint16(10, centrals.length, true); ev.setUint32(12, centralSize, true); ev.setUint32(16, offset, true)
  return new Blob([...locals, ...centrals, end])
}
async function load(options = {}, name = 'любой файл (1).xlsx') { const blob = makeXlsx(options); return readExcelFile({ name, size: blob.size, arrayBuffer: () => blob.arrayBuffer() }) }

function makeCostXlsx(sheet, { sheetName = 'OZON', extras = [] } = {}) {
  const allSheets = [...extras.map((name, index) => ({ name, id: `extra${index}`, target: `worksheets/extra${index}.xml` })), { name: sheetName, id: 'cost', target: 'worksheets/cost.xml' }]
  const workbook = `<workbook xmlns:r="relationships"><sheets>${allSheets.map((item, index) => `<sheet name="${item.name}" sheetId="${index + 1}" r:id="${item.id}"/>`).join('')}</sheets></workbook>`
  const rels = `<Relationships>${allSheets.map(item => `<Relationship Id="${item.id}" Target="${item.target}"/>`).join('')}</Relationships>`
  return zip({ 'xl/workbook.xml': workbook, 'xl/_rels/workbook.xml.rels': rels, 'xl/worksheets/cost.xml': `<worksheet><sheetData>${sheet}</sheetData></worksheet>`, ...Object.fromEntries(extras.map((_, index) => [`xl/worksheets/extra${index}.xml`, '<worksheet><sheetData><row><c t="str"><v>другой лист</v></c></row></sheetData></worksheet>'])) })
}
async function loadCost(sheet, options, name = 'OZON WB Себес и маржа 14.09 (1).xlsx') { const blob = makeCostXlsx(sheet, options); return readExcelFile({ name, size: blob.size, arrayBuffer: () => blob.arrayBuffer() }) }
const costHeader = (row = 6, reversed = false, label = 'СЕБЕС (руб)') => `<row r="${row}">${reversed ? `<c r="B${row}" t="str"><v>${label}</v></c><c r="D${row}" t="str"><v>АРТИКУЛ</v></c>` : `<c r="A${row}" t="str"><v>АРТИКУЛ</v></c><c r="F${row}" t="str"><v>${label}</v></c>`}</row>`

describe('real Ozon XLSX structures', () => {
  it('reads a normal one-sheet workbook whose filename contains spaces and (1)', async () => assert.equal(rowsFromSheet(await load(), OZON_ACCRUALS_SCHEMA).length, 1))
  it('decodes a sheet name made from numeric XML entities', async () => assert.ok((await load({ entityName: true })).Начисления))
  it('accepts an arbitrary relationship id and an absolute target', async () => assert.equal(rowsFromSheet(await load({ relId: 'sheet1', target: '/xl/worksheets/sheet1.xml' }), OZON_ACCRUALS_SCHEMA)[0].Артикул, 'ART'))
  it('infers row and column positions when r attributes are absent', async () => assert.equal(rowsFromSheet(await load({ noRefs: true }), OZON_ACCRUALS_SCHEMA)[0]['Сумма итого, руб.'], 100))
  it('reads str, shared strings, and inline strings', async () => { for (const strings of ['str', 'shared', 'inline']) assert.equal(rowsFromSheet(await load({ strings }), OZON_ACCRUALS_SCHEMA)[0]['Название товара'], 'Товар') })
  it('maps reordered columns by normalized header names', async () => assert.equal(rowsFromSheet(await load({ reordered: true }), OZON_ACCRUALS_SCHEMA)[0].Артикул, 'ART'))
  it('finds a header on a row other than the second', async () => assert.equal(rowsFromSheet(await load({ headerRow: 7 }), OZON_ACCRUALS_SCHEMA).length, 1))
  it('ignores unrelated sheets and falls back to headers when the report sheet was renamed', async () => assert.equal(rowsFromSheet(await load({ sheetName: 'Новый отчёт', extraSheet: true }), OZON_ACCRUALS_SCHEMA)[0].SKU, 'SKU'))
  it('imports at least 20,000 operation rows without an XML size limit', async () => assert.equal(rowsFromSheet(await load({ rows: 20000 }), OZON_ACCRUALS_SCHEMA).length, 20000))
  it('keeps an unknown service group, total, and warning in the P&L', () => {
    const accruals = { Начисления: [['Период'], headers, ['1', '', 'Совершенно новая услуга', 'Новый тип', '', '', '', 0, 0, -42]] }
    const costs = { OZON: [['АРТИКУЛ', 'СЕБЕС (руб)'], ['ART', 1]] }
    const report = buildArticleReport(accruals, costs)
    assert.equal(report.serviceGroups['Совершенно новая услуга'], -42); assert.equal(report.totals.accruals, -42); assert.match(report.warnings.join(' '), /Совершенно новая услуга/); assert.match(report.warnings.join(' '), /Новый тип/); assert.deepEqual(report.newCategoryDetails, [{ label: 'Совершенно новая услуга — Новый тип', amount: -42 }])
  })
})

describe('cost reference XLSX regression cases', () => {
  it('skips an empty article row even when it contains t="e" cells', async () => {
    const book = await loadCost(`${costHeader()}<row r="8"><c r="D8" t="e"><v>#N/A</v></c><c r="F8" t="e"><v>#N/A</v></c></row><row r="9"><c r="A9" t="str"><v>OK</v></c><c r="F9"><v>10</v></c></row>`)
    const result = readCostReference(book); assert.equal(result.costs.get('OK'), 10); assert.equal(result.summary.skippedEmptyRows, 1); assert.equal(result.issues.length, 0)
  })
  it('ignores t="e" in an unneeded column', async () => {
    const result = readCostReference(await loadCost(`${costHeader()}<row r="8"><c r="A8" t="str"><v>OK</v></c><c r="D8" t="e"><v>#N/A</v></c><c r="F8"><v>20</v></c></row>`)); assert.equal(result.costs.get('OK'), 20)
  })
  it('reports an article whose cost is #N/A with its cell address', async () => {
    const result = readCostReference(await loadCost(`${costHeader()}<row r="8"><c r="A8" t="str"><v>BAD</v></c><c r="F8" t="e"><v>#N/A</v></c></row>`)); assert.deepEqual(result.issues, [{ article: 'BAD', address: 'F8', error: '#N/A' }]); assert.equal(result.costs.has('BAD'), false)
  })
  it('uses the cached numeric value of a formula', async () => {
    const result = readCostReference(await loadCost(`${costHeader()}<row r="8"><c r="A8" t="str"><v>FORMULA</v></c><c r="F8"><f>VLOOKUP(A8,A:B,2,0)</f><v>720</v></c></row>`)); assert.equal(result.costs.get('FORMULA'), 720)
  })
  it('finds headers outside row 6 and normalizes whitespace and case', async () => {
    const result = readCostReference(await loadCost(`${costHeader(17, false, 'Себестоимость, руб.')}<row r="18"><c r="A18" t="str"><v> A-Б </v></c><c r="F18"><v>42</v></c></row>`)); assert.equal(result.summary.headerRow, 17); assert.equal(result.costs.get('A-Б'), 42)
  })
  it('maps article and cost columns after they are reordered', async () => {
    const result = readCostReference(await loadCost(`${costHeader(4, true)}<row r="5"><c r="B5"><v>99</v></c><c r="D5" t="str"><v>SWAP</v></c></row>`)); assert.equal(result.costs.get('SWAP'), 99)
  })
  it('does not depend on a filename containing spaces and (1)', async () => {
    const book = await loadCost(`${costHeader()}<row r="7"><c r="A7" t="str"><v>NAME</v></c><c r="F7"><v>1</v></c></row>`, {}, 'любой справочник (1).xlsx'); assert.equal(readCostReference(book).costs.get('NAME'), 1)
  })
  it('prefers OZON among WB, Calculator, and other sheets', async () => {
    const result = readCostReference(await loadCost(`${costHeader()}<row r="7"><c r="A7" t="str"><v>MULTI</v></c><c r="F7"><v>2</v></c></row>`, { extras: ['WB', 'Калькулятор', 'Другое'] })); assert.equal(result.summary.sheetName, 'OZON')
  })
  it('continues to import the old valid layout', async () => {
    const book = await loadCost(`${costHeader()}<row r="7"><c r="A7" t="str"><v>OLD</v></c><c r="F7"><v>300</v></c></row>`); assert.equal(rowsFromSheet(book, COST_PRICE_SCHEMA)[0].АРТИКУЛ, 'OLD'); assert.equal(readCostReference(book).costs.get('OLD'), 300)
  })
  it('accepts t="e" without the unsupported-cell-type error and exposes conflicts', async () => {
    const book = await loadCost(`${costHeader()}<row r="7"><c r="A7" t="str"><v>DUP</v></c><c r="F7"><v>10</v></c></row><row r="8"><c r="A8" t="str"><v>DUP</v></c><c r="F8"><v>11</v></c><c r="G8" t="e"><v>#N/A</v></c></row>`)
    const result = readCostReference(book); assert.equal(result.conflicts.length, 1); assert.doesNotMatch(JSON.stringify(result), /Неподдерживаемый тип ячейки/)
  })
})
