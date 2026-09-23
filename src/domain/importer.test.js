import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { buildArticleReport, ImportError, readExcelFile, rowsFromSheet } from './importer.js'
import { OZON_ACCRUALS_SCHEMA } from './importSchema.js'

const accrualHeaders = ['ID начисления', 'Дата начисления', 'Группа услуг', 'Тип начисления', 'Артикул', 'SKU', 'Название товара', 'Количество', 'Цена продавца', 'Комиссия Ozon, %', 'Сумма итого, руб.']
const accrualWorkbook = {
  Начисления: [
    ['Период: 01.08.2026–31.08.2026'],
    accrualHeaders,
    ['A-001', '2026-08-01', 'Продажи', 'Выручка', 'TEST-100', '900001', 'Тестовый товар А', 2, 1000, 15, 1500],
    ['A-002', '2026-08-01', 'Продажи', 'Баллы за скидки', 'TEST-100', '900001', 'Тестовый товар А', 2, 0, 0, 100],
    ['A-003', '2026-08-02', 'Вознаграждение Ozon', 'Комиссия', 'TEST-100', '900001', 'Тестовый товар А', 0, 0, 99, -200],
    ['A-004', '2026-08-03', 'Продажи', 'Выручка', 'TEST-200', '900002', 'Тестовый товар Б', 1, 500, 15, 400],
    ['A-005', '2026-08-03', 'Услуги доставки', 'Доставка', 0, '', '', 0, 0, 0, -100],
  ],
}
const costWorkbook = {
  OZON: [[], [], [], [], [], ['АРТИКУЛ', 'СЕБЕС (руб)'], ['TEST-100', 300]],
}

describe('Excel import domain', () => {
  it('builds the P&L from total amounts and counts only revenue rows', () => {
    const report = buildArticleReport(accrualWorkbook, costWorkbook)
    const item = report.articles.find(row => row.article === 'TEST-100')
    assert.equal(report.period, 'Период: 01.08.2026–31.08.2026')
    assert.equal(report.serviceGroups['Продажи'], 2000)
    assert.equal(report.serviceGroups['Вознаграждение Ozon'], -200)
    assert.deepEqual({ ...report.totals, profitability: Math.round(report.totals.profitability) }, { quantity: 3, revenue: 2000, accruals: 1700, cost: 600, profit: 1100, profitability: 55, revenueWithoutVat: 2000 / 1.22, profitWithoutVat: 1100 / 1.22 })
    assert.deepEqual(item, { article: 'TEST-100', sku: '900001', name: 'Тестовый товар А', quantity: 2, revenue: 1600, ozonReward: -200, otherOzonExpenses: 0, accruals: 1400, cost: 600, profit: 800, profitability: 50, missingCost: false, unassigned: false })
  })

  it('keeps expenses without an article separate but includes them in totals', () => {
    const report = buildArticleReport(accrualWorkbook, costWorkbook)
    const common = report.articles.find(row => row.unassigned)
    assert.equal(common.name, 'Общие расходы без артикула')
    assert.equal(common.accruals, -100)
    assert.equal(common.profitability, null)
    assert.equal(report.totals.accruals, 1700)
  })

  it('marks articles absent from the cost reference', () => {
    const item = buildArticleReport(accrualWorkbook, costWorkbook).articles.find(row => row.article === 'TEST-200')
    assert.equal(item.missingCost, true)
    assert.equal(item.cost, null)
    assert.equal(item.profit, null)
  })

  it('does not use the optional commission percentage as a financial amount', () => {
    const report = buildArticleReport(accrualWorkbook, costWorkbook)
    assert.equal(report.articles.find(row => row.article === 'TEST-100').ozonReward, -200)
  })

  it('matches the supplied control P&L figures', () => {
    const workbook = { Начисления: [
      ['Контрольный период'], accrualHeaders,
      ['1', '', 'Продажи', 'Выручка', 'CONTROL', '', '', 927, 0, 0, 1568052],
      ['2', '', 'Вознаграждение Ozon', 'Комиссия', 'CONTROL', '', '', 0, 0, 0, -300000],
      ['3', '', 'Услуги доставки', 'Доставка', 'CONTROL', '', '', 0, 0, 0, -551671.74],
    ] }
    const costs = { OZON: [[], [], [], [], [], ['АРТИКУЛ', 'СЕБЕС (руб)'], ['CONTROL', 367637.94 / 927]] }
    const { totals } = buildArticleReport(workbook, costs)
    assert.equal(totals.quantity, 927)
    assert.equal(totals.revenue, 1568052)
    assert.ok(Math.abs(totals.accruals - 716380.26) < 0.001)
    assert.ok(Math.abs(totals.cost - 367637.94) < 0.001)
    assert.ok(Math.abs(totals.profit - 348742.32) < 0.001)
    assert.ok(Math.abs(totals.profitability - 22.24) < 0.01)
    assert.ok(Math.abs(totals.revenueWithoutVat - 1285288.52) < 0.01)
    assert.ok(Math.abs(totals.profitWithoutVat - 285854.36) < 0.01)
  })

  it('reports a missing sheet and missing columns clearly', () => {
    assert.throws(() => rowsFromSheet({}, OZON_ACCRUALS_SCHEMA), error => error instanceof ImportError && error.code === 'MISSING_SHEET')
    assert.throws(() => rowsFromSheet({ Начисления: [[], ['Артикул'], ['TEST']] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'MISSING_COLUMNS' && error.message.includes('ID начисления'))
    assert.throws(() => rowsFromSheet({ Начисления: [[], accrualHeaders.filter(header => header !== 'Сумма итого, руб.'), ['TEST']] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'MISSING_COLUMNS' && error.message.includes('Сумма итого, руб.'))
  })

  it('reports a sheet without data rows as empty', () => {
    assert.throws(() => rowsFromSheet({ Начисления: [[], accrualHeaders] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'EMPTY_FILE')
  })

  it('rejects an unsuitable format and an empty xlsx file', async () => {
    await assert.rejects(() => readExcelFile({ name: 'report.csv', size: 10 }), error => error.code === 'INVALID_FILE')
    await assert.rejects(() => readExcelFile({ name: 'report.xlsx', size: 0 }), error => error.code === 'EMPTY_FILE')
  })
})
