import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { buildArticleReport, ImportError, readExcelFile, rowsFromSheet } from './importer.js'
import { OZON_ACCRUALS_SCHEMA } from './importSchema.js'

const headers = ['ID начисления', 'Дата начисления', 'Группа услуг', 'Тип начисления', 'Артикул', 'SKU', 'Название товара', 'Количество', 'Цена продавца', 'Комиссия Ozon, %', 'Сумма итого, руб.']
const costs = values => ({ OZON: [[], ['АРТИКУЛ', 'СЕБЕС (руб)'], ...values] })
const workbook = rows => ({ AnySheet: [['Период 2026'], headers, ...rows] })
const row = (id, group, type, article, quantity, amount) => [id, '2026-01-01', group, type, article, 'SKU-1', 'Товар', quantity, 0, 0, amount]

describe('P&L Ozon', () => {
  it('считает обычную продажу, НДС, валовую и маржинальную прибыль', () => {
    const report = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 2, 1220), row('2', 'Вознаграждение Ozon', 'Комиссия', 'A', 0, -100)]), costs([['A', 300]]))
    assert.deepEqual({ sales: report.totals.salesGross, vat: report.totals.vat, revenue: report.totals.netRevenue, cost: report.totals.cost, gross: report.totals.grossProfit, expenses: report.totals.ozonExpenses, margin: report.totals.marginalProfit }, { sales: 1220, vat: 220, revenue: 1000, cost: 600, gross: 400, expenses: 100, margin: 300 })
  })

  it('вычитает возврат из продажи, НДС, количества и себестоимости', () => {
    const report = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 3, 1830), row('2', 'Возвраты', 'Возврат выручки', 'A', -1, -610)]), costs([['A', 200]]))
    assert.deepEqual({ returns: report.totals.returnsGross, net: report.totals.netRevenueGross, quantity: report.totals.netQuantity, cost: report.totals.cost }, { returns: 610, net: 1220, quantity: 2, cost: 400 })
  })

  it('оставляет расход без артикула нераспределённым, но включает в магазин', () => {
    const report = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 1, 122), row('2', 'Услуги доставки', 'Доставка', 0, 0, -20)]), costs([['A', 50]]))
    assert.equal(report.articles.find(item => item.unassigned).marginalProfit, -20)
    assert.equal(report.totals.marginalProfit, 30)
    assert.equal(report.checks.skuPlusUnassigned, true)
  })

  it('добавляет знаковые компенсации отдельно', () => {
    const report = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 1, 122), row('2', 'Прочие начисления', 'Компенсация', 'A', 0, 30), row('3', 'Прочие начисления', 'Компенсация', 'A', 0, -5)]), costs([['A', 50]]))
    assert.equal(report.totals.compensations, 25)
    assert.equal(report.totals.marginalProfit, 75)
  })

  it('не подставляет 0 вместо отсутствующей себестоимости', () => {
    const report = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 1, 122)]), costs([]))
    assert.equal(report.articles[0].cost, null); assert.equal(report.totals.cost, null); assert.equal(report.totals.complete, false); assert.equal(report.totals.costCoveragePercent, 0)
  })

  it('исключает дублирующуюся строку по ID', () => {
    const duplicate = row('1', 'Продажи', 'Выручка', 'A', 1, 122)
    const report = buildArticleReport(workbook([duplicate, duplicate]), costs([['A', 50]]))
    assert.equal(report.totals.salesGross, 122); assert.equal(report.checks.noDuplicateRows, false)
  })

  it('сохраняет новый тип начисления в unknown один раз', () => {
    const report = buildArticleReport(workbook([row('1', 'Новая услуга', 'Новый тип', 'A', 0, -17)]), costs([]))
    assert.equal(report.expenseCategories.unknown, 17); assert.equal(report.unknownCategories[0].amount, -17); assert.equal(report.checks.unknownCategoriesKept, true)
  })

  it('корректно обрабатывает нулевую выручку и отрицательную маржинальную прибыль', () => {
    const zero = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 0, 0)]), costs([]))
    assert.equal(zero.totals.marginalMargin, null)
    const loss = buildArticleReport(workbook([row('1', 'Продажи', 'Выручка', 'A', 1, 122), row('2', 'Услуги доставки', 'Доставка', 'A', 0, -80)]), costs([['A', 50]]))
    assert.equal(loss.totals.marginalProfit, -30); assert.equal(loss.totals.marginalMargin, -30)
  })
})

describe('Excel validation', () => {
  it('понятно отклоняет пустую книгу и неверные колонки', () => {
    assert.throws(() => rowsFromSheet({}, OZON_ACCRUALS_SCHEMA), error => error instanceof ImportError && error.code === 'EMPTY_FILE')
    assert.throws(() => rowsFromSheet({ X: [['Артикул'], ['A']] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'MISSING_COLUMNS')
    assert.throws(() => rowsFromSheet({ X: [headers] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'EMPTY_FILE')
  })
  it('отклоняет неверный формат и пустой xlsx', async () => {
    await assert.rejects(() => readExcelFile({ name: 'report.csv', size: 10 }), error => error.code === 'INVALID_FILE')
    await assert.rejects(() => readExcelFile({ name: 'report.xlsx', size: 0 }), error => error.code === 'EMPTY_FILE')
  })
})
