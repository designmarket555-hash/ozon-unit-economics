import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { buildArticleReport, ImportError, readExcelFile, rowsFromSheet } from './importer.js'
import { OZON_ACCRUALS_SCHEMA } from './importSchema.js'

const accrualHeaders = ['ID начисления', 'Дата начисления', 'Группа услуг', 'Тип начисления', 'Артикул', 'SKU', 'Название товара', 'Количество', 'Цена продавца', 'Комиссия Ozon, %', 'Сумма итого, руб.']
const accrualWorkbook = {
  Начисления: [
    ['Обезличенный тестовый отчёт'],
    accrualHeaders,
    ['A-001', '2026-08-01', 'Продажи', 'Заказ', 'TEST-100', '900001', 'Тестовый товар А', 2, 1000, 15, 1500],
    ['A-002', '2026-08-02', 'Логистика', 'Доставка', 'TEST-100', '900001', 'Тестовый товар А', 0, 0, 0, -100],
    ['A-003', '2026-08-03', 'Продажи', 'Заказ', 'TEST-200', '900002', 'Тестовый товар Б', 1, 500, 15, 400],
  ],
}
const costWorkbook = {
  OZON: [[], [], [], [], [], ['АРТИКУЛ', 'СЕБЕС (руб)'], ['TEST-100', 300]],
}

describe('Excel import domain', () => {
  it('aggregates accruals and uses uploaded cost without VAT adjustment', () => {
    const result = buildArticleReport(accrualWorkbook, costWorkbook)
    assert.deepEqual(result[0], {
      article: 'TEST-100', sku: '900001', name: 'Тестовый товар А', quantity: 2,
      sales: 2000, accruals: 1400, cost: 600, profit: 800,
      profitability: 800 / 600 * 100, missingCost: false,
    })
  })

  it('marks articles absent from the cost reference', () => {
    const item = buildArticleReport(accrualWorkbook, costWorkbook)[1]
    assert.equal(item.article, 'TEST-200')
    assert.equal(item.missingCost, true)
    assert.equal(item.cost, null)
    assert.equal(item.profit, null)
  })

  it('reports a missing sheet and missing columns clearly', () => {
    assert.throws(() => rowsFromSheet({}, OZON_ACCRUALS_SCHEMA), error => error instanceof ImportError && error.code === 'MISSING_SHEET')
    assert.throws(() => rowsFromSheet({ Начисления: [[], ['Артикул'], ['TEST']] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'MISSING_COLUMNS' && error.message.includes('ID начисления'))
  })

  it('reports a sheet without data rows as empty', () => {
    assert.throws(() => rowsFromSheet({ Начисления: [[], accrualHeaders] }, OZON_ACCRUALS_SCHEMA), error => error.code === 'EMPTY_FILE')
  })

  it('rejects an unsuitable format and an empty xlsx file', async () => {
    await assert.rejects(() => readExcelFile({ name: 'report.csv', size: 10 }), error => error.code === 'INVALID_FILE')
    await assert.rejects(() => readExcelFile({ name: 'report.xlsx', size: 0 }), error => error.code === 'EMPTY_FILE')
  })
})
