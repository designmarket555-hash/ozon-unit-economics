export const OZON_ACCRUALS_SCHEMA = {
  sourceName: 'Excel-отчёт Ozon',
  sheet: 'Начисления',
  fields: ['ID начисления', 'Дата начисления', 'Группа услуг', 'Тип начисления', 'Артикул', 'SKU', 'Название товара', 'Количество', 'Цена продавца', 'Сумма итого, руб.'],
  optionalFields: ['Комиссия Ozon, %'],
}

export const COST_PRICE_SCHEMA = {
  sourceName: 'Справочник себестоимости',
  sheet: 'OZON',
  headerRow: 6,
  matchKey: 'АРТИКУЛ',
  costField: 'СЕБЕС (руб)',
}

export const PERIOD_STATUSES = ['Предварительный', 'Закрытый']
