export const DEFAULT_INPUTS = {
  price: 2490,
  cost: 780,
  commissionPercent: 15,
  logistics: 185,
  advertisingPercent: 8,
  acquiringPercent: 1.5,
  other: 45,
}

export function calculateUnitEconomics(values) {
  const price = number(values.price)
  const cost = number(values.cost)
  const commission = price * number(values.commissionPercent) / 100
  const logistics = number(values.logistics)
  const advertising = price * number(values.advertisingPercent) / 100
  const acquiring = price * number(values.acquiringPercent) / 100
  const other = number(values.other)
  const expenses = cost + commission + logistics + advertising + acquiring + other
  const profit = price - expenses

  return {
    revenue: price,
    cost,
    commission,
    logistics,
    advertising,
    acquiring,
    other,
    expenses,
    profit,
    salesMargin: price ? profit / price * 100 : 0,
    costReturn: expenses ? profit / expenses * 100 : 0,
  }
}

function number(value) {
  const parsed = Number(String(value).replace(',', '.'))
  return Number.isFinite(parsed) ? Math.max(0, parsed) : 0
}
