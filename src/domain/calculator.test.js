import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { calculateUnitEconomics, DEFAULT_INPUTS } from './calculator.js'

describe('calculateUnitEconomics', () => {
  it('calculates all proportional and fixed costs', () => {
    const result = calculateUnitEconomics(DEFAULT_INPUTS)
    assert.equal(result.commission, 373.5)
    assert.equal(result.advertising, 199.2)
    assert.equal(result.acquiring, 37.35)
    assert.ok(Math.abs(result.expenses - 1620.05) < 0.001)
    assert.ok(Math.abs(result.profit - 869.95) < 0.001)
    assert.ok(Math.abs(result.salesMargin - 34.94) < 0.01)
    assert.ok(Math.abs(result.costReturn - 53.7) < 0.1)
  })

  it('handles empty and invalid input safely', () => {
    const result = calculateUnitEconomics({ price: '', cost: 'text' })
    assert.equal(result.profit, 0)
    assert.equal(result.salesMargin, 0)
    assert.equal(result.costReturn, 0)
  })

  it('recalculates the result when the selling price changes', () => {
    const initial = calculateUnitEconomics(DEFAULT_INPUTS)
    const changed = calculateUnitEconomics({ ...DEFAULT_INPUTS, price: 3000 })
    assert.equal(initial.profit, 869.95)
    assert.equal(changed.profit, 1255)
    assert.notEqual(changed.salesMargin, initial.salesMargin)
  })
})
