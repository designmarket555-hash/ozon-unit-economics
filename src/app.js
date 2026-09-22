import { calculateUnitEconomics, DEFAULT_INPUTS } from './domain/calculator.js'
import { COST_PRICE_SCHEMA, OZON_ACCRUALS_SCHEMA, PERIOD_STATUSES } from './domain/importSchema.js'
import { buildArticleReport, readExcelFile, rowsFromSheet } from './domain/importer.js'

const root = document.querySelector('#root')
const money = new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 })
const decimal = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 1 })
let section = 'calculator'
let inputs = { ...DEFAULT_INPUTS }
let periodStatus = 'Закрытый'
const files = { ozon: null, cost: null }
const fileStates = { ozon: {}, cost: {} }
let report = null
let importError = ''

function layout(content) {
  return `<div class="app-shell">
    <aside class="sidebar">
      <div class="brand"><div class="brand-mark">O</div><div><strong>Ozon P&amp;L</strong><span>юнит-экономика</span></div></div>
      <nav>
        <button data-section="calculator" class="${section === 'calculator' ? 'active' : ''}"><i>⌗</i>Калькулятор цены</button>
        <button data-section="pnl" class="${section === 'pnl' ? 'active' : ''}"><i>▥</i>Факт P&amp;L Ozon</button>
      </nav>
      <div class="sidebar-note"><i>✓</i><div><strong>Данные остаются у вас</strong><p>Файлы обрабатываются локально и не публикуются.</p></div></div>
      <div class="user"><div class="avatar">А</div><div><strong>Аналитик</strong><span>Рабочее пространство</span></div><i>⌄</i></div>
    </aside>
    <main class="main-content">${content}</main>
  </div>`
}

function header(eyebrow, title, description, badge = '') {
  return `<header class="page-header"><div><span class="eyebrow">${eyebrow}</span><div class="title-line"><h1>${title}</h1>${badge ? `<span class="badge">${badge}</span>` : ''}</div><p>${description}</p></div><div class="period"><span>Рабочий период</span><strong>Август 2026 <i>⌄</i></strong></div></header>`
}

function field(key, label, suffix, options = '') {
  return `<label class="field ${options.includes('wide') ? 'wide' : ''}"><span>${label}</span><div class="${options.includes('highlight') ? 'highlight' : ''}"><input data-input="${key}" aria-label="${label}" type="number" min="0" step="any" value="${inputs[key]}"><b>${suffix}</b></div></label>`
}

function calculatorPage() {
  const result = calculateUnitEconomics(inputs)
  return layout(`${header('Плановая модель', 'Калькулятор цены', 'Оцените экономику одной продажи до запуска товара или изменения цены.')}
    <div class="calculator-grid">
      <section class="card input-card">
        <div class="card-heading"><div><span class="step">01</span><h2>Параметры продажи</h2></div><button class="text-button" id="reset">↻ Сбросить</button></div>
        <div class="form-grid">
          ${field('price', 'Цена продажи', '₽', 'highlight')}${field('cost', 'Себестоимость товара', '₽')}
          ${field('commissionPercent', 'Комиссия Ozon', '%')}${field('logistics', 'Логистика', '₽')}
          ${field('advertisingPercent', 'Реклама', '% от цены')}${field('acquiringPercent', 'Эквайринг', '%')}
          ${field('other', 'Прочие расходы', '₽', 'wide')}
        </div>
        <div class="helper">ⓘ Процентные расходы рассчитываются от цены продажи.</div>
      </section>
      <section class="results">
        <div class="result-hero"><div class="result-label"><span>Прибыль с продажи</span><i>₽</i></div><strong data-result="profit">${money.format(result.profit)}</strong><div class="profit-bar"><span style="width:${Math.max(0, Math.min(100, result.salesMargin))}%"></span></div><div class="hero-meta"><span>Выручка ${money.format(result.revenue)}</span><span>${decimal.format(result.salesMargin)}% маржа</span></div></div>
        <div class="metric-grid"><div class="metric"><span>Маржинальность продаж</span><strong>${decimal.format(result.salesMargin)}%</strong><small>Прибыль / выручка</small></div><div class="metric"><span>Рентабельность затрат</span><strong>${decimal.format(result.costReturn)}%</strong><small>Прибыль / расходы</small></div></div>
        <div class="card breakdown"><div class="card-heading"><div><span class="step">02</span><h2>Структура расчёта</h2></div></div>
          ${row('Выручка', result.revenue, true)}${row('Себестоимость', -result.cost)}${row('Комиссия Ozon', -result.commission)}${row('Логистика', -result.logistics)}${row('Реклама', -result.advertising)}${row('Эквайринг', -result.acquiring)}${row('Прочие расходы', -result.other)}
          <div class="total"><span>Итого расходов</span><strong>${money.format(result.expenses)}</strong></div>
        </div>
      </section>
    </div>`)
}

function row(name, value, positive = false) {
  return `<div class="row"><span>${name}</span><strong class="${positive ? 'positive' : ''}">${value > 0 ? '+' : '−'} ${money.format(Math.abs(value))}</strong></div>`
}

function uploadCard(key, index, title, description, schema) {
  const file = files[key]
  const state = fileStates[key]
  return `<article class="upload-card ${file ? 'has-file' : ''} ${state.error ? 'has-error' : ''}"><div class="upload-top"><span class="file-index">${index}</span><div><h3>${title}</h3><p>${description}</p></div></div>
    <label class="dropzone"><input data-file="${key}" type="file" accept=".xlsx"><span class="upload-icon">⇧</span><strong>${file ? escapeHtml(file.name) : 'Выберите Excel-файл'}</strong><small>${state.loading ? 'Читаем файл…' : file ? 'Файл проверен · можно заменить' : 'формат .xlsx'}</small></label>
    ${state.error ? `<p class="file-error" role="alert">⚠ ${escapeHtml(state.error)}</p>` : ''}
    <div class="schema"><div><span>Лист</span><strong>${schema.sheet}</strong></div><div><span>Строка заголовков</span><strong>${schema.headerRow}</strong></div>${schema.matchKey ? `<div><span>Ключ</span><strong>${schema.matchKey}</strong></div>` : ''}</div>
    ${schema.ignoredSheets ? '<p class="ignore-note">ⓘ Лист «Начисления (2)» будет исключён</p>' : ''}</article>`
}

function pnlPage() {
  return layout(`${header('Фактические данные', 'Факт P&amp;L Ozon', 'Загрузите два Excel-файла: обработка полностью выполняется в вашем браузере.', 'Локальный импорт')}
    <div class="period-bar card"><div><label>Отчётный период</label><div class="date-range"><span>01.08.2026</span><i>—</i><span>31.08.2026</span></div></div><div><label>Статус периода</label><div class="segmented">${PERIOD_STATUSES.map(item => `<button data-status="${item}" class="${periodStatus === item ? 'selected' : ''}">${item === 'Закрытый' ? '▣ ' : ''}${item}</button>`).join('')}</div></div><div class="source-pill"><span>Источник данных</span><strong>▤ Excel-отчёт Ozon</strong></div></div>
    <section class="upload-section"><div class="section-title"><span class="step">01</span><div><h2>Загрузите исходные файлы</h2><p>Реальные данные не сохраняются в репозитории</p></div></div><div class="upload-grid">${uploadCard('ozon', 1, 'Отчёт Ozon по начислениям', 'Основной источник выручки и расходов', OZON_ACCRUALS_SCHEMA)}${uploadCard('cost', 2, 'Справочник себестоимости', 'Сопоставление по артикулу продавца', COST_PRICE_SCHEMA)}</div></section>
    <section class="import-actions"><div><strong>Себестоимость используется без корректировки НДС</strong><span>Берём значение «СЕБЕС (руб)» ровно из справочника.</span></div><button id="run-import" ${files.ozon && files.cost && !fileStates.ozon.loading && !fileStates.cost.loading ? '' : 'disabled'}>Сформировать отчёт</button></section>
    ${importError ? `<div class="global-error" role="alert">⚠ ${escapeHtml(importError)}</div>` : ''}
    ${report ? reportView(report) : ''}`)
}

function reportView(rows) {
  const missing = rows.filter(row => row.missingCost)
  return `<section class="report-section"><div class="section-title"><span class="step">02</span><div><h2>Результат по артикулам</h2><p>${rows.length} товаров · ${missing.length} без себестоимости</p></div></div>
    <div class="report-table card"><table><thead><tr><th>Артикул</th><th>SKU и название</th><th>Количество</th><th>Продажи</th><th>Начисления Ozon</th><th>Себестоимость</th><th>Прибыль</th><th>Рентабельность</th><th>Статус</th></tr></thead><tbody>${rows.map(item => `<tr class="${item.missingCost ? 'missing-row' : ''}"><td><strong>${escapeHtml(item.article)}</strong></td><td><strong>${escapeHtml(item.sku || '—')}</strong><small>${escapeHtml(item.name || 'Без названия')}</small></td><td>${decimal.format(item.quantity)}</td><td>${money.format(item.sales)}</td><td>${money.format(item.accruals)}</td><td>${item.cost == null ? '—' : money.format(item.cost)}</td><td class="${item.profit != null && item.profit >= 0 ? 'profit' : 'loss'}">${item.profit == null ? '—' : money.format(item.profit)}</td><td>${item.profitability == null ? '—' : `${decimal.format(item.profitability)}%`}</td><td><span class="status ${item.missingCost ? 'missing' : 'found'}">${item.missingCost ? '⚠ Не найдена' : '✓ Найдена'}</span></td></tr>`).join('')}</tbody></table></div>
    ${missing.length ? `<aside class="missing-list"><h3>Товары без найденной себестоимости</h3><p>Добавьте эти артикулы на лист «OZON» и загрузите справочник повторно.</p><div>${missing.map(item => `<span>${escapeHtml(item.article)}${item.name ? ` · ${escapeHtml(item.name)}` : ''}</span>`).join('')}</div></aside>` : '<div class="all-found">✓ Себестоимость найдена для всех товаров.</div>'}</section>`
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char])
}

function render() {
  root.innerHTML = section === 'calculator' ? calculatorPage() : pnlPage()
  root.querySelectorAll('[data-section]').forEach(button => button.addEventListener('click', () => { section = button.dataset.section; render() }))
  root.querySelectorAll('[data-input]').forEach(input => input.addEventListener('change', () => { inputs[input.dataset.input] = input.value; render() }))
  root.querySelector('#reset')?.addEventListener('click', () => { inputs = { ...DEFAULT_INPUTS }; render() })
  root.querySelectorAll('[data-status]').forEach(button => button.addEventListener('click', () => { periodStatus = button.dataset.status; render() }))
  root.querySelectorAll('[data-file]').forEach(input => input.addEventListener('change', async () => {
    const key = input.dataset.file
    const file = input.files[0] || null
    files[key] = file
    fileStates[key] = file ? { loading: true } : {}
    report = null
    importError = ''
    render()
    if (!file) return
    try {
      const workbook = await readExcelFile(file)
      rowsFromSheet(workbook, key === 'ozon' ? OZON_ACCRUALS_SCHEMA : COST_PRICE_SCHEMA)
      fileStates[key] = { workbook }
    } catch (error) { files[key] = null; fileStates[key] = { error: error.message } }
    render()
  }))
  root.querySelector('#run-import')?.addEventListener('click', () => {
    try {
      report = buildArticleReport(fileStates.ozon.workbook, fileStates.cost.workbook)
      importError = report.length ? '' : 'В отчёте нет строк с заполненным артикулом.'
    } catch (error) { report = null; importError = error.message }
    render()
  })
}

render()
