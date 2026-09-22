import { calculateUnitEconomics, DEFAULT_INPUTS } from './domain/calculator.js'
import { COST_PRICE_SCHEMA, OZON_ACCRUALS_SCHEMA, PERIOD_STATUSES } from './domain/importSchema.js'

const root = document.querySelector('#root')
const money = new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 })
const decimal = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 1 })
let section = 'calculator'
let inputs = { ...DEFAULT_INPUTS }
let periodStatus = 'Закрытый'
let vatMode = 'included'
const files = { ozon: null, cost: null }

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
  return `<article class="upload-card ${file ? 'has-file' : ''}"><div class="upload-top"><span class="file-index">${index}</span><div><h3>${title}</h3><p>${description}</p></div></div>
    <label class="dropzone"><input data-file="${key}" type="file" accept=".xlsx,.xls"><span class="upload-icon">⇧</span><strong>${file ? file.name : 'Выберите Excel-файл'}</strong><small>${file ? 'Файл выбран · можно заменить' : 'формат .xlsx или .xls'}</small></label>
    <div class="schema"><div><span>Лист</span><strong>${schema.sheet}</strong></div><div><span>Строка заголовков</span><strong>${schema.headerRow}</strong></div>${schema.matchKey ? `<div><span>Ключ</span><strong>${schema.matchKey}</strong></div>` : ''}</div>
    ${schema.ignoredSheets ? '<p class="ignore-note">ⓘ Лист «Начисления (2)» будет исключён</p>' : ''}</article>`
}

function pnlPage() {
  return layout(`${header('Фактические данные', 'Факт P&amp;L Ozon', 'Подготовьте отчёты закрытого периода для сверки с учётной моделью.', 'Импорт · этап 1')}
    <div class="period-bar card"><div><label>Отчётный период</label><div class="date-range"><span>01.08.2026</span><i>—</i><span>31.08.2026</span></div></div><div><label>Статус периода</label><div class="segmented">${PERIOD_STATUSES.map(item => `<button data-status="${item}" class="${periodStatus === item ? 'selected' : ''}">${item === 'Закрытый' ? '▣ ' : ''}${item}</button>`).join('')}</div></div><div class="source-pill"><span>Источник данных</span><strong>▤ Excel-отчёт Ozon</strong></div></div>
    <section class="upload-section"><div class="section-title"><span class="step">01</span><div><h2>Загрузите исходные файлы</h2><p>Реальные данные не сохраняются в репозитории</p></div></div><div class="upload-grid">${uploadCard('ozon', 1, 'Отчёт Ozon по начислениям', 'Основной источник выручки и расходов', OZON_ACCRUALS_SCHEMA)}${uploadCard('cost', 2, 'Справочник себестоимости', 'Сопоставление по артикулу продавца', COST_PRICE_SCHEMA)}</div></section>
    <section class="card vat-card"><div class="section-title"><span class="step">02</span><div><h2>Настройка себестоимости и НДС</h2><p>Единый режим будет применён ко всем значениям «СЕБЕС (руб)»</p></div></div><div class="vat-content"><div class="vat-choice">${vatChoice('included', 'Себестоимость указана с НДС', 'Значение в файле уже включает налог')}${vatChoice('excluded', 'Себестоимость указана без НДС', 'Налог не включён в значение файла')}</div><div class="vat-rate"><label>Ставка НДС себестоимости</label><div><input value="22" readonly><span>%</span></div><small>Настройка фиксируется вместе с периодом</small></div></div><div class="warning">ⓘ <span><strong>Суммы не смешиваются автоматически.</strong> Режим НДС будет явно отображаться в расчёте и экспорте.</span></div></section>
    <div class="future-note"><div class="future-icon">✓</div><div><span>Архитектура источников</span><strong>Готово к подключению Ozon API</strong><p>В будущем здесь появится второй источник только для чтения. Управление ценами, остатками, товарами и рекламой не предусмотрено.</p></div><em>Позже</em></div>`)
}

function vatChoice(value, title, note) {
  return `<label class="${vatMode === value ? 'checked' : ''}"><input data-vat="${value}" type="radio" name="vat" ${vatMode === value ? 'checked' : ''}><span><strong>${title}</strong><small>${note}</small></span></label>`
}

function render() {
  root.innerHTML = section === 'calculator' ? calculatorPage() : pnlPage()
  root.querySelectorAll('[data-section]').forEach(button => button.addEventListener('click', () => { section = button.dataset.section; render() }))
  root.querySelectorAll('[data-input]').forEach(input => input.addEventListener('change', () => { inputs[input.dataset.input] = input.value; render() }))
  root.querySelector('#reset')?.addEventListener('click', () => { inputs = { ...DEFAULT_INPUTS }; render() })
  root.querySelectorAll('[data-status]').forEach(button => button.addEventListener('click', () => { periodStatus = button.dataset.status; render() }))
  root.querySelectorAll('[data-vat]').forEach(input => input.addEventListener('change', () => { vatMode = input.dataset.vat; render() }))
  root.querySelectorAll('[data-file]').forEach(input => input.addEventListener('change', () => { files[input.dataset.file] = input.files[0] || null; render() }))
}

render()
