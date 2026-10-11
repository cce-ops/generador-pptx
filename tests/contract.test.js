import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import vm from 'node:vm'

const src = fs.readFileSync(new URL('../app.js', import.meta.url), 'utf8')

function load() {
  const sandbox = { console }
  vm.createContext(sandbox)
  vm.runInContext(src, sandbox, { filename: 'app.js' })
  return sandbox.__generador
}

// Los objetos creados dentro del vm tienen otro prototipo; se clonan al
// realm actual antes de comparar en profundidad.
const plain = v => structuredClone(v)

test('expone el contrato en __generador', () => {
  const g = load()
  for (const fn of ['buildPrompt', 'extractJson', 'normalizeSlides', 'truncateText']) {
    assert.equal(typeof g[fn], 'function', `${fn} debe existir`)
  }
})

test('buildPrompt incluye el esquema JSON y el tope', () => {
  const g = load()
  const p = g.buildPrompt('mis apuntes', 15)
  assert.ok(p.includes('mis apuntes'))
  assert.ok(p.includes('15'))
  assert.ok(p.includes('"slides"'))
  assert.ok(p.includes('12 palabras'))
})

test('extractJson acepta JSON directo, con fences y con prosa', () => {
  const g = load()
  assert.deepEqual(plain(g.extractJson('{"slides": []}')), { slides: [] })
  assert.deepEqual(plain(g.extractJson('```json\n{"slides": []}\n```')), { slides: [] })
  assert.deepEqual(plain(g.extractJson('Aquí tienes:\n{"slides": []}\nEspero que sirva.')), { slides: [] })
})

test('extractJson rechaza basura con mensaje accionable', () => {
  const g = load()
  assert.throws(() => g.extractJson('hola sin llaves'), /Respuesta sin JSON válido/)
})

test('normalizeSlides trunca al tope y sanea formas', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [
    { title: 'A', bullets: ['x'], table: { header: ['H'], rows: [['c']] } },
    { title: 'B' },
    { title: 'C', bullets: 'no-es-array', table: { header: 'roto' } },
  ] }, 2)
  assert.equal(slides.length, 2)
  assert.equal(slides[0].kind, 'cover')
  assert.deepEqual(plain(slides[1].bullets), ['x'])
  assert.deepEqual(plain(slides[1].table), { header: ['H'], rows: [['c']] })
})

test('normalizeSlides elimina tablas malformadas y exige lista', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [{ title: 'C', table: { header: 'roto' } }] }, 5)
  assert.ok(!('table' in slides[0]))
  assert.throws(() => g.normalizeSlides({}, 5), /lista de diapositivas/)
})

test('truncateText corta a MAX_CHARS avisando', () => {
  const g = load()
  assert.deepEqual(plain(g.truncateText('corto')), { text: 'corto', truncated: false })
  const r = g.truncateText('x'.repeat(70000))
  assert.equal(r.text.length, 60000)
  assert.equal(r.truncated, true)
})

test('normalizeSlides trunca bullets a 12 palabras, no 12 caracteres', () => {
  const g = load()
  const long = 'uno dos tres cuatro cinco seis siete ocho nueve diez once doce trece catorce quince dieciseis diecisiete dieciocho diecinueve veinte'
  const slides = g.normalizeSlides({ slides: [
    { kind: 'content', title: 'T', bullets: [long] },
  ] }, 5)
  const b = slides.find(s => s.kind === 'content').bullets[0]
  assert.ok(b.split(/\s+/).length <= 12, b)
  assert.ok(b.length > 30, b)
  assert.ok(b.endsWith('…'), b)
})

test('normalizeSlides limpia strings vacios y descarta lo que quede vacio', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [
    { kind: 'content', title: '  Espacios  ', bullets: ['   ', 'real'] },
    { kind: 'stats', title: 'S', items: [{ value: '  ', label: '' }, { value: '1', label: 'uno' }] },
    { kind: 'process', title: 'P', steps: [['', ''], ['A', 'b']] },
  ] }, 10)
  const content = slides.find(s => s.kind === 'content' && s.title === 'Espacios')
  assert.ok(content, 'content conservada por tener bullet real')
  assert.deepEqual(plain(content.bullets), ['real'])
  assert.equal(content.title, 'Espacios')
  const stats = slides.find(s => s.kind === 'stats')
  assert.equal(stats.items.length, 1)
})

test('normalizeSlides descarta items stats/compare sin valor', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [
    { kind: 'stats', title: 'S', items: [{ value: '', label: 'vacio' }, { value: '1', label: 'uno' }] },
    { kind: 'compare', title: 'C', items: [{ value: '  ', label: 'x' }, { value: 'A', label: 'a' }] },
  ] }, 10)
  assert.equal(slides.find(s => s.kind === 'stats').items.length, 1)
  assert.equal(slides.find(s => s.kind === 'compare').compare.length, 1)
})

test('splitKeywords parte **negrita** en segmentos', () => {
  const g = load()
  assert.deepEqual(plain(g.splitKeywords('hola **mundo** adios')), [
    { text: 'hola ', key: false },
    { text: 'mundo', key: true },
    { text: ' adios', key: false },
  ])
  assert.deepEqual(plain(g.splitKeywords('sin marcas')), [{ text: 'sin marcas', key: false }])
})

test('normalizeSlides descarta slides sin contenido renderizable', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [
    { kind: 'content', title: 'Vacia', bullets: [] },
    { kind: 'stats', title: 'Sin items', items: [] },
    { kind: 'process', title: 'Sin pasos', steps: [] },
    { kind: 'content', title: 'Llena', bullets: ['a', 'b', 'c'] },
  ] }, 10)
  const titles = slides.map(s => s.title)
  assert.ok(!titles.includes('Vacia'))
  assert.ok(!titles.includes('Sin items'))
  assert.ok(!titles.includes('Sin pasos'))
  assert.ok(titles.includes('Llena'))
})

test('orderModels empieza por el elegido y sigue la lista', () => {
  const g = load()
  assert.deepEqual(plain(g.orderModels(['a', 'b', 'c'], 'b')), ['b', 'a', 'c'])
  assert.deepEqual(plain(g.orderModels(['a', 'b'], 'z')), ['z', 'a', 'b'])
})

test('buildPrompt incluye glosario de kinds y reglas estrictas', () => {
  const g = load()
  const p = g.buildPrompt('apuntes', 10)
  assert.ok(p.includes('kind'))
  assert.ok(p.includes('cover'))
  assert.ok(p.includes('stats'))
  assert.ok(p.includes('process'))
  assert.ok(p.includes('table'))
  assert.ok(p.includes('formula'))
  assert.ok(p.includes('compare'))
  assert.ok(p.includes('content'))
  assert.ok(p.includes('closing'))
  assert.ok(p.includes('1 idea visual por slide'))
  assert.ok(p.includes('3-5 bullets'))
  assert.ok(p.includes('3 kinds distintos'))
  assert.ok(p.includes('vacíos'))
  assert.ok(p.includes('doble asterisco'))
  assert.ok(p.includes('12 palabras'))
  assert.ok(p.includes('notes'))
  assert.ok(p.includes('5 palabras'), 'sin regla de value corto en compare')
  assert.ok(p.includes('chart'), 'sin kind chart en el contrato')
  assert.ok(p.includes('bar|line|pie|doughnut'), 'sin tipos de grafico en el esquema')
  assert.ok(p.includes('Si hay series de datos'), 'sin regla para usar chart cuando hay series')
  assert.ok(p.includes('axisTitle'), 'sin campo axisTitle en el esquema del chart')
  assert.ok(p.includes('magnitud y sus unidades'), 'sin regla de unidades en el eje de valores')
})

test('buildPrompt añade sección obligatoria con las instrucciones del usuario', () => {
  const g = load()
  const p = g.buildPrompt('apuntes', 10, 'Debe aparecer: huella de carbono, ISO 14006')
  assert.ok(p.includes('OBLIGATORIOS'), 'sin sección obligatoria')
  assert.ok(p.includes('no te los puedes saltar'), 'sin lenguaje de obligatoriedad')
  assert.ok(p.includes('huella de carbono, ISO 14006'), 'sin el texto del usuario')
})

test('buildPrompt omite la sección si las instrucciones estan vacias', () => {
  const g = load()
  for (const v of [undefined, null, '', '   ', '\n\t ']) {
    const p = g.buildPrompt('apuntes', 10, v)
    assert.ok(!p.includes('OBLIGATORIOS'), 'sección presente con instrucciones=' + JSON.stringify(v))
  }
})

test('buildPrompt trunca instrucciones a 2000 caracteres', () => {
  const g = load()
  const p = g.buildPrompt('apuntes', 10, 'x'.repeat(3000))
  assert.ok(p.includes('x'.repeat(2000) + '…'), 'sin truncar')
  assert.ok(!p.includes('x'.repeat(2001)), 'pasó de 2000')
})

test('buildPrompt pide exactamente N diapositivas contando portada y cierre', () => {
  const g = load()
  const p = g.buildPrompt('apuntes', 15)
  assert.ok(p.includes('exactamente 15 diapositivas'), 'sin exact count')
  assert.ok(p.includes('incluyendo la portada y el cierre'), 'sin aclarar qué cuenta')
  assert.ok(!p.includes('como máximo'), 'sigue pidiendo máximo')
})

test('normalizeSlides rellena títulos vacíos con fallback por kind', () => {
  const g = load()
  const f = g.normalizeSlides({ slides: [
    { kind: 'cover', title: 'Portada', subtitle: '' },
    { kind: 'formula', title: '', formula: 'E = mc2', columns: [['DA', 'dato']] },
    { kind: 'table', title: '   ', table: { header: ['a'], rows: [['1']] } },
    { kind: 'stats', title: null, items: [{ value: '1', label: 'uno' }] },
  ] }, 10)
  const formula = f.find(s => s.kind === 'formula')
  assert.ok(formula && formula.title.trim().length > 0, 'formula sin título')
  const tabla = f.find(s => s.kind === 'table')
  assert.ok(tabla && tabla.title.trim().length > 0, 'table sin título')
  const stats = f.find(s => s.kind === 'stats')
  assert.ok(stats && stats.title.trim().length > 0, 'stats sin título')
  const c = g.normalizeSlides({ slides: [
    { kind: 'cover', title: 'P', subtitle: '' },
    { kind: 'content', title: '', bullets: ['**Huella** de carbono enorme'] },
  ] }, 10)
  const cont = c.find(s => s.kind === 'content')
  assert.ok(cont && cont.title.includes('Huella'), 'content no hereda el bullet: ' + (cont && cont.title))
  assert.ok(!cont.title.includes('**'), 'fallback conserva **')
  const cov = g.normalizeSlides({ slides: [{ kind: 'cover', title: ' ', subtitle: 'sub' }] }, 5)
  assert.equal(cov[0].title, 'Presentación')
})

test('normalizeSlides valida kind, añade cover/closing, trunca bullets', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [
    { kind: 'stats', title: 'KPIs', items: [{ value: '1', label: 'a' }] },
    { kind: 'unknown', title: 'X', bullets: ['uno dos tres cuatro cinco seis siete ocho nueve diez once doce trece'] },
    { title: 'Sin kind', bullets: ['b'] }
  ] }, 5)
  assert.ok(slides.length >= 2)
  assert.equal(slides[0].kind, 'cover')
  assert.ok(slides.some(s => s.kind === 'closing'))
  const statsSlide = slides.find(s => s.kind === 'stats')
  assert.ok(statsSlide)
  const contentSlide = slides.find(s => s.kind === 'content' && s.title === 'X')
  assert.ok(contentSlide)
  assert.ok(contentSlide.bullets[0].split(/\s+/).length <= 12)
  assert.ok(contentSlide.bullets[0].endsWith('…') || contentSlide.bullets[0].length < 20)
})

test('normalizeSlides sanea chart: type, valores rotos y límites', () => {
  const g = load()
  const out = g.normalizeSlides({ slides: [
    { kind: 'cover', title: 'P' },
    { kind: 'chart', title: 'Producción', chart: {
      type: 'column',
      categories: ['2022', '2023', 'mala'],
      series: [{ name: 'Eólica', values: [245, 'x', null] }]
    } },
    { kind: 'chart', title: 'Mix', chart: {
      type: 'pie',
      categories: Array(15).fill('c').map((c, i) => c + i),
      series: [
        { name: 's1', values: [1, 2, 3] },
        { name: 's2', values: [4, 5, 6] },
        { name: 's3', values: [7, 8, 9] },
        { name: 's4', values: [10, 11, 12] },
        { name: 's5', values: [13, 14, 15] }
      ]
    } },
    { kind: 'closing', title: 'F', closing: ['i'] }
  ] }, 10)
  const col = out.find(s => s.title === 'Producción')
  assert.ok(col && col.kind === 'chart', 'chart no conservado')
  assert.equal(col.chart.type, 'bar', 'type column debe mapearse a bar')
  assert.deepEqual(col.chart.categories, ['2022', '2023', 'mala'], 'categories alteradas')
  assert.deepEqual(col.chart.series[0].values, [245, 0, 0], 'valores rotos no saneados a 0')
  const pie = out.find(s => s.title === 'Mix')
  assert.equal(pie.chart.type, 'pie', 'type pie perdido')
  assert.equal(pie.chart.series.length, 1, 'pie debe quedarse con 1 serie')
  assert.equal(pie.chart.categories.length, 3, 'categories no alineadas a la longitud de la serie')
  assert.ok(pie.chart.series[0].values.length === 3, 'values sin alinear')
})

test('normalizeSlides sanea axisTitle del chart', () => {
  const g = load()
  const out = g.normalizeSlides({ slides: [
    { kind: 'cover', title: 'P' },
    { kind: 'chart', title: 'G', chart: { type: 'bar', axisTitle: '  Consumo anual (MWh)  ', categories: ['a'], series: [{ name: 's', values: [1] }] } },
    { kind: 'chart', title: 'Numerico', chart: { type: 'bar', axisTitle: 12345, categories: ['a'], series: [{ name: 's', values: [1] }] } },
    { kind: 'chart', title: 'Larga', chart: { type: 'bar', axisTitle: 'x'.repeat(100), categories: ['a'], series: [{ name: 's', values: [1] }] } },
    { kind: 'closing', title: 'F', closing: ['i'] }
  ] }, 10)
  assert.equal(out.find(s => s.title === 'G').chart.axisTitle, 'Consumo anual (MWh)', 'axisTitle sin limpiar espacios')
  assert.equal(out.find(s => s.title === 'Numerico').chart.axisTitle, '12345', 'axisTitle numerico no convertido a texto')
  assert.ok(out.find(s => s.title === 'Larga').chart.axisTitle.length <= 48, 'axisTitle sin tope de longitud')
})

test('normalizeSlides cae a content si el chart no tiene datos renderizables', () => {
  const g = load()
  const out = g.normalizeSlides({ slides: [
    { kind: 'cover', title: 'P' },
    { kind: 'chart', title: 'Rotto', chart: { type: 'bar', categories: [], series: [] }, bullets: ['Bullet válido'] },
    { kind: 'chart', title: 'Vacio', chart: null },
    { kind: 'closing', title: 'F', closing: ['i'] }
  ] }, 10)
  const roto = out.find(s => s.title === 'Rotto')
  assert.ok(roto, 'slide chart-perdido descartado en vez de reconvertido')
  assert.equal(roto.kind, 'content', 'chart sin datos debe reconvertirse a content')
  assert.deepEqual(roto.bullets, ['Bullet válido'], 'bullets del slide reconvertido perdidos')
  assert.ok(!out.some(s => s.title === 'Vacio'), 'chart vacío sin bullets debe descartarse')
})

test('SAMPLE_DECK existe con 10 kinds (una por kind, incl. figure)', () => {
  const g = load()
  assert.ok(Array.isArray(g.SAMPLE_DECK))
  assert.equal(g.SAMPLE_DECK.length, 10)
  const kinds = g.SAMPLE_DECK.map(s => s.kind).sort()
  assert.equal(JSON.stringify(kinds), JSON.stringify(['chart', 'closing', 'compare', 'content', 'cover', 'figure', 'formula', 'process', 'stats', 'table']))
  g.SAMPLE_DECK.forEach(s => {
    assert.ok(s.kind)
    assert.ok(s.title)
    if (s.bullets) assert.ok(Array.isArray(s.bullets))
  })
})

test('buildPrompt con imágenes: sección figure en modo visión y ciego', () => {
  const g = load()
  const vis = g.buildPrompt('apuntes', 10, '', { imageCount: 3, vision: true })
  assert.ok(vis.includes('3 imágenes'), 'cuenta las imágenes')
  assert.ok(vis.includes('ADJUNTAS'), 'dice que van adjuntas')
  assert.ok(vis.includes('[IMAGEN n]'), 'explica los marcadores')
  assert.ok(vis.includes('kind:"figure"'), 'pide figure')
  const ciego = g.buildPrompt('apuntes', 10, '', { imageCount: 3, vision: false })
  assert.ok(ciego.includes('no puedes verlas'), 'modo ciego lo dice')
  assert.ok(ciego.includes('kind:"figure"'), 'aun asi pide figure')
  const sin = g.buildPrompt('apuntes', 10, '')
  assert.ok(!sin.includes('kind:"figure"'), 'sin imágenes no pide figure')
})

test('normalizeSlides valida el kind figure y limpia marcadores', () => {
  const g = load()
  const slides = g.normalizeSlides({ slides: [
    { kind: 'cover', title: 'Portada' },
    { kind: 'figure', title: 'Gráfico de barras [IMAGEN 2]', figure: '2', takeaway: 'Sube el pico [IMAGEN 1]' },
    { kind: 'closing', title: 'Cierre', closing: ['a', 'b', 'c'] }
  ] }, 10)
  const fig = slides.find(s => s.kind === 'figure')
  assert.ok(fig, 'figure conservado')
  assert.equal(fig.figure, 2, 'índice numérico')
  assert.ok(!/IMAGEN/.test(fig.title), 'marcador fuera del título: ' + fig.title)
  assert.ok(!/IMAGEN/.test(fig.takeaway), 'marcador fuera del takeaway')
})

test('assignFigures modo visión respeta el índice y descarta lo no usado', () => {
  const g = load()
  const imgs = [{ data: 'd1', w: 10, h: 10 }, { data: 'd2', w: 10, h: 10 }, { data: 'd3', w: 10, h: 10 }]
  const slides = [
    { kind: 'cover', title: 'P' },
    { kind: 'figure', title: 'FigA', figure: 2 },
    { kind: 'figure', title: 'FigB', figure: 1 },
    { kind: 'closing', title: 'C', closing: ['a'] }
  ]
  const out = g.assignFigures(slides, imgs, 'vision')
  const figs = out.filter(s => s.kind === 'figure')
  assert.equal(figs.length, 2, 'solo las figure pedidas')
  assert.equal(figs[0].img.data, 'd2', 'figure 2 → imagen 2')
  assert.equal(figs[1].img.data, 'd1', 'figure 1 → imagen 1')
  // la imagen 3 queda descartada a propósito (no se añaden anexos en visión)
  assert.ok(!out.some(s => s.title && /Imagen 3/.test(s.title)), 'sin anexo en visión')
})

test('assignFigures modo ciego coloca en orden y anexa sobrantes antes del cierre', () => {
  const g = load()
  const imgs = [{ data: 'd1', w: 10, h: 10 }, { data: 'd2', w: 10, h: 10 }]
  const slides = [
    { kind: 'cover', title: 'P' },
    { kind: 'figure', title: 'Fig1' },
    { kind: 'content', title: 'Texto', bullets: ['b'] },
    { kind: 'closing', title: 'C', closing: ['a'] }
  ]
  const out = g.assignFigures(slides, imgs, 'blind')
  const figs = out.filter(s => s.kind === 'figure')
  assert.equal(figs.length, 2, '1 pedida + 1 anexo')
  assert.equal(figs[0].img.data, 'd1', 'secuencial')
  const ci = out.findIndex(s => s.kind === 'closing')
  assert.ok(out[ci - 1].kind === 'figure', 'el anexo va justo antes del cierre')
  assert.ok(/Imagen 2/.test(out[ci - 1].title), 'anexo titulado con su nº')
})

test('assignFigures sin imágenes convierte figure en content placeholder', () => {
  const g = load()
  const out = g.assignFigures([{ kind: 'figure', title: 'Fig', takeaway: 'pie' }], [], 'vision')
  assert.equal(out[0].kind, 'content', 'figure sin imagen → content')
  assert.ok(/no disponible/.test(out[0].bullets[0]), 'mensaje placeholder')
})
