import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';

const src = fs.readFileSync(new URL('../app.js', import.meta.url), 'utf8');

function load() {
  const MockPptx = class {
    constructor() {
      this.layout = ''
      this.author = ''
      this.title = ''
      this._slides = []
      this.shapes = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse', RECTANGLE: 'rect' }
      if (this.constructor.instances) this.constructor.instances.push(this)
    }
    defineSlideMaster() {}
    addSlide(opts) {
      const s = makeSlide()
      s.pres = this
      this._slides.push(s)
      return s
    }
    write() { return Promise.resolve(new Blob(['test'], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })) }
  }
  const sandbox = { console, PptxGenJS: MockPptx };
  vm.createContext(sandbox);
  vm.runInContext(src, sandbox, { filename: 'app.js' });
  MockPptx.instances = [];
  sandbox.__engine.__mockPptx = MockPptx;
  return sandbox.__engine;
}

const plain = (v) => structuredClone(v);

function makeSlide() {
  return {
    _texts: [],
    _shapes: [],
    _tables: [],
    _charts: [],
    _images: [],
    addText(t, o) {
      if (Array.isArray(t)) {
        this._texts.push({ text: '', opts: o });
        t.forEach(item => this._texts.push({ text: item.text, opts: item.options }));
      } else {
        this._texts.push({ text: t, opts: o });
      }
    },
    addShape(type, o) { this._shapes.push({ type, opts: o }); },
    addTable(data, o) { this._tables.push({ data, opts: o }); },
    addChart(type, data, o) { this._charts.push({ type, data, opts: o }); },
    addImage(o) { this._images.push(o); }
  };
}

const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse', RECTANGLE: 'rect' };
const theme = {
  bg: 'FFFFFF', ink: '1A1A1A', muted: '555555', accent: '4D83C3',
  fontTitle: 'Georgia', fontBody: 'Georgia',
  cardBg: 'F5F5F5', cardTitle: '1A1A1A', cardText: '1A1A1A',
  tableHeaderBg: '1A1A1A', tableHeaderFg: 'FFFFFF',
  footer: '888888', badgeFill: '4D83C3', badgeText: 'FFFFFF',
  coverBg: '4D83C3', coverTitle: 'FFFFFF', coverSub: 'FFFFFF',
  accent2: '9BC53D', accent3: '1C7293', accent4: 'E0A526',
  accent5: 'C8553D', accent6: '6B7F76',
  redTint: 'F6E3DE', grid: 'C9D8CF'
};

test('render_cover añade título, subtítulo y motivos', () => {
  const g = load();
  const slide = makeSlide();
  g.render_cover(slide, { title: 'Título portada', subtitle: 'Subtítulo' }, theme);
  assert.ok(slide._texts.some(t => t.text === 'Título portada'));
  assert.ok(slide._texts.some(t => t.text === 'Subtítulo'));
  assert.ok(slide._shapes.some(s => s.type === 'ellipse'));
});

test('autoFont reduce hasta que el texto cabe en la caja (sin normAutofit)', () => {
  const g = load();
  const t = '4. Dimensionamiento del Sistema Solar Fotovoltaico';
  const sz = g.autoFont(t, 9, 0.8, 28, 18);
  assert.ok(sz <= 28 && sz >= 18, 'sz fuera de rango: ' + sz);
  const cpl = Math.max(1, (9 * 72) / (sz * 0.60));
  const lines = Math.max(1, Math.ceil(t.length / cpl));
  assert.ok(lines * sz * 1.25 <= 0.8 * 72 + 1, 'no cabe: ' + lines + 'x' + sz);
  assert.equal(g.autoFont('KPIs', 9, 0.8, 28, 18), 28, 'texto corto debe quedar en el máximo');
});

test('titulos y values largos bajan de tamaño para no solaparse', () => {
  const g = load();
  // contentTitle largo
  let slide = makeSlide();
  g.render_content(slide, { title: '4. Dimensionamiento del Sistema Solar Fotovoltaico', bullets: ['b'] }, theme);
  const t1 = slide._texts.find(t => t.text && t.text.startsWith('4. Dimens'));
  assert.ok(t1 && t1.opts.fontSize < 28, 'titulo content sin reducir: ' + (t1 && t1.opts.fontSize));
  // cover largo
  slide = makeSlide();
  g.render_cover(slide, { title: 'Diseño de Comunidad Energética Renovable en Rubí', subtitle: '' }, theme);
  const t2 = slide._texts.find(t => t.text && t.text.startsWith('Diseño de Comunidad'));
  assert.ok(t2 && t2.opts.fontSize < 38, 'titulo cover sin reducir: ' + (t2 && t2.opts.fontSize));
  // compare: value largo reduce, corto se queda (3 tarjetas = cajas de 2.4" como en los decks reales)
  slide = makeSlide();
  g.render_compare(slide, { title: 'C', items: [
    { value: '1.291 € en CER frente a 5.000 € en autoconsumo individual.', label: 'Coste' },
    { value: '18 usuarios', label: 'Users' },
    { value: '9,9 años', label: 'Retorno' },
  ] }, theme);
  const t3 = slide._texts.find(t => t.text && t.text.startsWith('1.291'));
  const t4 = slide._texts.find(t => t.text === '18 usuarios');
  assert.ok(t3 && t3.opts.fontSize < 24, 'value compare sin reducir: ' + (t3 && t3.opts.fontSize));
  assert.ok(t4 && t4.opts.fontSize === 24, 'value corto cambiado: ' + (t4 && t4.opts.fontSize));
  // stats: value largo reduce, corto se queda
  slide = makeSlide();
  g.render_stats(slide, { title: 'K', items: [
    { value: '60 % de la huella de carbono total', label: 'a' },
    { value: '1,7', label: 'b' },
  ] }, theme);
  const t5 = slide._texts.find(t => t.text && t.text.startsWith('60 %'));
  const t6 = slide._texts.find(t => t.text === '1,7');
  assert.ok(t5 && t5.opts.fontSize < 36, 'value stats sin reducir: ' + (t5 && t5.opts.fontSize));
  assert.ok(t6 && t6.opts.fontSize === 36, 'value corto stats cambiado: ' + (t6 && t6.opts.fontSize));
});

test('render_stats añade stats y takeaway', () => {
  const g = load();
  const slide = makeSlide();
  g.render_stats(slide, { title: 'KPIs', items: [{ value: '1,7', label: 'planetas', color: 'accent5' }, { value: '>60%', label: 'carbono', color: 'accent1' }], takeaway: 'Vivimos por encima' }, theme);
  assert.ok(slide._texts.some(t => t.text === 'KPIs'));
  assert.ok(slide._texts.some(t => t.text && t.text.includes('1,7')));
  assert.ok(slide._texts.some(t => typeof t.text === 'string' && t.text.includes('Vivimos')));
  const cards = slide._shapes.filter(s => s.type === 'roundRect');
  assert.ok(cards.some(c => c.opts.y + c.opts.h >= 4.3), 'cards de stats no llenan la diapo hacia abajo');
});

test('render_process añade badges y pasos', () => {
  const g = load();
  const slide = makeSlide();
  g.render_process(slide, { title: 'Proceso ACV', steps: [['Fase 1', 'Detalle 1'], ['Fase 2', 'Detalle 2']], takeaway: '4 fases' }, theme);
  assert.ok(slide._texts.some(t => t.text === 'Proceso ACV'));
  assert.ok(slide._shapes.some(s => s.type === 'ellipse'));
  const fase = slide._texts.find(t => t.text === 'Fase 1');
  assert.ok(fase && fase.opts.fit === 'shrink', 'paso del proceso sin auto-ajuste');
  assert.ok(slide._texts.some(t => t.text && t.text.includes('Fase 1')));
});

test('render_stats sanitiza colores invalidos al acento del tema', () => {
  const g = load();
  const slide = makeSlide();
  g.render_stats(slide, { title: 'K', items: [{ value: '9', label: 'x', color: 'green' }] }, theme);
  const num = slide._texts.find(t => t.text === '9');
  assert.ok(num, 'falta el numero');
  assert.equal(num.opts.color, theme.accent);
});

test('render_table crea tabla con cabecera oscura y takeaway', () => {
  const g = load();
  const slide = makeSlide();
  g.render_table(slide, { title: 'Comparativa', headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']], formula: 'LCSA = LCA + LCC + S-LCA' }, theme);
  assert.ok(slide._tables.length === 1);
  const tbl = slide._tables[0];
  assert.ok(tbl.data[0][0].options.fill.color === theme.tableHeaderBg);
  assert.ok(tbl.data[0][0].options.color === theme.tableHeaderFg);
  assert.ok(slide._texts.some(t => typeof t.text === 'string' && t.text.includes('LCSA')));
});

test('render_formula añade banner y 3 columnas', () => {
  const g = load();
  const slide = makeSlide();
  g.render_formula(slide, { title: 'Fórmula', formula: 'E = A × F', columns: [['Alcance 1', 'Directas'], ['Alcance 2', 'Energía'], ['Alcance 3', 'Cadena']] }, theme);
  assert.ok(slide._texts.some(t => t.text === 'Fórmula'));
  assert.ok(slide._texts.some(t => t.text && t.text.includes('E = A × F')));
  assert.ok(slide._shapes.some(s => s.type === 'roundRect'));
});

test('render_compare añade 3 cards', () => {
  const g = load();
  const slide = makeSlide();
  g.render_compare(slide, { title: 'Lineal vs Circular', items: [{ value: 'Tomar→Hacer→Tirar', label: 'Lineal' }, { value: 'Diseñar sin residuos', label: 'Circular 1' }, { value: 'Mantener en uso', label: 'Circular 2' }] }, theme);
  assert.ok(slide._shapes.filter(s => s.type === 'roundRect').length >= 3);
  assert.ok(slide._texts.some(t => t.text && t.text.includes('Lineal')));
  assert.equal(slide._texts.filter(t => t.text === 'Lineal').length, 1, 'label del compare duplicado');
  const valor = slide._texts.find(t => t.text === 'Tomar→Hacer→Tirar');
  assert.ok(valor && valor.opts.fontSize <= 26, 'valor de la card demasiado grande: ' + (valor && valor.opts.fontSize));
  const cards2 = slide._shapes.filter(s => s.type === 'roundRect');
  assert.ok(cards2.some(c => c.opts.y + c.opts.h >= 4.3), 'cards del compare no llenan la diapo hacia abajo');
});

test('render_content bullets con espacio', () => {
  const g = load();
  const slide = makeSlide();
  g.render_content(slide, { title: 'Contenido', bullets: ['Bullet corto', 'Otro bullet aquí'] }, theme);
  assert.ok(slide._texts.some(t => t.text === 'Contenido'));
  const bullets = slide._texts.filter(t => t.opts && t.opts.bullet);
  assert.ok(bullets.length >= 2);
  assert.ok(bullets[0].opts.paraSpaceAfter === 12);
  assert.ok(bullets[0].opts.lineSpacingMultiple === 1.25);
  assert.equal(bullets[0].opts.fontSize, 18);
  const caja = slide._texts.find(t => t.text === '');
  assert.ok(caja && caja.opts.valign === 'top', 'bullets sin valign top');
  assert.ok(caja.opts.y + caja.opts.h <= 5.1, 'caja de bullets sale por abajo');
  assert.equal(caja.opts.fit, 'shrink', 'bullets sin auto-ajuste');
});

test('paginateSlides parte bullets en paginas de 4 (presupuesto de altura)', () => {
  const g = load();
  const pages = g.paginateSlides([
    { kind: 'content', title: 'X', bullets: Array(9).fill('b') },
  ]);
  const contents = pages.filter(p => !p.cover && p.bullets);
  assert.equal(contents.length, 3, 'esperaba 3 paginas de 4: ' + contents.length);
  assert.ok(contents.every(p => p.bullets.length <= 4), 'pagina con mas de 4 bullets');
});

test('render_closing fondo coverBg y 3 badges', () => {
  const g = load();
  const slide = makeSlide();
  slide.background = { color: 'FFFFFF' };
  g.render_closing(slide, { title: 'Ideas clave', items: ['Idea 1', 'Idea 2', 'Idea 3'] }, theme);
  assert.equal(slide.background.color, 'FFFFFF', 'closing pisa el fondo');
  assert.ok(slide._shapes.filter(s => s.type === 'ellipse').length >= 3);
  assert.ok(slide._texts.some(t => t.text === 'Ideas clave'));
  const cuerpo = slide._texts.find(t => t.text === 'Idea 1');
  assert.ok(cuerpo && cuerpo.opts.color === theme.ink, 'texto del cierre en color de tema claro');
  assert.equal(cuerpo.opts.fit, 'shrink', 'items del cierre sin auto-ajuste');
});

test('render_cover titulo a la izquierda en bloque estrecho', () => {
  const g = load();
  const slide = makeSlide();
  g.render_cover(slide, { title: 'T', subtitle: 'S' }, theme);
  const title = slide._texts.find(t => t.text === 'T');
  assert.ok(title);
  assert.equal(title.opts.x, 0.7);
  assert.ok(title.opts.w <= 6.5, 'portada a ancho completo: ' + title.opts.w);
  assert.equal(title.opts.align, 'left');
});

test('ningun renderer usa placeholders huerfanos y el titulo va con tema', () => {
  const g = load();
  const cases = [
    ['render_cover', { title: 'T', subtitle: 'S' }],
    ['render_stats', { title: 'T', items: [{ value: '1', label: 'a' }] }],
    ['render_process', { title: 'T', steps: [['A', 'b']] }],
    ['render_table', { title: 'T', headers: ['H'], rows: [['c']] }],
    ['render_formula', { title: 'T', formula: 'E=A', columns: [['C', 'd']] }],
    ['render_compare', { title: 'T', items: [{ value: 'v', label: 'l' }] }],
    ['render_content', { title: 'T', bullets: ['b'] }],
    ['render_closing', { title: 'T', items: ['i'] }],
    ['render_chart', { title: 'T', chart: { type: 'bar', categories: ['a'], series: [{ name: 's', values: [1] }] } }],
  ];
  for (const [fn, data] of cases) {
    const slide = makeSlide();
    g[fn](slide, data, theme);
    const texts = slide._texts;
    assert.ok(!texts.some(t => t.opts && t.opts.placeholder), fn + ' usa placeholder huerfano');
    const title = texts.find(t => t.text === 'T');
    assert.ok(title, fn + ' no pinta el titulo');
    assert.equal(title.opts.x, fn === 'render_cover' ? 0.7 : 0.5, fn + ' titulo en x erronea');
    assert.ok(title.opts.fontSize >= 28, fn + ' titulo pequeno');
    assert.equal(title.opts.fit, 'shrink', fn + ' titulo sin auto-ajuste');
    if (fn === 'render_cover') {
      assert.ok(title.opts.w <= 6.5, fn + ' portada a ancho completo');
      assert.equal(title.opts.align, 'left', fn + ' portada no alineada');
    }
  }
});

test('paginateSlides expande bullets y tablas', () => {
  const g = load();
  const slides = [
    { kind: 'cover', title: 'Portada', subtitle: 'Sub' },
    { kind: 'content', title: 'Muchos bullets', bullets: Array(15).fill('b') },
    { kind: 'table', title: 'Tabla grande', headers: ['H'], rows: Array(12).fill(['r']) }
  ];
  const pages = g.paginateSlides(slides);
  assert.ok(pages.length > 3);
  assert.ok(pages.some(p => p.title && p.title.includes('(cont.)')));
});

test('paginateSlides conserva takeaways, formula de tabla y chart (bug: se perdian)', () => {
  const g = load();
  const pages = g.paginateSlides([
    { kind: 'stats', title: 'S', items: [{ value: '1', label: 'a' }], takeaway: 'TA stats' },
    { kind: 'process', title: 'P', steps: [['A', 'b']], takeaway: 'TA proceso' },
    { kind: 'table', title: 'T', table: { header: ['H'], rows: [['c']] }, formula: 'LCSA = LCA' },
    { kind: 'chart', title: 'G', chart: { type: 'bar', categories: ['a'], series: [{ name: 's', values: [1] }] }, takeaway: 'TA chart' },
  ]);
  assert.equal(pages.find(p => p.stats).takeaway, 'TA stats', 'takeaway de stats perdido en paginate');
  assert.equal(pages.find(p => p.process).takeaway, 'TA proceso', 'takeaway de process perdido en paginate');
  assert.equal(pages.find(p => p.table).formula, 'LCSA = LCA', 'formula-banner de tabla perdida en paginate');
  const ch = pages.find(p => p.chart);
  assert.ok(ch && ch.chart, 'pagina chart sin emitir');
  assert.equal(ch.takeaway, 'TA chart', 'takeaway de chart perdido en paginate');
});

test('render_chart crea grafico nativo tematizado y dentro del lienzo', () => {
  const g = load();
  const slide = makeSlide();
  g.render_chart(slide, { title: 'Producción renovable', chart: {
    type: 'bar',
    categories: ['2022', '2023'],
    series: [{ name: 'Eólica', values: [245, 260] }, { name: 'Solar', values: [180, 200] }]
  } }, theme);
  assert.equal(slide._charts.length, 1, 'sin addChart');
  const c = slide._charts[0];
  assert.equal(c.type, 'bar', 'tipo de grafico: ' + c.type);
  assert.equal(c.data.length, 2, 'series perdidas');
  assert.deepEqual(c.data[0].labels, ['2022', '2023'], 'categories perdidas');
  assert.deepEqual(c.data[0].values, [245, 260], 'values perdidos');
  assert.equal(c.data[1].name, 'Solar', 'nombre de serie perdido');
  const o = c.opts;
  assert.ok(o.x >= 0 && o.y >= 1.3, 'chart muy arriba: y=' + o.y);
  assert.ok(o.x + o.w <= 10.01 && o.y + o.h <= 5.625, 'chart fuera del lienzo: ' + (o.x + o.w) + 'x' + (o.y + o.h));
  assert.ok(o.y + o.h >= 4.3, 'chart no llena la diapo hacia abajo: ' + (o.y + o.h));
  assert.ok(o.chartColors.includes(theme.accent), 'chart sin color del tema');
  assert.equal(o.showLegend, true, 'multi-serie sin leyenda');
  assert.ok(slide._texts.some(t => t.text === 'Producción renovable'), 'titulo del chart sin pintar');
});

test('render_chart pinta el titulo del eje de valores (magnitud y unidades)', () => {
  const g = load();
  const slide = makeSlide();
  g.render_chart(slide, { title: 'Consumo', chart: {
    type: 'bar',
    axisTitle: 'Consumo anual (MWh)',
    categories: ['Escuela', 'Bar'],
    series: [{ name: 'Demanda', values: [80.9, 74.3] }]
  } }, theme);
  const o = slide._charts[0].opts;
  assert.equal(o.showValAxisTitle, true, 'eje de valores sin titulo');
  assert.equal(o.valAxisTitle, 'Consumo anual (MWh)', 'titulo del eje perdido');
  // sin axisTitle no se muestra titulo de eje
  const s2 = makeSlide();
  g.render_chart(s2, { title: 'X', chart: { type: 'bar', categories: ['a'], series: [{ name: 's', values: [1] }] } }, theme);
  assert.ok(!s2._charts[0].opts.showValAxisTitle, 'showValAxisTitle activo sin axisTitle');
});

test('chartPalette genera n colores distintos aunque el tema sea monocromo', () => {
  const g = load();
  const mono = { ...theme, accent2: theme.accent, accent3: theme.accent, accent4: theme.accent };
  const p = g.chartPalette(mono, 6);
  assert.equal(p.length, 6, 'paleta corta: ' + p.length);
  assert.equal(new Set(p).size, 6, 'colores repetidos: ' + p.join(','));
  assert.equal(p[0], theme.accent, 'el primer color debe ser el acento del tema');
  p.forEach(c => assert.match(c, /^[0-9A-F]{6}$/, 'color no hex valido: ' + c));
  assert.equal(g.chartPalette(theme, 1).length, 1, 'n=1 debe dar 1 color');
});

test('render_chart colorea cada barra, serie y rebanada aunque el tema sea monocromo', () => {
  const g = load();
  const mono = { ...theme, accent2: theme.accent, accent3: theme.accent, accent4: theme.accent };
  // barra simple: color por punto → cada "cosa" su color
  let slide = makeSlide();
  g.render_chart(slide, { title: 'T', chart: {
    type: 'bar', categories: ['Escuela', 'Bar', 'Biblioteca', 'Casas'],
    series: [{ name: 'Consumo', values: [1, 2, 3, 4] }]
  } }, mono);
  const o1 = slide._charts[0].opts;
  assert.equal(o1.chartColors.length, 4, 'bar simple sin color por barra: ' + JSON.stringify(o1.chartColors));
  assert.equal(new Set(o1.chartColors).size, 4, 'barras con colores repetidos: ' + o1.chartColors.join(','));
  // multi-serie: color por serie
  slide = makeSlide();
  g.render_chart(slide, { title: 'T', chart: {
    type: 'bar', categories: ['a', 'b'],
    series: [{ name: 'E1', values: [1, 2] }, { name: 'E2', values: [3, 4] }, { name: 'E3', values: [5, 6] }]
  } }, mono);
  const o2 = slide._charts[0].opts;
  assert.equal(o2.chartColors.length, 3, 'multi-serie sin color por serie');
  assert.equal(new Set(o2.chartColors).size, 3, 'series con colores repetidos: ' + o2.chartColors.join(','));
  // pie: color por rebanada
  slide = makeSlide();
  g.render_chart(slide, { title: 'T', chart: {
    type: 'pie', categories: ['a', 'b', 'c', 'd', 'e', 'f'],
    series: [{ name: 's', values: [1, 2, 3, 4, 5, 6] }]
  } }, mono);
  const o3 = slide._charts[0].opts;
  assert.equal(o3.chartColors.length, 6, 'pie sin color por rebanada: ' + JSON.stringify(o3.chartColors));
  assert.equal(new Set(o3.chartColors).size, 6, 'rebanadas con colores repetidos: ' + o3.chartColors.join(','));
});

test('render_chart en pie usa axisTitle como titulo del grafico', () => {
  const g = load();
  const slide = makeSlide();
  g.render_chart(slide, { title: 'Mix', chart: {
    type: 'pie',
    axisTitle: 'Reparto del mix (%)',
    categories: ['A', 'B'],
    series: [{ name: 's', values: [60, 40] }]
  } }, theme);
  const o = slide._charts[0].opts;
  assert.equal(o.showTitle, true, 'pie sin titulo del grafico con axisTitle');
  assert.equal(o.title, 'Reparto del mix (%)', 'titulo del pie perdido');
  assert.ok(!o.showValAxisTitle, 'pie no debe mostrar eje de valores');
});

test('render_chart pie usa solo la 1a serie y reserva hueco al takeaway', () => {
  const g = load();
  const slide = makeSlide();
  g.render_chart(slide, { title: 'Mix', chart: {
    type: 'pie',
    categories: ['A', 'B'],
    series: [{ name: 's1', values: [60, 40] }, { name: 's2', values: [1, 2] }]
  }, takeaway: 'La eólica lidera' }, theme);
  const c = slide._charts[0];
  assert.equal(c.type, 'pie', 'tipo pie: ' + c.type);
  assert.equal(c.data.length, 1, 'pie debe quedarse con 1 serie');
  assert.ok(c.opts.y + c.opts.h <= 4.55, 'chart pisa el takeaway: ' + (c.opts.y + c.opts.h));
  assert.ok(c.opts.showLegend, 'pie sin leyenda');
  assert.ok(slide._texts.some(t => t.text && t.text.includes('La eólica lidera')), 'takeaway del chart sin pintar');
});

test('paginateSlides pasa compare/closing normalizados a las paginas', () => {
  const g = load();
  const pages = g.paginateSlides([
    { kind: 'compare', title: 'C', compare: [{ value: 'v', label: 'l' }] },
    { kind: 'closing', title: 'F', closing: ['idea'] },
  ]);
  const cmp = pages.find(p => !p.cover && p.title === 'C');
  const cls = pages.find(p => !p.cover && p.title === 'F');
  assert.ok(cmp && cmp.compare && cmp.compare.length === 1, 'compare sin llevar a la pagina');
  assert.ok(cls && cls.closing && cls.closing.length === 1, 'closing sin llevar a la pagina');
});

test('paginateSlides no emite paginas sin contenido', () => {
  const g = load();
  const pages = g.paginateSlides([
    { kind: 'content', title: 'Vacia', bullets: [] },
    { kind: 'table', title: 'SinFilas', table: { header: ['H'], rows: [] } },
    { kind: 'content', title: 'Llena', bullets: ['a'] },
  ]);
  assert.ok(!pages.some(p => !p.cover && p.title === 'Vacia'), 'pagina solo-titulo emitida');
  assert.ok(!pages.some(p => !p.cover && p.title === 'SinFilas'), 'pagina tabla-vacia emitida');
  assert.ok(pages.some(p => p.title === 'Llena'), 'pagina con contenido perdida');
});

test('render_closing acepta el campo closing (ademas de items)', () => {
  const g = load();
  const slide = makeSlide();
  g.render_closing(slide, { title: 'F', closing: ['unico'] }, theme);
  assert.ok(slide._texts.some(t => t.text === 'unico'), 'closing no pintado desde campo closing');
  const slide2 = makeSlide();
  g.render_closing(slide2, { title: 'F', closing: ['a', 'b', 'c', 'd'] }, theme);
  assert.equal(slide2._texts.filter(t => ['a', 'b', 'c', 'd'].includes(t.text)).length, 3, 'cierre sin cap de 3 items');
});

test('buildPptx usa LAYOUT_16x9 (10x5.625) y todo cabe en el lienzo', () => {
  const g = load();
  const deck = [
    { kind: 'cover', title: 'Portada', subtitle: 'Sub' },
    { kind: 'stats', title: 'S', items: [{ value: '1', label: 'a' }], takeaway: 't' },
    { kind: 'process', title: 'P', steps: [['A', 'b'], ['C', 'd']], takeaway: 't' },
    { kind: 'table', title: 'T', table: { header: ['H1', 'H2'], rows: [['a', 'b'], ['c', 'd']] } },
    { kind: 'formula', title: 'F', formula: 'E = mc2', columns: [['A', 'b'], ['C', 'd'], ['E', 'f']] },
    { kind: 'compare', title: 'C', items: [{ value: 'v', label: 'l' }, { value: 'v2', label: 'l2' }, { value: 'v3', label: 'l3' }] },
    { kind: 'content', title: 'X', bullets: ['b1', 'b2'] },
    { kind: 'chart', title: 'G', chart: { type: 'bar', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] } },
    { kind: 'closing', title: 'Z', closing: ['i1', 'i2', 'i3'] },
  ];
  return g.buildPptx(deck, { title: 'T', theme: 'seriph' }).then(() => {
    const inst = g.__mockPptx.instances.at(-1);
    assert.equal(inst.layout, 'LAYOUT_16x9', 'layout no aplicado: ' + inst.layout);
    // titulos numerados por apartado (portada sin numero)
    const t0 = inst._slides[0]._texts.map(t => t.text);
    assert.ok(t0.includes('Portada'), 'portada sin titulo propio');
    assert.ok(!t0.includes('1. Portada'), 'portada numerada');
    assert.ok(inst._slides[1]._texts.some(t => t.text === '1. S'), 'slide 2 sin numerar');
    assert.ok(inst._slides[3]._texts.some(t => t.text === '3. T'), 'slide 4 sin numerar');
    assert.ok(inst._slides.at(-1)._texts.some(t => t.text === '8. Z'), 'cierre sin numerar');
    // cierre con fondo normal (no coverBg)
    assert.equal(inst._slides.at(-1).background.color, 'FFFFFF', 'cierre con fondo especial');
    // el chart llego al slide correspondiente (seccion 7)
    assert.equal(inst._slides[7]._charts.length, 1, 'chart no renderizado en buildPptx');
    // el contenido llega hasta al menos y=4.3: sin dejarse la base vacia
    inst._slides.slice(1).forEach((s, i) => {
      const bottoms = [
        ...s._texts.filter(t => t.opts && t.opts.y !== undefined && !/^\d+\/\d+$/.test(String(t.text))).map(t => t.opts.y + t.opts.h),
        ...s._shapes.map(sh => sh.opts.y + sh.opts.h),
        ...s._tables.map(tb => tb.opts.y + tb.opts.rowH.reduce((a, c) => a + c, 0)),
        ...s._charts.map(c => c.opts.y + c.opts.h),
      ];
      const max = Math.max(...bottoms);
      assert.ok(max >= 4.3, 'slide ' + (i + 2) + ' termina muy arriba: ' + max.toFixed(2));
    });
    const W = 10, H = 5.625;
    inst._slides.forEach((s, i) => {
      const boxes = [
        ...s._texts.map(t => t.opts),
        ...s._shapes.map(sh => sh.opts),
        ...s._tables.map(tb => tb.opts),
        ...s._charts.map(c => c.opts),
      ].filter(b => b && b.x !== undefined);
      for (const b of boxes) {
        const h = b.h !== undefined ? b.h : (Array.isArray(b.rowH) ? b.rowH.reduce((a, c) => a + c, 0) : 0)
        assert.ok(b.x >= -0.01 && b.y >= -0.01, 'slide ' + (i + 1) + ' posicion negativa');
        assert.ok(b.x + b.w <= W + 0.01, 'slide ' + (i + 1) + ' desborda ancho: ' + (b.x + b.w));
        assert.ok(b.y + h <= H + 0.01, 'slide ' + (i + 1) + ' desborda alto: ' + (b.y + h));
      }
    });
  });
});

test('buildPptx devuelve Blob y usa masters', () => {
  const g = load();
  // Mock PptxGenJS constructor
  const origPptx = globalThis.PptxGenJS;
  const blobs = [];
  globalThis.PptxGenJS = class {
    constructor() { this.layout = ''; this.author = ''; this.title = ''; this._slides = []; }
    defineSlideMaster() {}
    addSlide() { const s = makeSlide(); this._slides.push(s); return s; }
    write(opts) { return Promise.resolve(new Blob(['test'], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })); }
  };
  const slides = [{ kind: 'cover', title: 'Test' }, { kind: 'content', title: 'Hola', bullets: ['a'] }];
  return g.buildPptx(slides, { title: 'Test', theme: 'seriph' }).then(blob => {
    globalThis.PptxGenJS = origPptx;
    assert.ok(blob instanceof Blob);
  });
});

test('render_figure inserta la imagen centrada respetando proporción', () => {
  const g = load();
  const slide = makeSlide();
  // horizontal (w>h): debe caber por ancho
  g.render_figure(slide, { title: 'Figura', img: { data: 'data:image/png;base64,AAA', w: 400, h: 200 } }, theme);
  assert.equal(slide._images.length, 1, 'una imagen');
  const im = slide._images[0];
  assert.equal(im.data, 'data:image/png;base64,AAA');
  // limitado por alto (3.55): sc = min(9/400, 3.55/200) = 3.55/200
  const sc = Math.min(9 / 400, 3.55 / 200);
  assert.ok(Math.abs(im.w - 400 * sc) < 0.01, 'ancho proporcional: ' + im.w);
  assert.ok(Math.abs(im.h - 200 * sc) < 0.01, 'alto proporcional: ' + im.h);
  assert.ok(im.x >= 0.5 && im.x <= 1.5, 'centrada en x: ' + im.x);
  assert.ok(slide._texts.some(t => t.text === 'Figura'), 'título');
});

test('render_figure sin imagen dibuja placeholder (no rompe)', () => {
  const g = load();
  const slide = makeSlide();
  g.render_figure(slide, { title: 'Figura', img: null }, theme);
  assert.equal(slide._images.length, 0, 'sin imagen no addImage');
  assert.ok(slide._texts.some(t => /no disponible/.test(String(t.text))), 'mensaje placeholder');
});

test('render_figure reserve hueco al takeaway', () => {
  const g = load();
  const slide = makeSlide();
  g.render_figure(slide, { title: 'F', img: { data: 'd', w: 900, h: 900 }, takeaway: 'pie' }, theme);
  const im = slide._images[0];
  // con takeaway: alto máx 3.0 (una imagen cuadrada se ajusta por alto)
  assert.ok(im.h <= 3.01, 'alto limitado por takeaway: ' + im.h);
  assert.ok(slide._texts.some(t => /pie/.test(String(t.text))), 'takeaway visible');
});

test('paginateSlides pasa figure con su imagen e índice', () => {
  const g = load();
  const img = { data: 'x', w: 100, h: 50 };
  const pages = g.paginateSlides([
    { kind: 'cover', title: 'P' },
    { kind: 'figure', title: 'Fig', figure: 1, img },
    { kind: 'closing', title: 'C', closing: ['a'] }
  ]);
  const fp = pages.find(p => p.figure !== undefined);
  assert.ok(fp, 'pagina figure');
  assert.equal(fp.figure, 1);
  assert.equal(fp.img.data, 'x', 'imagen propagada');
});
