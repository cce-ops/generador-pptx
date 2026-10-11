import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';

const src = fs.readFileSync(new URL('../app.js', import.meta.url), 'utf8');

function load() {
  const sandbox = { console };
  vm.createContext(sandbox);
  vm.runInContext(src, sandbox, { filename: 'app.js' });
  return sandbox.__engine;
}

const plain = (v) => structuredClone(v);

test('THEMES tiene 6 estilos incluyendo verdeSkill', () => {
  const g = load();
  const themes = plain(g.THEMES);
  assert.equal(Object.keys(themes).length, 6);
  assert.ok('verdeSkill' in themes);
  for (const [k, t] of Object.entries(themes)) {
    assert.ok(t.name);
    assert.ok(t.bg);
    assert.ok(t.ink);
    assert.ok(t.accent);
    assert.ok(t.fontTitle);
    assert.ok(t.fontBody);
  }
});

test('defineMasters crea 3 masters con fondos correctos', () => {
  const g = load();
  const PptxGenJS = { shapes: {} };
  const pres = { defineSlideMaster: (m) => { pres._masters = pres._masters || []; pres._masters.push(m); } };
  const theme = g.THEMES.seriph;
  const masters = g.defineMasters(pres, theme);
  assert.ok(pres._masters);
  assert.equal(pres._masters.length, 3);
  const names = pres._masters.map(m => m.title);
  assert.deepEqual(names.sort(), ['CIERRE', 'CONTENIDO', 'PORTADA']);
  const port = pres._masters.find(m => m.title === 'PORTADA');
  const cont = pres._masters.find(m => m.title === 'CONTENIDO');
  const cier = pres._masters.find(m => m.title === 'CIERRE');
  assert.equal(port.background.color, theme.coverBg);
  assert.equal(cont.background.color, theme.bg);
  assert.equal(cier.background.color, theme.coverBg);
  assert.ok(masters.PORTADA === 'PORTADA' && masters.CONTENIDO === 'CONTENIDO' && masters.CIERRE === 'CIERRE');
});

test('tx helper existe y escribe texto con theme', () => {
  const g = load();
  const theme = g.THEMES.seriph;
  const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse' };
  const slide = { addText: (t, o) => { slide._last = { text: t, opts: o }; } };
  g.tx(slide, 'Hola', 1, 2, 3, 4, { fontSize: 18 }, theme);
  assert.equal(slide._last.text, 'Hola');
  assert.equal(slide._last.opts.fontFace, theme.fontBody);
  assert.equal(slide._last.opts.color, theme.ink);
});

test('card helper crea roundedRect con fill del theme', () => {
  const g = load();
  const theme = g.THEMES.seriph;
  const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse' };
  const slide = { addShape: (type, o) => { slide._lastShape = { type, opts: o }; } };
  g.card(slide, 1, 2, 3, 4, theme.cardBg, 'Test', theme, SH);
  assert.equal(slide._lastShape.type, SH.ROUNDED_RECTANGLE);
  assert.equal(slide._lastShape.opts.fill.color, theme.cardBg);
  assert.equal(slide._lastShape.opts.rectRadius, 0.12);
});

test('badge helper crea oval con fill badgeFill', () => {
  const g = load();
  const theme = g.THEMES.seriph;
  const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse' };
  const slide = { 
    addShape: (type, o) => { slide._lastShape = { type, opts: o }; },
    addText: (t, o) => { slide._lastText = { text: t, opts: o }; }
  };
  g.badge(slide, 3, 1, 2, 0.5, theme, SH);
  assert.equal(slide._lastShape.type, SH.OVAL);
  assert.equal(slide._lastShape.opts.fill.color, theme.badgeFill);
});

test('stat helper escribe número y etiqueta con colores', () => {
  const g = load();
  const theme = g.THEMES.seriph;
  const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse' };
  const slide = { addText: (t, o) => { slide._texts = slide._texts || []; slide._texts.push({ text: t, opts: o }); } };
  g.stat(slide, 1, 2, 3, '99%', 'label', theme.accent, theme, SH);
  assert.equal(slide._texts.length, 2);
  assert.equal(slide._texts[0].text, '99%');
  assert.equal(slide._texts[0].opts.color, theme.accent);
});

test('takeaway helper escribe con acento', () => {
  const g = load();
  const theme = g.THEMES.seriph;
  const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse' };
  const slide = { addText: (t, o) => { slide._texts = slide._texts || []; slide._texts.push({ text: t, opts: o }); } };
  g.takeaway(slide, 'Idea clave de prueba', 4.6, theme, SH);
  assert.ok(slide._texts.length >= 1);
  const txt = slide._texts[slide._texts.length - 1];
  assert.ok(Array.isArray(txt.text));
  assert.equal(txt.text[0].options.color, theme.accent2);
});

test('coverMotif añade dos óvalos translúcidos', () => {
  const g = load();
  const theme = g.THEMES.seriph;
  const SH = { ROUNDED_RECTANGLE: 'roundRect', OVAL: 'ellipse' };
  const slide = { addShape: (type, o) => { slide._shapes = slide._shapes || []; slide._shapes.push({ type, opts: o }); } };
  g.coverMotif(slide, theme, SH);
  assert.equal(slide._shapes.length, 2);
  assert.equal(slide._shapes[0].type, SH.OVAL);
  assert.ok('transparency' in slide._shapes[0].opts.fill);
});
