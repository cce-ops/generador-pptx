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
    }
    defineSlideMaster() {}
    addSlide(opts) { 
      console.log('addSlide called with:', opts)
      const s = { 
        addText: (t, o) => { console.log('addText:', t, o); s._texts = s._texts || []; s._texts.push({ text: t, opts: o }); },
        addShape: (type, o) => { console.log('addShape:', type, o); s._shapes = s._shapes || []; s._shapes.push({ type, opts: o }); },
        addTable: (data, o) => { console.log('addTable:', data, o); s._tables = s._tables || []; s._tables.push({ data, opts: o }); },
        _texts: [], _shapes: [], _tables: []
      }
      s.pres = this
      this._slides.push(s)
      console.log('addSlide returning slide:', !!s)
      return s
    }
    write() { return Promise.resolve(new Blob(['test'], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })) }
  }
  const sandbox = { console, PptxGenJS: MockPptx };
  vm.createContext(sandbox);
  vm.runInContext(src, sandbox, { filename: 'app.js' });
  return sandbox.__engine;
}

const makeSlide = () => ({})

test('buildPptx basic', async () => {
  const g = load()
  const slides = [{ kind: 'cover', title: 'Test' }, { kind: 'content', title: 'Hola', bullets: ['a'] }]
  const blob = await g.buildPptx(slides, { title: 'Test', theme: 'seriph' })
  assert.ok(blob instanceof Blob)
})
