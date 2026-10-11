import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import vm from 'node:vm'

const src = fs.readFileSync(new URL('../app.js', import.meta.url), 'utf8')

function gemini(text) {
  return {
    ok: true,
    json: () => Promise.resolve({ candidates: [{ content: { parts: [{ text }] } }] }),
  }
}
const deckJson = (slides) => JSON.stringify({ slides })

function load(fetchStub) {
  const sandbox = {
    console,
    document: { getElementById: () => null },
    fetch: fetchStub,
    AbortController,
    setTimeout,
    clearTimeout,
  }
  vm.createContext(sandbox)
  vm.runInContext(src, sandbox, { filename: 'app.js' })
  return sandbox.__generador
}

const content = (i) => ({ kind: 'content', title: 'T' + i, bullets: ['b ' + i] })
const cover = () => ({ kind: 'cover', title: 'P', subtitle: 's' })
const closing = () => ({ kind: 'closing', title: 'Cierre', closing: ['a', 'b', 'c'] })
const shortDeck = () => [cover(), content(1), content(2), closing()] // 4
function fullDeck(n) {
  const d = [cover()]
  for (let i = 1; i <= n - 2; i++) d.push(content(i))
  d.push(closing())
  return d
}

test('structureSlides reintenta si faltan diapositivas y para al completar', async () => {
  const prompts = []
  let call = 0
  const g = load((url, opts) => {
    call++
    prompts.push(JSON.parse(opts.body).contents[0].parts[0].text)
    return Promise.resolve(gemini(deckJson(call === 1 ? shortDeck() : fullDeck(6))))
  })
  const slides = await g.structureSlides('apuntes', { provider: 'aistudio', apiKey: 'K', maxSlides: 6, timeoutMs: 5000 })
  assert.equal(slides.length, 6, 'deck final: ' + slides.length)
  assert.equal(call, 2, 'esperaba 2 llamadas: ' + call)
  assert.ok(prompts[1].includes('EXACTAMENTE 6'), 'segunda llamada sin correctivo')
})

test('structureSlides no reintenta si ya entrega todas', async () => {
  let call = 0
  const g = load(() => { call++; return Promise.resolve(gemini(deckJson(fullDeck(5)))) })
  const slides = await g.structureSlides('apuntes', { provider: 'aistudio', apiKey: 'K', maxSlides: 5, timeoutMs: 5000 })
  assert.equal(slides.length, 5)
  assert.equal(call, 1, 'llamadas: ' + call)
})

test('structureSlides acepta el mejor resultado tras agotar reintentos', async () => {
  let call = 0
  const g = load(() => { call++; return Promise.resolve(gemini(deckJson(shortDeck()))) })
  const slides = await g.structureSlides('apuntes', { provider: 'aistudio', apiKey: 'K', maxSlides: 15, timeoutMs: 5000 })
  assert.equal(call, 3, 'máximo 1 + 2 reintentos: ' + call)
  assert.equal(slides.length, shortDeck().length)
})
