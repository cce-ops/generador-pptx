'use strict'
/* Generador PPTX. Tarea 1: shell + setStatus. Tarea 2: entradas a texto plano. */

const MAX_BYTES = 5000000

function setStatus(phase, msg) {
  const el = document.getElementById('out-status')
  if (!el)
    return
  el.setAttribute('data-phase', phase)
  el.textContent = msg
}

function readAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = function () { resolve(r.result) }
    r.onerror = function () { reject(new Error('No se pudo leer el fichero.')) }
    r.readAsArrayBuffer(file)
  })
}

// ===== EXTRACCIÓN DE IMÁGENES (parte B) =====
// Reglas comunes a docx y pdf: solo "figuras" (no iconos), sin duplicados,
// tope de cantidad y de peso. El texto queda con marcas [IMAGEN n] para que
// la IA sepa dónde encaja cada figura.
const IMG_MIN_SIDE = 100
const IMG_MAX_SIDE = 1400
const IMG_MAX_BYTES = 1500000
const IMG_MAX_COUNT = 20
const IMG_MAX_TOTAL = 14000000 // ~14 MB de data-URLs: corta si un PDF lleva escaneos gigantes

function loadImg(src) {
  return new Promise((resolve, reject) => {
    const im = new Image()
    im.onload = () => { resolve(im) }
    im.onerror = () => { reject(new Error('imagen ilegible')) }
    im.src = src
  })
}

// Normaliza una imagen: descarta iconos (<100px), reescala gigantes (>1400px),
// recodifica si pesa demasiado. Devuelve null si no es una figura usable.
function finishImage(src) {
  return loadImg(src).then((im) => {
    const w0 = im.naturalWidth || im.width
    const h0 = im.naturalHeight || im.height
    if (!w0 || !h0 || w0 < IMG_MIN_SIDE || h0 < IMG_MIN_SIDE) return null
    const sc = Math.min(1, IMG_MAX_SIDE / Math.max(w0, h0))
    const w = Math.round(w0 * sc)
    const h = Math.round(h0 * sc)
    const cv = document.createElement('canvas')
    cv.width = w
    cv.height = h
    cv.getContext('2d').drawImage(im, 0, 0, w, h)
    let data = cv.toDataURL('image/png')
    if (data.length > IMG_MAX_BYTES) {
      const cv2 = document.createElement('canvas')
      cv2.width = w
      cv2.height = h
      cv2.getContext('2d').drawImage(cv, 0, 0)
      data = cv2.toDataURL('image/jpeg', 0.85)
    }
    return { data, w, h }
  }).catch(() => null)
}

// Procesa la lista de data-URLs en orden: filtra, deduplica y renumera.
// map[orig] = nuevo índice (1-based) para renumerar los marcadores del texto.
function finishImageList(srcs) {
  const kept = []
  const map = {}
  const seen = new Set()
  let total = 0
  let chain = Promise.resolve()
  srcs.forEach((src, i) => {
    chain = chain.then(() => {
      if (kept.length >= IMG_MAX_COUNT || total >= IMG_MAX_TOTAL || !src) return
      return finishImage(src).then((im) => {
        if (!im) return
        const key = im.w + 'x' + im.h + ':' + im.data.length + ':' + im.data.slice(-48)
        if (seen.has(key)) return
        seen.add(key)
        map[i + 1] = kept.length + 1
        total += im.data.length
        kept.push(im)
      })
    })
  })
  return chain.then(() => ({ images: kept, map }))
}

function renumberMarkers(text, map) {
  return String(text).replace(/\[IMAGEN\s*(\d+)\]/g, (m, n) => {
    const ni = map[Number(n)]
    return ni ? '[IMAGEN ' + ni + ']' : ''
  })
}

function htmlToText(html) {
  return String(html)
    .replace(/<\/(p|h[1-6]|li|tr|div)>/gi, '\n')
    .replace(/<br[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n')
    .trim()
}

// docx: convertToHtml con imágenes como data-URL, en orden del documento.
function extractDocx(buf) {
  return mammoth.convertToHtml({ arrayBuffer: buf }, {
    convertImage: mammoth.images.imgElement((image) => {
      return image.read('base64').then((b64) => ({ src: 'data:' + image.contentType + ';base64,' + b64 }))
    })
  }).then((r) => {
    const srcs = []
    const html = String(r.value).replace(/<img[^>]*?src="([^"]*)"[^>]*?>/gi, (m, src) => {
      srcs.push(src)
      return '\n[IMAGEN ' + srcs.length + ']\n'
    })
    const text = htmlToText(html)
    return finishImageList(srcs).then((res) => ({ text: renumberMarkers(text, res.map), images: res.images }))
  })
}

// pdf: recorre la lista de operadores, resuelve las imágenes pintadas y las
// compone en canvas. El marcador va al final de la página donde aparecen.
function pdfImgToSrc(img) {
  if (!img) return null
  const w = img.width
  const h = img.height
  if (!w || !h) return null
  try {
    const cv = document.createElement('canvas')
    cv.width = w
    cv.height = h
    const ctx = cv.getContext('2d')
    if (img.bitmap) {
      // pdf.js moderno deja el píxel decodificado en un ImageBitmap
      ctx.drawImage(img.bitmap, 0, 0)
    } else if (img.data && img.data.length) {
      const d = img.data
      let rgba
      if (d.length === w * h * 4) rgba = d
      else if (d.length === w * h * 3) {
        rgba = new Uint8ClampedArray(w * h * 4)
        for (let i = 0, j = 0; i < w * h * 3; i += 3, j += 4) {
          rgba[j] = d[i]; rgba[j + 1] = d[i + 1]; rgba[j + 2] = d[i + 2]; rgba[j + 3] = 255
        }
      } else return null
      ctx.putImageData(new ImageData(rgba, w, h), 0, 0)
    } else return null
    return cv.toDataURL('image/png')
  } catch (e) { return null }
}

function extractPdf(buf) {
  if (typeof pdfjsLib !== 'undefined' && pdfjsLib.GlobalWorkerOptions && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
    // worker autoalojado en el repo (mismo origen): sin dependencia del CDN y con worker real.
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'pdf.worker.js'
  }
  // isEvalSupported:false cierra el CVE-2024-4367 (JS arbitrario al parsear fuentes maliciosas).
  return pdfjsLib.getDocument({ data: buf, isEvalSupported: false }).promise.then((pdf) => {
    const OPS = pdfjsLib.OPS
    const pages = []
    const srcs = []
    const srcPage = []
    let chain = Promise.resolve()
    for (let i = 1; i <= pdf.numPages; i++) {
      (function (n) {
        chain = chain.then(() => {
          return pdf.getPage(n).then((page) => {
            return Promise.all([
              page.getTextContent().then((tc) => {
                pages.push({ page: n, text: tc.items.map((it) => { return it.str }).join(' ') })
              }),
              page.getOperatorList().then((opList) => {
                for (let j = 0; j < opList.fnArray.length; j++) {
                  const fn = opList.fnArray[j]
                  if (fn !== OPS.paintImageXObject && fn !== OPS.paintJpegXObject && fn !== OPS.paintInlineImageXObject) continue
                  let img = null
                  try {
                    img = (fn === OPS.paintInlineImageXObject) ? opList.argsArray[j] : page.objs.get(opList.argsArray[j][0])
                  } catch (e) { img = null }
                  const src = pdfImgToSrc(img)
                  if (src) { srcs.push(src); srcPage.push(n) }
                }
              })
            ])
          })
        })
      })(i)
    }
    return chain.then(() => finishImageList(srcs)).then((res) => {
      const byPage = {}
      Object.keys(res.map).forEach((orig) => {
        const ni = res.map[orig]
        const pg = srcPage[orig - 1]
        ;(byPage[pg] = byPage[pg] || []).push(ni)
      })
      pages.sort((a, b) => a.page - b.page)
      const text = pages.map((t) => {
        let s = t.text
        ;(byPage[t.page] || []).forEach((ni) => { s += '\n[IMAGEN ' + ni + ']\n' })
        return s
      }).join('\n')
      return { text, images: res.images }
    })
  })
}

function getInputText() {
  setStatus('leyendo', 'Leyendo apuntes…')
  const mdFile = document.getElementById('in-md').files[0]
  const docxFile = document.getElementById('in-docx').files[0]
  const pdfFile = document.getElementById('in-pdf').files[0]
  const file = mdFile || docxFile || pdfFile
  const done = function (text, label, images) {
    const n = (images || []).length
    setStatus('listo', `Apuntes leídos: ${label}${n ? ` — ${n} imagen${n === 1 ? '' : 'es'}` : ''}.`)
    return { text, label, images: images || [] }
  }
  const fail = function (msg) {
    setStatus('error', msg)
    return new Error(msg)
  }
  if (file) {
    if (file.size > MAX_BYTES)
      return Promise.reject(fail('Archivo demasiado grande (máx. 5 MB).'))
    if (docxFile && file === docxFile) {
      return readAsArrayBuffer(file).then(extractDocx).then((r) => { return done(r.text, file.name, r.images) }, () => { throw fail('No se pudo leer el fichero.') })
    }
    if (pdfFile && file === pdfFile) {
      return readAsArrayBuffer(file).then(extractPdf).then((r) => { return done(r.text, file.name, r.images) }, () => { throw fail('No se pudo leer el fichero.') })
    }
    return file.text().then((t) => { return done(t, file.name, []) }, () => { throw fail('No se pudo leer el fichero.') })
  }
  const pasted = document.getElementById('in-text').value
  if (!pasted.trim())
    return Promise.reject(fail('Pega texto o sube un fichero (.md, .docx, .pdf).'))
  return Promise.resolve(done(pasted, 'texto pegado', []))
}

const DEFAULT_MAX_SLIDES = 15
const MAX_CHARS = 60000

// ===== PATTERN ENGINE: THEMES, HELPERS, MASTERS =====
const THEMES = {
  seriph: {
    name: 'Seriph claro',
    bg: 'FFFFFF', ink: '1A1A1A', muted: '555555', accent: '4D83C3',
    fontTitle: 'Georgia', fontBody: 'Georgia',
    cardBg: 'F5F5F5', cardTitle: '1A1A1A', cardText: '1A1A1A',
    tableHeaderBg: '1A1A1A', tableHeaderFg: 'FFFFFF',
    footer: '888888', badgeFill: '4D83C3', badgeText: 'FFFFFF',
    coverBg: '4D83C3', coverTitle: 'FFFFFF', coverSub: 'FFFFFF',
    accent2: '4D83C3', accent3: '4D83C3', accent4: '4D83C3',
    accent5: '4D83C3', accent6: '4D83C3',
    redTint: 'F5F5F5', grid: 'DDDDDD'
  },
  verde: {
    name: 'Verde sostenible',
    bg: 'FFFFFF', ink: '1A2B23', muted: '4A5A52', accent: '1B7A4D',
    fontTitle: 'Georgia', fontBody: 'Georgia',
    cardBg: 'EDF4EA', cardTitle: '1A2B23', cardText: '1A2B23',
    tableHeaderBg: '1B7A4D', tableHeaderFg: 'FFFFFF',
    footer: '888888', badgeFill: '1B7A4D', badgeText: 'FFFFFF',
    coverBg: '1B7A4D', coverTitle: 'FFFFFF', coverSub: 'FFFFFF',
    accent2: '1B7A4D', accent3: '1B7A4D', accent4: '1B7A4D',
    accent5: '1B7A4D', accent6: '1B7A4D',
    redTint: 'F6E3DE', grid: 'C9D8CF'
  },
  oscuro: {
    name: 'Oscuro',
    bg: '1A1A1A', ink: 'F2F2F2', muted: 'B8B8B8', accent: '7FD1A8',
    fontTitle: 'Georgia', fontBody: 'Georgia',
    cardBg: '2A2A2A', cardTitle: 'F2F2F2', cardText: 'E8E8E8',
    tableHeaderBg: '7FD1A8', tableHeaderFg: '1A1A1A',
    footer: '999999', badgeFill: '7FD1A8', badgeText: '1A1A1A',
    coverBg: '1A1A1A', coverTitle: 'FFFFFF', coverSub: '7FD1A8',
    accent2: '7FD1A8', accent3: '7FD1A8', accent4: '7FD1A8',
    accent5: '7FD1A8', accent6: '7FD1A8',
    redTint: '3A3A3A', grid: '444444'
  },
  corporativo: {
    name: 'Azul corporativo',
    bg: 'FFFFFF', ink: '1A2B3C', muted: '4A6076', accent: '0F3B63',
    fontTitle: 'Georgia', fontBody: 'Georgia',
    cardBg: 'E8EFF5', cardTitle: '1A2B3C', cardText: '1A2B3C',
    tableHeaderBg: '0F3B63', tableHeaderFg: 'FFFFFF',
    footer: '888888', badgeFill: '0F3B63', badgeText: 'FFFFFF',
    coverBg: '0F3B63', coverTitle: 'FFFFFF', coverSub: 'FFFFFF',
    accent2: '0F3B63', accent3: '0F3B63', accent4: '0F3B63',
    accent5: '0F3B63', accent6: '0F3B63',
    redTint: 'E8EFF5', grid: 'CCCCCC'
  },
  upc: {
    name: 'Blanco y azul',
    bg: 'FFFFFF', ink: '1A1A1A', muted: '4A5568', accent: '4285F4',
    fontTitle: 'Georgia', fontBody: 'Georgia',
    cardBg: 'E8F0FE', cardTitle: '1A1A1A', cardText: '1A1A1A',
    tableHeaderBg: '4285F4', tableHeaderFg: 'FFFFFF',
    footer: '888888', badgeFill: '4285F4', badgeText: 'FFFFFF',
    coverBg: '4285F4', coverTitle: 'FFFFFF', coverSub: 'FFFFFF',
    accent2: '4285F4', accent3: '4285F4', accent4: '4285F4',
    accent5: '4285F4', accent6: '4285F4',
    redTint: 'E8F0FE', grid: 'CCCCCC'
  },
  verdeSkill: {
    name: 'Verde skill (bosque/lima/ámbar)',
    bg: 'FFFFFF', ink: '17302A', muted: '6B7F76', accent: '2E7D4F',
    fontTitle: 'Cambria', fontBody: 'Calibri',
    cardBg: 'EDF4EA', cardTitle: '17302A', cardText: '17302A',
    tableHeaderBg: '1F4D3A', tableHeaderFg: 'FFFFFF',
    footer: '6B7F76', badgeFill: '9BC53D', badgeText: '17302A',
    coverBg: '1F4D3A', coverTitle: 'FFFFFF', coverSub: 'E0A526',
    accent2: '9BC53D', accent3: '1C7293', accent4: 'E0A526',
    accent5: 'C8553D', accent6: '6B7F76',
    redTint: 'F6E3DE', grid: 'C9D8CF'
  }
}

// Helpers (puras, reciben theme y SH)
function tx(s, text, x, y, w, h, o, theme) {
  const opts = Object.assign({
    x, y, w, h,
    isTextBox: true,
    margin: 0,
    fontFace: o.fontFace || theme.fontBody,
    fontSize: o.fontSize || 14,
    color: o.color || theme.ink,
    valign: o.valign || 'top'
  }, o)
  return s.addText(text, opts)
}

function card(s, x, y, w, h, fill, name, theme, SH) {
  return s.addShape(SH.ROUNDED_RECTANGLE, {
    x, y, w, h,
    rectRadius: 0.12,
    fill: { color: fill },
    line: { color: fill },
    objectName: name || 'Tarjeta'
  })
}

function badge(s, n, x, y, d, theme, SH) {
  d = d || 0.5
  const shape = s.addShape(SH.OVAL, {
    x, y, w: d, h: d,
    fill: { color: theme.badgeFill },
    line: { color: theme.badgeFill },
    objectName: 'Badge ' + n
  })
  tx(s, String(n), x, y, d, d, {
    align: 'center',
    valign: 'middle',
    bold: true,
    fontSize: 16,
    color: theme.badgeText
  }, theme)
  return shape
}

function stat(s, x, y, w, num, label, color, theme, SH, fontSize) {
  tx(s, num, x, y, w, 0.9, {
    fontFace: theme.fontTitle,
    fontSize: fontSize || 36,
    bold: true,
    color: color || theme.accent,
    valign: 'middle'
  }, theme)
  tx(s, label, x, y + 0.9, w, 0.8, {
    fontSize: 14,
    valign: 'top'
  }, theme)
}

function takeaway(s, text, y, theme, SH) {
  y = y || 4.6
  return tx(s, [
    { text: 'Idea clave ', options: { bold: true, color: theme.accent2 } },
    { text: text, options: { color: theme.ink } }
  ], 0.8, y, 8.4, 0.85, { valign: 'middle', fit: 'shrink' }, theme)
}

function coverMotif(s, theme, SH) {
  s.addShape(SH.OVAL, {
    x: 6.5, y: 0.7, w: 3.0, h: 3.0,
    fill: { color: theme.accent2, transparency: 15 },
    line: { color: theme.accent2, transparency: 15 }
  })
  s.addShape(SH.OVAL, {
    x: 7.9, y: 3.0, w: 1.7, h: 1.7,
    fill: { color: theme.accent4, transparency: 20 },
    line: { color: theme.accent4, transparency: 20 }
  })
}

// Define 3 masters: PORTADA, CONTENIDO, CIERRE
function defineMasters(pres, theme) {
  const SH = pres.shapes

  // PORTADA: fondo coverBg, placeholders título/subtítulo
  pres.defineSlideMaster({
    title: 'PORTADA',
    background: { color: theme.coverBg },
    objects: [
      {
        placeholder: {
          options: {
            name: 'title', type: 'title',
            x: 0.7, y: 1.4, w: 5.4, h: 1.8,
            fontSize: 38, bold: true, color: theme.coverTitle,
            margin: 0, valign: 'bottom', align: 'left'
          },
          text: ''
        }
      },
      {
        placeholder: {
          options: {
            name: 'body', type: 'body',
            x: 0.7, y: 3.4, w: 5.4, h: 0.9,
            fontSize: 18, color: theme.coverSub,
            margin: 0
          },
          text: ''
        }
      }
    ]
  })

  // CONTENIDO: fondo bg, pie con número de página en y=5.2
  pres.defineSlideMaster({
    title: 'CONTENIDO',
    background: { color: theme.bg },
    slideNumber: {
      x: 9.0, y: 5.2, w: 0.5, h: 0.3,
      fontSize: 10, color: theme.footer
    },
    objects: [
      {
        placeholder: {
          options: {
            name: 'title', type: 'title',
            x: 0.5, y: 0.3, w: 9, h: 0.8,
            fontSize: 28, bold: true, color: theme.ink,
            margin: 0, valign: 'middle', align: 'left'
          },
          text: ''
        }
      },
      {
        text: {
          text: 'Generador PPTX · Diseño sostenible',
          options: {
            x: 0.5, y: 5.2, w: 6, h: 0.3,
            fontSize: 10, color: theme.muted, margin: 0
          }
        }
      }
    ]
  })

  // CIERRE: fondo coverBg, solo título
  pres.defineSlideMaster({
    title: 'CIERRE',
    background: { color: theme.coverBg },
    objects: [
      {
        placeholder: {
          options: {
            name: 'title', type: 'title',
            x: 0.5, y: 0.3, w: 9, h: 0.8,
            fontSize: 28, bold: true, color: theme.coverTitle,
            margin: 0, valign: 'middle', align: 'left'
          },
          text: ''
        }
      }
    ]
  })

  return { PORTADA: 'PORTADA', CONTENIDO: 'CONTENIDO', CIERRE: 'CIERRE' }
}

// Export for testing (Node vm context)
if (typeof globalThis !== 'undefined') {
  globalThis.__engine = {
    THEMES,
    tx,
    card,
    badge,
    stat,
    takeaway,
    coverMotif,
    defineMasters
  }
}

// ===== RENDERIZADORES POR KIND =====
// Tamaño de fuente que cabe en la caja de forma determinista.
// normAutofit no basta: PowerPoint no recompone el texto hasta que se edita,
// así que aquí se calcula (factor 0.60 de ancho de letra + 1.25 de interlineado).
function autoFont(text, w, h, max, min) {
  const t = String(text || '')
  const lo = min || 12
  let sz = max
  while (sz > lo) {
    const cpl = Math.max(1, (w * 72) / (sz * 0.60))
    const lines = Math.max(1, Math.ceil(t.length / cpl))
    if (lines * sz * 1.25 <= h * 72) break
    sz -= 1
  }
  return sz
}

function contentTitle(s, title, theme) {
  tx(s, title, 0.5, 0.3, 9, 0.8, { fontFace: theme.fontTitle, fontSize: autoFont(title, 9, 0.8, 28, 18), bold: true, color: theme.ink, align: 'left', valign: 'middle', fit: 'shrink' }, theme)
}

function titleBar(s, theme, SH) {
  s.addShape('rect', { x: 0.5, y: 1.18, w: 1.6, h: 0.06, fill: { color: theme.accent }, line: { color: theme.accent } })
}

function render_cover(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  tx(s, d.title, 0.7, 1.5, 5.6, 1.8, { fontFace: theme.fontTitle, fontSize: autoFont(d.title, 5.6, 1.8, 38, 18), bold: true, color: theme.coverTitle, align: 'left', valign: 'middle', fit: 'shrink' }, theme)
  if (d.subtitle) tx(s, d.subtitle, 0.7, 3.4, 5.6, 1.0, { fontFace: theme.fontBody, fontSize: autoFont(d.subtitle, 5.6, 1.0, 18, 12), color: theme.coverSub, align: 'left', valign: 'middle', fit: 'shrink' }, theme)
  s.addShape(SH.OVAL, {
    x: 6.6, y: 0.55, w: 2.9, h: 2.9,
    fill: { color: theme.accent2, transparency: 15 },
    line: { color: theme.accent2, transparency: 15 }
  })
  s.addShape(SH.OVAL, {
    x: 7.8, y: 2.75, w: 1.5, h: 1.5,
    fill: { color: theme.accent4, transparency: 20 },
    line: { color: theme.accent4, transparency: 20 }
  })
}

function render_stats(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const items = d.items || []
  const n = items.length
  const gap = 9 / Math.max(n, 1)
  items.forEach((it, i) => {
    const x = 0.5 + i * gap
    const w = gap - 0.3
    card(s, x, 1.5, w, 3.0, theme.cardBg, null, theme, SH)
    tx(s, it.value, x + 0.25, 1.95, w - 0.5, 1.25, { fontFace: theme.fontTitle, fontSize: autoFont(it.value, w - 0.5, 1.25, 36, 16), bold: true, color: themeColor(it.color, theme, theme.accent), valign: 'middle', fit: 'shrink' }, theme)
    tx(s, it.label, x + 0.25, 3.3, w - 0.5, 1.0, { fontSize: 14, color: theme.ink, valign: 'top', fit: 'shrink' }, theme)
  })
  if (d.takeaway) takeaway(s, d.takeaway, 4.6, theme, SH)
}

function render_process(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const steps = d.steps || []
  const n = steps.length
  const gap = 9 / Math.max(n, 1)
  const y0 = d.takeaway ? 1.55 : 1.9
  steps.forEach((st, i) => {
    const x = 0.5 + i * gap
    const w = gap - 0.2
    badge(s, i + 1, x + w / 2 - 0.3, y0, 0.6, theme, SH)
    tx(s, st[0], x, y0 + 0.75, w, 0.7, { fontFace: theme.fontTitle, fontSize: 16, bold: true, color: theme.ink, align: 'center', fit: 'shrink' }, theme)
    tx(s, st[1], x, y0 + 1.5, w, d.takeaway ? 1.3 : 1.5, { fontSize: 13, align: 'center', valign: 'top', fit: 'shrink' }, theme)
  })
  if (d.takeaway) takeaway(s, d.takeaway, 4.6, theme, SH)
}

function render_table(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const rows = d.rows || []
  const data = [d.headers].concat(rows).map((r, i) =>
    r.map((c, j) => i === 0
      ? { text: c, options: { bold: true, color: theme.tableHeaderFg, fill: { color: theme.tableHeaderBg }, fontFace: theme.fontBody, fontSize: 14, valign: 'middle' } }
      : { text: c, options: { fontFace: theme.fontBody, fontSize: 12, bold: j === 0, color: theme.ink, valign: 'middle', paraSpaceAfter: 4 } })
  )
  const colW = Array(d.headers.length).fill(9 / d.headers.length)
  const top = 1.35
  const bottom = d.formula ? 4.3 : 5.15
  const rowH = Math.min(1.6, Math.max(0.45, (bottom - top - 0.5) / Math.max(rows.length, 1)))
  s.addTable(data, {
    x: 0.5, y: top, w: 9,
    colW,
    rowH: [0.5].concat(rows.map(() => rowH)),
    border: { type: 'solid', color: theme.grid, pt: 1 },
    fill: { color: theme.redTint }
  })
  if (d.formula) takeaway(s, d.formula, 4.5, theme, {})
}

function render_formula(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  tx(s, d.formula, 0.5, 1.35, 9, 1.1, {
    fontFace: theme.fontTitle, fontSize: 28, bold: true, color: theme.accent, align: 'center', valign: 'middle', fit: 'shrink'
  }, theme)
  const cols = d.columns || []
  const n = cols.length
  const gap = 9 / Math.max(n, 1)
  cols.forEach((c, i) => {
    const x = 0.5 + i * gap
    const w = gap - 0.2
    card(s, x, 2.7, w, 2.25, theme.cardBg, null, theme, s.pres ? s.pres.shapes : { ROUNDED_RECTANGLE: 'roundRect' })
    tx(s, c[0], x + 0.2, 2.9, w - 0.4, 0.5, { bold: true, fontSize: 16, color: theme.cardTitle, fit: 'shrink' }, theme)
    tx(s, c[1], x + 0.2, 3.5, w - 0.4, 1.3, { color: theme.cardText, fit: 'shrink' }, theme)
  })
}

function render_compare(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const items = d.items || []
  const n = items.length
  const gap = 9 / Math.max(n, 1)
  items.forEach((it, i) => {
    const x = 0.5 + i * gap
    const w = gap - 0.2
    card(s, x, 1.5, w, 3.15, theme.cardBg, null, theme, s.pres ? s.pres.shapes : { ROUNDED_RECTANGLE: 'roundRect' })
    tx(s, it.value, x + 0.2, 1.8, w - 0.4, 1.5, { fontFace: theme.fontTitle, fontSize: autoFont(it.value, w - 0.4, 1.5, 24, 13), bold: true, color: themeColor(it.color, theme, theme.accent), valign: 'middle', fit: 'shrink' }, theme)
    tx(s, it.label, x + 0.2, 3.45, w - 0.4, 1.0, { fontSize: 14, color: theme.ink, valign: 'top', fit: 'shrink' }, theme)
  })
}

// Paleta de gráficos: primer color = acento del tema, el resto girando el tono.
// Garantiza colores distintos aunque el tema sea monocromo (5 de los 6 temas
// copian accent2..6 = accent, lo que dejaba todas las series/rebanadas iguales).
function hslToHex(h, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1))
  const m = l - c / 2
  const rgb = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x]
    : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x]
  return rgb.map(v => Math.round((v + m) * 255).toString(16).padStart(2, '0')).join('').toUpperCase()
}

function chartPalette(theme, n) {
  const total = Math.max(1, Math.min(n || 1, 12))
  let accent = String((theme && theme.accent) || '4D83C3').replace('#', '').toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(accent)) accent = '4D83C3'
  const out = [accent]
  if (total === 1) return out
  const r = parseInt(accent.slice(0, 2), 16) / 255
  const g = parseInt(accent.slice(2, 4), 16) / 255
  const b = parseInt(accent.slice(4, 6), 16) / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  let h = 0
  let s = 0
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? ((g - b) / d + (g < b ? 6 : 0)) : max === g ? ((b - r) / d + 2) : ((r - g) / d + 4)
    h *= 60
  }
  // S/L normalizados para que se vean tanto en fondo claro como oscuro
  const li = Math.min(0.66, Math.max(0.42, l))
  const si = Math.min(0.85, Math.max(0.5, s))
  for (let i = 1; i < total; i++) out.push(hslToHex((h + i * (360 / total)) % 360, si, li))
  return out
}

// Gráfico nativo de PowerPoint (editable, no imagen). El tipo llega saneado
// por sanitizeChart; aquí se aplica tema y se reserva hueco al takeaway.
function render_chart(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const ch = d.chart
  if (!ch || !Array.isArray(ch.series) || !ch.series.length || !Array.isArray(ch.categories) || !ch.categories.length) return
  const isPie = ch.type === 'pie' || ch.type === 'doughnut'
  const axisTitle = String(ch.axisTitle || '').trim()
  const series = (isPie ? ch.series.slice(0, 1) : ch.series.slice(0, 4))
  const data = series.map(sr => ({ name: sr.name || '', labels: ch.categories, values: sr.values }))
  // pptxgenjs: 1 serie → chartColors por punto (barra/rebanada); multi-serie → por serie
  const nColors = isPie ? ch.categories.length : (data.length > 1 ? data.length : (data[0].values.length || 1))
  const hasTakeaway = !!d.takeaway
  const CT = (s.pres && s.pres.ChartType) || (typeof PptxGenJS !== 'undefined' && PptxGenJS.ChartType) || {}
  s.addChart(CT[ch.type] || ch.type, data, {
    x: 0.5, y: 1.45, w: 9, h: hasTakeaway ? 3.0 : 3.55,
    chartColors: chartPalette(theme, nColors),
    // el eje de valores dice qué mide (magnitud + unidades); en tarta no hay
    // eje, así que el mismo texto va como título del gráfico
    showTitle: isPie && !!axisTitle,
    title: axisTitle, titleFontFace: theme.fontTitle, titleFontSize: 14, titleColor: theme.ink,
    showValAxisTitle: !isPie && !!axisTitle,
    valAxisTitle: axisTitle, valAxisTitleFontFace: theme.fontBody, valAxisTitleFontSize: 12, valAxisTitleColor: theme.muted,
    showLegend: data.length > 1 || isPie,
    legendPos: 'b', legendFontFace: theme.fontBody, legendFontSize: 11, legendColor: theme.ink,
    catAxisLabelFontFace: theme.fontBody, catAxisLabelFontSize: 11, catAxisLabelColor: theme.ink,
    valAxisLabelFontFace: theme.fontBody, valAxisLabelFontSize: 11, valAxisLabelColor: theme.muted,
    valGridLine: { color: theme.grid, style: 'solid', size: 0.5 },
    catGridLine: { style: 'none' },
    barGapWidthPct: 60,
    showValue: false,
    showPercent: isPie, dataLabelColor: theme.ink, dataLabelFontSize: 11, dataLabelFontFace: theme.fontBody
  })
  if (hasTakeaway) takeaway(s, d.takeaway, 4.6, theme, SH)
}

// Diapositiva con una imagen (figura extraída de los apuntes): se centra
// respetando proporción dentro del área de contenido. Sin imagen → placeholder.
function render_figure(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const im = d.img
  const hasTakeaway = !!d.takeaway
  const boxY = 1.45
  const boxH = hasTakeaway ? 3.0 : 3.55
  if (im && im.data && im.w && im.h) {
    const sc = Math.min(9 / im.w, boxH / im.h)
    const w = im.w * sc
    const h = im.h * sc
    s.addImage({ data: im.data, x: 0.5 + (9 - w) / 2, y: boxY + (boxH - h) / 2, w, h })
  } else {
    card(s, 0.5, boxY, 9, boxH, theme.cardBg, 'Figura', theme, SH)
    tx(s, 'Imagen no disponible en el material', 0.5, boxY, 9, boxH, { fontSize: 16, color: theme.muted, align: 'center', valign: 'middle' }, theme)
  }
  if (hasTakeaway) takeaway(s, d.takeaway, 4.6, theme, SH)
}

function splitKeywords(text) {
  const parts = String(text).split(/\*\*(.+?)\*\*/g)
  const out = []
  for (let i = 0; i < parts.length; i++) {
    if (!parts[i]) continue
    out.push({ text: parts[i], key: i % 2 === 1 })
  }
  return out.length ? out : [{ text: String(text), key: false }]
}

function render_content(s, d, theme) {
  const SH = s.pres ? s.pres.shapes : { OVAL: 'ellipse', RECTANGLE: 'rect', ROUNDED_RECTANGLE: 'roundRect' }
  if (d.title) { contentTitle(s, d.title, theme); titleBar(s, theme, SH) }
  const bullets = d.bullets || []
  const paras = []
  bullets.forEach((b) => {
    const segs = splitKeywords(b)
    segs.forEach((seg, j) => {
      const opts = { fontFace: theme.fontBody, fontSize: 18, color: theme.ink, paraSpaceAfter: 12, lineSpacingMultiple: 1.25 }
      if (j === 0) {
        opts.bullet = { characterCode: '2022' }
        opts.breakLine = segs.length === 1
      }
      if (j === segs.length - 1) opts.breakLine = true
      if (seg.key) { opts.bold = true; opts.color = theme.accent }
      paras.push({ text: seg.text, options: opts })
    })
  })
  if (paras.length) {
    tx(s, paras, 0.7, 1.45, 9, 3.55, { valign: 'top', fit: 'shrink' }, theme)
  }
}

function render_closing(s, d, theme) {
  if (d.title) tx(s, d.title, 0.5, 0.3, 9, 0.8, { fontFace: theme.fontTitle, fontSize: 28, bold: true, color: theme.ink, align: 'left', valign: 'middle', fit: 'shrink' }, theme)
  const items = (d.closing || d.items || []).slice(0, 3)
  items.forEach((t, i) => {
    const y = 1.55 + i * 1.25
    badge(s, i + 1, 0.7, y + 0.2, 0.6, theme, s.pres ? s.pres.shapes : { OVAL: 'ellipse', ROUNDED_RECTANGLE: 'roundRect' })
    tx(s, t, 1.6, y, 7.8, 1.0, { fontSize: 20, color: theme.ink, valign: 'middle', fit: 'shrink' }, theme)
  })
}

function paginateSlides(slides) {
  const pages = [{ cover: true, title: (slides[0] && slides[0].title) || 'Presentación', bullets: [], table: null }]
  slides.forEach((sld, idx) => {
    if (idx === 0 && sld.kind === 'cover') return
    const kind = sld.kind || 'content'
    const title = sld.title || ''
    if (kind === 'content') {
      const bullets = sld.bullets || []
      const chunks = []
      for (let i = 0; i < bullets.length; i += 4) chunks.push(bullets.slice(i, i + 4))
      chunks.forEach((b, k) => pages.push({ title: k ? title + ' (cont.)' : title, bullets: b, table: null }))
    } else if (kind === 'table') {
      const rows = sld.table?.rows || sld.rows || []
      const header = sld.table?.header || sld.table?.headers || sld.headers || []
      const chunks = []
      for (let i = 0; i < rows.length; i += 8) chunks.push(rows.slice(i, i + 8))
      chunks.forEach((r, k) => pages.push({ title: k ? title + ' (cont.)' : title, bullets: [], table: { header, rows: r }, formula: sld.formula }))
    } else if (kind === 'process') {
      const steps = sld.steps || []
      const chunks = []
      for (let i = 0; i < steps.length; i += 4) chunks.push(steps.slice(i, i + 4))
      chunks.forEach((st, k) => pages.push({ title: k ? title + ' (cont.)' : title, bullets: [], process: st, takeaway: sld.takeaway, table: null }))
    } else if (kind === 'stats') {
      if ((sld.items || []).length) pages.push({ title, bullets: [], stats: sld.items, takeaway: sld.takeaway, table: null })
    } else if (kind === 'formula') {
      if (sld.formula) pages.push({ title, bullets: [], formula: sld.formula, columns: sld.columns, table: null })
    } else if (kind === 'chart') {
      if (sld.chart) pages.push({ title, bullets: [], chart: sld.chart, takeaway: sld.takeaway, table: null })
    } else if (kind === 'figure') {
      pages.push({ title, bullets: [], figure: sld.figure, img: sld.img, takeaway: sld.takeaway, table: null })
    } else if (kind === 'compare') {
      const items = sld.compare || sld.items || []
      if (items.length) pages.push({ title, bullets: [], compare: items, table: null })
    } else if (kind === 'closing') {
      const items = sld.closing || sld.items || []
      if (items.length) pages.push({ title, bullets: [], closing: items, table: null })
    } else {
      pages.push({ title, bullets: sld.bullets || [], table: null })
    }
  })
  return pages
}

// Export for testing
if (typeof globalThis !== 'undefined') {
  globalThis.__engine = {
    ...globalThis.__engine,
    autoFont,
    chartPalette,
    assignFigures,
    finishImageList,
    stripMarkers,
    render_cover,
    render_stats,
    render_process,
    render_table,
    render_formula,
    render_compare,
    render_chart,
    render_figure,
    render_content,
    render_closing,
    paginateSlides,
    buildPptx
  }
}

function truncateText(text) {
  const t = String(text)
  if (t.length <= MAX_CHARS)
    return { text: t, truncated: false }
  return { text: t.slice(0, MAX_CHARS), truncated: true }
}

function buildPrompt(text, maxSlides, instrucciones, imgOpts) {
  const n = Math.max(1, Math.min(maxSlides || DEFAULT_MAX_SLIDES, 30))
  const inst = String(instrucciones || '').trim()
  const instSection = inst
    ? `\nCONCEPTOS Y REGLAS OBLIGATORIOS (no te los puedes saltar; prevalecen sobre todo lo anterior):\n`
      + `- Asegúrate de que cada concepto, término o regla indicado por el usuario figure en el título, un bullet o las notas de alguna diapositiva:\n`
      + `${inst.slice(0, 2000)}${inst.length > 2000 ? '…' : ''}\n`
    : ''
  const io = imgOpts || {}
  const nimg = io.imageCount || 0
  const imgSection = !nimg ? ''
    : io.vision
      ? `\nEl documento contiene ${nimg} imágenes, que van ADJUNTAS a esta petición en orden `
        + `(imagen 1 = primera, …). En el texto verás marcas [IMAGEN n] donde encajan. `
        + `Cuando una imagen aporte contenido, crea una diapositiva kind:"figure": `
        + `"title" = pie de foto descriptivo (qué muestra y por qué importa), "figure" = n (número de imagen), `
        + `"takeaway" = una frase corta de conclusión opcional. `
        + `NO copies el marcador [IMAGEN n] a title, bullets ni notes. `
        + `Si una imagen es decorativa o redundante, no le dediques diapositiva.\n`
      : `\nEl documento contiene ${nimg} imágenes, marcadas en el texto como [IMAGEN n] `
        + `(no puedes verlas). Crea diapositivas kind:"figure" para las que parezcan relevantes por su contexto: `
        + `"title" = pie de foto genérico basado en el texto cercano, "figure" = n. `
        + `NO copies los marcadores a title, bullets ni notes.\n`
  return `Genera la estructura de una presentación de exactamente ${n} diapositivas (incluyendo la portada y el cierre) `
    + `a partir de los siguientes apuntes. Una idea por diapositiva; usa "table" solo `
    + `cuando los apuntes contengan datos tabulares. Cada bullet: una sola idea en un máximo de 12 palabras.\n`
    + `Devuelve SOLO JSON válido con este esquema exacto, sin texto adicional:\n`
    + `{"slides": [{"kind": "cover|stats|process|table|formula|compare|chart|figure|content|closing", `
    + `"title": "...", "bullets": ["..."], "items": [{"value": "...", "label": "...", "color": "..."}], `
    + `"steps": [["Fase", "detalle"]], "table": {"header": ["..."], "rows": [["..."]]}, `
    + `"formula": "...", "columns": [["Etiqueta", "Detalle"]], "compare": [{"value": "...", "label": "..."}], `
    + `"chart": {"type": "bar|line|pie|doughnut", "axisTitle": "...", "categories": ["..."], "series": [{"name": "...", "values": [0]}]}, `
    + `"figure": 1, "closing": ["..."], "takeaway": "...", "notes": "..."}]}`
    + `\nKinds obligatorios: cover (portada), stats (KPIs+takeaway), process (pasos+badges+takeaway), `
    + `table (tabla+cabecera+takeaway), formula (fórmula grande+3 columnas), compare (3 cards), `
    + `chart (gráfico nativo de barras/líneas/tarta con categories y series), `
    + (nimg ? `figure (una imagen del documento; solo "figure" y "title"/"figure"/"takeaway"), ` : '')
    + `content (título+bullets≤12 palabras), closing (fondo acento+3 badges).\n`
    + `Reglas: 1 idea visual por slide. Cada slide "content" lleva 3-5 bullets (nunca 1 solo). `
    + `Usa al menos 3 kinds distintos además de cover/closing/content (stats, process, table, formula o compare). `
    + `Nunca devuelvas slides con bullets, items, steps o rows vacíos. `
    + `En bullets, marca 1-2 conceptos clave por bullet con **doble asterisco**. `
    + `Bullets ≤12 palabras, 1 idea c/u. Detalles y fuentes → "notes". `
    + `En "compare", "value" = número o ≤5 palabras y "label" = frase larga. `
    + `Si hay datos tabulares → kind:"table". Si hay secuencia 4-5 pasos → kind:"process". `
    + `Si hay series de datos (valores por categoría, años, cifras) → kind:"chart" (mismo número de values y categories). `
    + `En "chart", incluye siempre "axisTitle" con la magnitud y sus unidades (p. ej. "Consumo anual (MWh)", "Inversión (€)", "Reparto (%)"). `
    + `Si hay 3 dimensiones comparadas → kind:"compare". Si hay fórmula aislada → kind:"formula". `
    + `Portada y cierre obligatorios (el builder los inserta si faltan).\n`
    + imgSection
    + instSection
    + `\nAPUNTES:\n${text}`
}

function extractJson(raw) {
  const t = String(raw).trim()
  const m = t.match(/```(?:json)?\s*([\s\S]*?)```/)
  const candidate = m ? m[1].trim() : t
  try { return JSON.parse(candidate) }
  catch (e) {}
  const start = candidate.indexOf('{')
  const end = candidate.lastIndexOf('}')
  if (start >= 0 && end > start) {
    try { return JSON.parse(candidate.slice(start, end + 1)) }
    catch (e2) {}
  }
  throw new Error('Respuesta sin JSON válido')
}

function themeColor(value, theme, fallback) {
  const fb = fallback || (theme && theme.accent) || '000000'
  if (value === undefined || value === null) return fb
  const v = String(value).trim().replace(/^#/, '')
  if (/^[0-9A-Fa-f]{6}$/.test(v)) return v.toUpperCase()
  if (theme && typeof theme[v.toLowerCase()] === 'string' && /^[0-9A-Fa-f]{6}$/.test(theme[v.toLowerCase()])) {
    return theme[v.toLowerCase()].toUpperCase()
  }
  return fb
}

function hasRenderableContent(slide) {
  switch (slide.kind) {
    case 'cover': return true
    case 'stats': return Array.isArray(slide.items) && slide.items.length > 0
    case 'process': return Array.isArray(slide.steps) && slide.steps.length > 0
    case 'table': return !!(slide.table && Array.isArray(slide.table.rows) && slide.table.rows.length > 0)
    case 'formula': return !!(slide.formula && slide.formula.trim())
    case 'compare': return Array.isArray(slide.compare) && slide.compare.length > 0
    case 'chart': return !!(slide.chart && Array.isArray(slide.chart.categories) && slide.chart.categories.length
      && Array.isArray(slide.chart.series) && slide.chart.series.length)
    case 'figure': return true
    case 'closing': return Array.isArray(slide.closing) && slide.closing.length > 0
    default: return Array.isArray(slide.bullets) && slide.bullets.length > 0
  }
}

// Título de reserva cuando la IA devuelve un slide sin título (bug "7. " vacío)
const KIND_TITLES = {
  cover: 'Presentación', stats: 'Cifras clave', process: 'Proceso paso a paso',
  table: 'Tabla de datos', formula: 'Fórmula clave', compare: 'Comparativa',
  chart: 'Evolución de datos', closing: 'Ideas clave', content: 'Contenido',
  figure: 'Referencia visual',
}
function fallbackTitle(kind, slide) {
  if (kind === 'content' && Array.isArray(slide.bullets) && slide.bullets.length) {
    const w = String(slide.bullets[0]).replace(/\*\*/g, '').split(/\s+/).slice(0, 6).join(' ').trim()
    if (w) return w
  }
  return KIND_TITLES[kind] || 'Contenido'
}

// Marcadores [IMAGEN n] que no deben colarse en títulos, bullets ni notas
function stripMarkers(x) {
  return String(x == null ? '' : x).replace(/\[IMAGEN\s*\d*\]/g, '').replace(/\s+/g, ' ').trim()
}

// Bullets: una idea por bullet, tope de 12 palabras (no 12 caracteres)
function normBullets(bullets) {
  const b = Array.isArray(bullets) ? bullets.map(x => stripMarkers(x)).filter(x => x) : []
  return b.map((x) => {
    const words = x.split(/\s+/)
    return words.length > 12 ? words.slice(0, 12).join(' ') + '…' : x
  })
}

// Tipos de gráfico que acepta el contrato (la IA puede escribir alias)
const CHART_TYPES = {
  bar: 'bar', column: 'bar', col: 'bar', barras: 'bar',
  line: 'line', lines: 'line', lineas: 'line',
  pie: 'pie', tarta: 'pie', doughnut: 'doughnut', donut: 'doughnut',
}

// Sanea el gráfico de la IA: type conocido, series ≤4, categories ≤12,
// valores numéricos (no finitos → 0) y categorías/valores alineados por longitud.
// Devuelve null si no hay datos renderizables.
function sanitizeChart(c) {
  if (!c || typeof c !== 'object') return null
  // Object.hasOwn evita resolver propiedades heredadas (__proto__, constructor…) que
  // serian truthy y saltarian el fallback (robustez; ver auditoria hallazgo 5).
  const tkey = String(c.type == null ? '' : c.type).trim().toLowerCase()
  const type = (Object.prototype.hasOwnProperty.call(CHART_TYPES, tkey) ? CHART_TYPES[tkey] : null) || 'bar'
  const axisTitle = String(c.axisTitle == null ? '' : c.axisTitle).trim().slice(0, 48)
  let cats = (Array.isArray(c.categories) ? c.categories : [])
    .map(v => String(v == null ? '' : v).trim())
  let series = (Array.isArray(c.series) ? c.series : [])
    .map(sr => ({
      name: String((sr && sr.name) || '').trim() || 'Serie',
      values: (Array.isArray(sr && sr.values) ? sr.values : [])
        .map(v => { const n = Number(v); return Number.isFinite(n) ? n : 0 })
    }))
    .filter(sr => sr.values.length > 0)
    .slice(0, 4)
  if ((type === 'pie' || type === 'doughnut') && series.length > 1) series = series.slice(0, 1)
  if (!series.length) return null
  if (!cats.length) cats = series[0].values.map((_, i) => String(i + 1))
  const len = Math.min(cats.length, 12, ...series.map(sr => sr.values.length))
  if (len < 1) return null
  return {
    type,
    ...(axisTitle ? { axisTitle } : {}),
    categories: cats.slice(0, len).map((c2, i) => c2 || String(i + 1)),
    series: series.map(sr => ({ name: sr.name, values: sr.values.slice(0, len) }))
  }
}

function normalizeSlides(data, max) {
  if (!data || !Array.isArray(data.slides))
    throw new Error('Respuesta sin lista de diapositivas')
  const n = Math.max(1, Math.min(max || DEFAULT_MAX_SLIDES, 30))
  const validKinds = ['cover', 'stats', 'process', 'table', 'formula', 'compare', 'chart', 'figure', 'content', 'closing']
  const out = []
  for (let i = 0; i < data.slides.length && out.length < n; i++) {
    const s = data.slides[i] || {}
    let kind = validKinds.includes(s.kind) ? s.kind : 'content'
    const slide = { kind, title: stripMarkers(s.title) }
    if (kind === 'stats') {
      slide.items = Array.isArray(s.items) ? s.items.map(it => ({
        value: String(it.value || '').trim(), label: String(it.label || '').trim(),
        color: it.color || 'accent'
      })).filter(it => it.value) : []
      if (s.takeaway) slide.takeaway = String(s.takeaway).trim()
    } else if (kind === 'process') {
      slide.steps = Array.isArray(s.steps) ? s.steps.filter(Array.isArray).map(st => [String(st[0] || '').trim(), String(st[1] || '').trim()]).filter(st => st[0] || st[1]) : []
      if (s.takeaway) slide.takeaway = String(s.takeaway).trim()
    } else if (kind === 'table') {
      slide.table = {}
      if (Array.isArray(s.table?.header) && Array.isArray(s.table?.rows)) {
        slide.table.header = s.table.header.map(h => String(h).trim())
        slide.table.rows = s.table.rows.filter(Array.isArray).map(r => r.map(c => String(c).trim())).filter(r => r.some(c => c))
      }
      if (s.formula) slide.formula = String(s.formula).trim()
    } else if (kind === 'formula') {
      slide.formula = String(s.formula || '').trim()
      slide.columns = Array.isArray(s.columns) ? s.columns.filter(Array.isArray).map(c => [String(c[0] || '').trim(), String(c[1] || '').trim()]).filter(c => c[0] || c[1]) : []
    } else if (kind === 'compare') {
      const citems = Array.isArray(s.compare) ? s.compare : (Array.isArray(s.items) ? s.items : [])
      slide.compare = citems.map(it => ({
        value: String(it.value || '').trim(), label: String(it.label || '').trim()
      })).filter(it => it.value)
    } else if (kind === 'chart') {
      const ch = sanitizeChart(s.chart)
      if (ch) slide.chart = ch
      else { kind = 'content'; slide.bullets = normBullets(s.bullets) }
    } else if (kind === 'figure') {
      const fi = parseInt(s.figure, 10)
      if (Number.isFinite(fi) && fi > 0) slide.figure = fi
      if (s.takeaway) slide.takeaway = stripMarkers(s.takeaway)
      slide.bullets = []
    } else if (kind === 'closing') {
      slide.closing = Array.isArray(s.closing) ? s.closing.map(s2 => String(s2).trim()).filter(s2 => s2) : []
    } else {
      slide.bullets = normBullets(s.bullets)
    }
    // Procesar tabla si existe (para cualquier kind que la tenga)
    if (s.table && Array.isArray(s.table?.header) && Array.isArray(s.table?.rows)) {
      slide.table = {
        header: s.table.header.map(h => String(h).trim()),
        rows: s.table.rows.filter(Array.isArray).map(r => r.map(c => String(c).trim())).filter(r => r.some(c => c))
      }
      if (s.formula) slide.formula = String(s.formula).trim()
    }
    if (s.notes) slide.notes = stripMarkers(s.notes)
    if (s.takeaway && !slide.takeaway) slide.takeaway = stripMarkers(s.takeaway)
    if (s.subtitle && !slide.subtitle) slide.subtitle = String(s.subtitle).trim()
    const full = { ...slide, kind }
    if (!full.title) full.title = fallbackTitle(kind, full)
    if (kind === 'cover' || hasRenderableContent(full)) out.push(full)
  }
  // Añadir cover si falta (titula con la primera slide conservada, no con descartes)
  if (!out.length || out[0].kind !== 'cover') {
    out.unshift({ kind: 'cover', title: (out[0] && out[0].title) || 'Presentación', subtitle: '' })
  }
  // Añadir closing si falta
  if (out.length && out[out.length - 1].kind !== 'closing') {
    out.push({ kind: 'closing', title: 'Ideas clave', closing: ['Idea principal', 'Segunda idea', 'Tercera idea'] })
  }
  return out.slice(0, n)
}

// Asigna las imágenes extraídas a los slides kind:"figure".
// mode 'vision' → respeta el índice "figure" que decidió la IA (lo no usado se
// descarta a propósito, p. ej. logos). mode 'blind' → coloca en orden secuencial
// y las sobrantes van como anexo al final (antes del cierre) para no perderlas.
function assignFigures(slides, images, mode) {
  const imgs = Array.isArray(images) ? images : []
  const placeholder = (s) => ({
    kind: 'content', title: s.title, bullets: ['Imagen no disponible en el material'],
    takeaway: s.takeaway, notes: s.notes
  })
  if (!imgs.length) return slides.map((s) => (s.kind === 'figure' ? placeholder(s) : s))
  const used = new Set()
  const out = []
  slides.forEach((s) => {
    if (s.kind !== 'figure') { out.push(s); return }
    let idx = -1
    const want = parseInt(s.figure, 10)
    if (Number.isFinite(want) && want >= 1 && want <= imgs.length && !used.has(want - 1)) idx = want - 1
    if (idx < 0) { for (let k = 0; k < imgs.length; k++) if (!used.has(k)) { idx = k; break } }
    if (idx < 0) { out.push(placeholder(s)); return }
    used.add(idx)
    out.push({ ...s, figure: idx + 1, img: imgs[idx] })
  })
  if (mode === 'blind') {
    const leftovers = []
    imgs.forEach((im, k) => { if (!used.has(k)) leftovers.push(k) })
    if (leftovers.length) {
      const annex = leftovers.map((k) => ({ kind: 'figure', title: 'Imagen ' + (k + 1) + ' del material', figure: k + 1, img: imgs[k] }))
      const ci = out.findIndex((s) => s.kind === 'closing')
      if (ci >= 0) out.splice(ci, 0, ...annex)
      else out.push(...annex)
    }
  }
  return out
}

// Imagen de muestra para la demo (canvas, solo en navegador). Genera un
// gráfico de barras sencillo para que el slide figure no salga vacío.
function sampleFigureImg() {
  if (typeof document === 'undefined') return null
  const cv = document.createElement('canvas')
  cv.width = 480
  cv.height = 300
  const ctx = cv.getContext('2d')
  ctx.fillStyle = '#EEF3EE'
  ctx.fillRect(0, 0, 480, 300)
  const cols = ['#2E7D4F', '#363EA1', '#A13674', '#A19836']
  const vals = [0.55, 0.78, 0.62, 0.9]
  vals.forEach((v, i) => {
    const h = Math.round(v * 210)
    ctx.fillStyle = cols[i % cols.length]
    ctx.fillRect(45 + i * 105, 250 - h, 66, h)
  })
  ctx.fillStyle = '#2A3B32'
  ctx.font = '15px sans-serif'
  ctx.fillText('Muestra: reparto por escenario', 45, 282)
  return { data: cv.toDataURL('image/png'), w: 480, h: 300 }
}

// SAMPLE_DECK: 10 slides, una por kind, para demo sin IA
const SAMPLE_DECK = [
  { kind: 'cover', title: 'Diseño Sostenible', subtitle: 'Economía circular y evaluación de la sostenibilidad' },
  { kind: 'stats', title: 'Contexto clave', items: [
    { value: '1,7', label: 'planetas', color: 'accent5' },
    { value: '>60 %', label: 'de la huella es carbono', color: 'accent1' }
  ], takeaway: 'Vivimos por encima de la biocapacidad' },
  { kind: 'process', title: 'Fases del ACV', steps: [
    ['Objetivos y alcance', 'Límites del sistema y unidad funcional'],
    ['Inventario', 'Entradas y salidas: energía, agua, residuos'],
    ['Evaluación', 'Midpoint y endpoint'],
    ['Interpretación', 'Puntos críticos y recomendaciones']
  ], takeaway: 'El ACV integra 4 fases normalizadas' },
  { kind: 'table', title: 'Comparativa de indicadores', table: { header: ['Indicador', 'Qué mide', 'Unidad', 'Alcance'], rows: [
    ['Huella ecológica', 'Superficie para sostener estilo de vida', 'hag/hab', 'Global'],
    ['ACV', 'Múltiples impactos de un producto', 'kg CO₂ eq, MJ', 'Producto'],
    ['Huella de carbono', 'Emisiones GEI', 'kg CO₂ eq', 'Actividad']
  ] }, formula: 'LCSA = LCA + LCC + S-LCA' },
  { kind: 'formula', title: 'Huella de carbono', formula: 'Emisiones = Actividad × Factor de emisión', columns: [
    ['Alcance 1', 'Directas: combustión, procesos, fugas'],
    ['Alcance 2', 'Energía adquirida: electricidad, calor'],
    ['Alcance 3', '>70 %: cadena de valor, uso de productos']
  ] },
  { kind: 'compare', title: 'Lineal vs Circular', items: [
    { value: 'Tomar → Hacer → Tirar', label: 'Lineal' },
    { value: 'Diseñar sin residuos', label: 'Circular 1' },
    { value: 'Mantener en uso', label: 'Circular 2' }
  ] },
  { kind: 'content', title: 'Ecodiseño LiDS', bullets: [
    'Conceptos innovadores: desmaterialización',
    'Componentes de bajo impacto',
    'Estructura para desmontaje',
    'Sistema de producto circular'
  ], notes: 'Apoyos: ISO 14006, Reglamento ESPR' },
  { kind: 'chart', title: 'Producción renovable', chart: {
    type: 'bar',
    axisTitle: 'Generación (TWh)',
    categories: ['2020', '2021', '2022', '2023'],
    series: [
      { name: 'Eólica (TWh)', values: [220, 235, 245, 260] },
      { name: 'Solar (TWh)', values: [140, 160, 180, 205] }
    ]
  }, takeaway: 'La eólica mantiene el liderazgo' },
  { kind: 'figure', title: 'Reparto por escenario', figure: 1, takeaway: 'El escenario con baterías reduce el pico' },
  { kind: 'closing', title: 'Ideas clave', closing: [
    'Medir para decidir: huella, ACV y carbono se complementan',
    'Circular: mantener materiales en uso con energía renovable',
    'Ecodiseñar desde el inicio: corregir después es caro'
  ] }
]

if (typeof globalThis !== 'undefined') {
  globalThis.__generador = {
    buildPrompt,
    extractJson,
    normalizeSlides,
    structureSlides,
    autoFont,
    chartPalette,
    assignFigures,
    finishImageList,
    stripMarkers,
    extractDocx,
    extractPdf,
    sampleFigureImg,
    truncateText,
    orderModels,
    SAMPLE_DECK,
    splitKeywords,
    paginateSlides,
    render_cover,
    render_stats,
    render_process,
    render_table,
    render_formula,
    render_compare,
    render_chart,
    render_figure,
    render_content,
    render_closing,
    buildPptx
  }
}

const KEY_NAME = 'pptxgen.key.v1'
const PROVIDERS = {
  aistudio: { label: 'Google AI Studio', model: 'gemini-3.8-flash', models: ['gemini-3.8-flash', 'gemini-3.7-flash', 'gemini-3.6-flash', 'gemini-3.5-flash', 'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite'], keyUrl: 'https://aistudio.google.com/apikey' },
  openrouter: { label: 'OpenRouter', model: 'meta-llama/llama-3.3-70b-instruct:free', models: ['meta-llama/llama-3.3-70b-instruct:free'], keyUrl: 'https://openrouter.ai/keys' },
}
let lastSlides = null

function loadKey() {
  try { return localStorage.getItem(KEY_NAME) || '' }
  catch (e) { return '' }
}
function saveKey(k) {
  try { localStorage.setItem(KEY_NAME, k) }
  catch (e) {}
}
function clearKey() {
  try { localStorage.removeItem(KEY_NAME) }
  catch (e) {}
}
function authError(provider) {
  const p = PROVIDERS[provider] || PROVIDERS.aistudio
  return new Error(`Clave inválida o sin acceso (${p.label}). Revisa la consola del proveedor: ${p.keyUrl}`)
}
function providerError(provider, model, status, body) {
  if (status === 401 || status === 403 || /api[\s_-]?key|unauthorized|unauthenticated/i.test(body || ''))
    return authError(provider)
  return new Error(`Error del proveedor (HTTP ${status}). Revisa el modelo usado: ${model}.`)
}

const REQUEST_TIMEOUT_MS = 60000

function fetchWithTimeout(url, options, ms) {
  const ctrl = new AbortController()
  const t = setTimeout(() => ctrl.abort(), ms)
  const opts = Object.assign({}, options, { signal: ctrl.signal })
  return fetch(url, opts).then(
    (resp) => { clearTimeout(t); return resp },
    (e) => {
      clearTimeout(t)
      if (e && e.name === 'AbortError')
        throw new Error(`Tiempo agotado (${Math.round(ms / 1000)} s) esperando al modelo.`)
      throw e
    },
  )
}

function callAistudio(apiKey, model, prompt, ms, images) {
  // la clave viaja en la cabecera x-goog-api-key, no en la URL, para no dejarla en
  // registros de proxies/extensiones que inspeccionan la linea de peticion.
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)
  }:generateContent`
  const parts = [{ text: prompt }]
  ;(Array.isArray(images) ? images : []).forEach((im) => {
    const m = /^data:([^;]+);base64,(.*)$/.exec(im.data || '')
    if (m) parts.push({ inlineData: { mimeType: m[1], data: m[2] } })
  })
  return fetchWithTimeout(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
    body: JSON.stringify({
      contents: [{ parts }],
      generationConfig: { responseMimeType: 'application/json' },
    }),
  }, ms).then((resp) => {
    if (!resp.ok)
      return resp.text().then((b) => { throw providerError('aistudio', model, resp.status, b) })
    return resp.json()
  }).then((data) => {
    const c = ((data.candidates || [])[0] || {})
    const parts = (c.content || {}).parts || []
    return parts.map((p) => { return p.text || '' }).join('')
  })
}

function callOpenrouter(apiKey, model, prompt, ms, images) {
  // OpenRouter con llama-3.3 solo texto: las imágenes se ignoran (modo ciego).
  return fetchWithTimeout('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
      // origen+path (sin query ni fragmento): atribucion sin filtrar la URL completa.
      'HTTP-Referer': location.origin + location.pathname,
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: 'system', content: 'Devuelves solo JSON válido.' },
        { role: 'user', content: prompt },
      ],
    }),
  }, ms).then((resp) => {
    if (!resp.ok)
      return resp.text().then((b) => { throw providerError('openrouter', model, resp.status, b) })
    return resp.json()
  }).then((data) => {
    const ch = ((data.choices || [])[0] || {})
    return (ch.message || {}).content || ''
  })
}

function orderModels(models, chosen) {
  const list = Array.isArray(models) ? models.slice() : []
  const out = []
  if (chosen)
    out.push(chosen)
  list.forEach((m) => { if (m !== chosen) out.push(m) })
  return out
}

function sleep(ms) {
  return new Promise((resolve) => { setTimeout(resolve, ms) })
}

function structureSlides(text, opts) {
  const o = opts || {}
  const provider = o.provider || 'aistudio'
  const p = PROVIDERS[provider] || PROVIDERS.aistudio
  const maxSlides = Math.max(1, Math.min(o.maxSlides || DEFAULT_MAX_SLIDES, 30))
  const waitMs = (typeof o.waitMs === 'number') ? o.waitMs : 3000
  const timeoutMs = (typeof o.timeoutMs === 'number') ? o.timeoutMs : REQUEST_TIMEOUT_MS
  const t = truncateText(text)
  const notice = t.truncated ? ` (nota: texto recortado a ${MAX_CHARS} caracteres)` : ''
  const images = Array.isArray(o.images) ? o.images : []
  // Visión: solo Gemini (AI Studio) puede ver las imágenes adjuntas; el resto → modo ciego.
  const wantVision = provider === 'aistudio' && images.length > 0
  const prompt = buildPrompt(t.text, maxSlides, o.instrucciones, { imageCount: images.length, vision: wantVision })
  const blindPrompt = buildPrompt(t.text, maxSlides, o.instrucciones, { imageCount: images.length, vision: false })
  const call = provider === 'openrouter' ? callOpenrouter : callAistudio
  const chain = provider === 'aistudio' ? orderModels(p.models, o.model || p.model) : [(o.model || p.model)]
  const ask = (model, pr, useImages) => call(o.apiKey || '', model, pr, timeoutMs, useImages ? images : [])
    .then((raw) => assignFigures(normalizeSlides(extractJson(raw), maxSlides), images, useImages ? 'vision' : 'blind'))
  // Si la IA entrega menos diapositivas de las pedidas, se repite (máx. 2 veces)
  // con un correctivo antes de aceptar el resultado.
  const complete = (model, slides, attempt, useImages) => {
    if (slides.length >= maxSlides || attempt > 2)
      return slides
    setStatus('estructurando', `La IA entregó ${slides.length} de ${maxSlides} diapositivas; pidiendo de nuevo…`)
    const corrective = prompt
      + `\n\nIMPORTANTE: tu respuesta anterior solo contenía ${slides.length} diapositivas válidas. `
      + `Devuelve EXACTAMENTE ${maxSlides} diapositivas en total, contando portada y cierre, `
      + `y ninguna con bullets, items, steps o rows vacíos.`
    return ask(model, corrective, useImages).then((s2) => complete(model, s2, attempt + 1, useImages))
  }
  const tryModel = (model) => {
    setStatus('estructurando', `Probando ${model} con ${p.label}${wantVision ? ' (viendo las imágenes)' : ''}…${notice}`)
    return ask(model, prompt, wantVision).then((slides) => complete(model, slides, 1, wantVision))
      .catch((e) => {
        // Si la llamada con imágenes fue rechazada, reintenta en modo ciego (solo marcadores).
        if (wantVision && !/^Clave inválida/.test(e.message)) {
          setStatus('estructurando', `${model} no admitió las imágenes; reintentando en modo ciego…`)
          return ask(model, blindPrompt, false).then((s2) => complete(model, s2, 1, false))
        }
        throw e
      })
  }
  const runChain = (i) => {
    if (i >= chain.length) {
      const msg = 'Ningún modelo respondió. Revisa tu clave y tu conexión.'
      setStatus('error', msg)
      throw new Error(msg)
    }
    return tryModel(chain[i]).catch((e) => {
      if (/^Clave inválida/.test(e.message)) {
        setStatus('error', e.message)
        throw e
      }
      setStatus('estructurando', `${chain[i]} falló. Reintentando en ${Math.round(waitMs / 1000)} s…`)
      return sleep(waitMs).then(() => tryModel(chain[i])).catch(() => {
        return runChain(i + 1)
      })
    })
  }
  setStatus('estructurando', `Generando estructura con ${p.label}…${notice}`)
  return runChain(0)
}

function currentOpts() {
  return {
    provider: document.getElementById('sel-provider').value,
    apiKey: document.getElementById('in-key').value.trim(),
    maxSlides: parseInt(document.getElementById('in-max').value, 10) || DEFAULT_MAX_SLIDES,
    model: document.getElementById('in-model').value.trim(),
    instrucciones: document.getElementById('in-instrucciones').value.trim(),
  }
}

function onGenerate() {
  const o = currentOpts()
  if (!o.apiKey) {
    setStatus('error', 'Falta la API key. Pégala arriba (queda solo en tu navegador).')
    return
  }
  saveKey(o.apiKey)
  setStatus('leyendo', 'Leyendo apuntes…')
  getInputText().then((input) => {
    return structureSlides(input.text, { ...o, images: input.images })
  }).then((slides) => {
    lastSlides = slides
    const falta = slides.length < o.maxSlides ? ` (pediste ${o.maxSlides}: la IA no dio más contenido)` : ''
    setStatus('listo', `Estructura lista: ${slides.length} diapositivas${falta}. Ya puedes descargar el .pptx.`)
    document.getElementById('btn-download').disabled = false
  }).catch((e) => {
    if (document.getElementById('out-status').getAttribute('data-phase') !== 'error') {
      setStatus('error', e.message)
    }
  })
}

const SAMPLE_SLIDES = [
  { title: 'Energía eólica', bullets: [] },
  { title: 'Ventajas', bullets: ['Sin emisiones durante la operación', 'Coste nivelado a la baja', 'Escalable de kW a GW'] },
  { title: 'Datos 2024', bullets: [], table: { header: ['Fuente', 'TWh'], rows: [['Eólica', '245'], ['Solar', '180'], ['Hidro', '320']] } },
]

function chunk(arr, n) {
  const out = []
  for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n))
  return out
}

function buildPptx(slides, meta) {
  const m = meta || {}
  const th = THEMES[m.theme] || THEMES.seriph
  const pptx = new PptxGenJS({
    theme: {
      headFontFace: th.fontTitle,
      bodyFontFace: th.fontBody,
      colors: {
        dk1: th.ink,
        lt1: th.bg,
        dk2: th.accent,
        lt2: th.cardBg,
        accent1: th.accent,
        accent2: th.accent2 || th.accent,
        accent3: th.accent3 || th.accent,
        accent4: th.accent4 || th.accent,
        accent5: th.accent5 || th.accent,
        accent6: th.accent6 || th.accent,
        hlink: th.accent3 || th.accent,
        folHlink: th.accent6 || th.accent
      }
    }
  })
  pptx.author = 'Generador PPTX'
  pptx.layout = 'LAYOUT_16x9'
  pptx.title = m.title || 'Presentación'
  const pages = paginateSlides(slides)
  const total = pages.length
  let sec = 0
  pages.forEach((pg, i) => {
    const slide = pptx.addSlide()
    slide.background = { color: pg.cover ? th.coverBg : th.bg }
    if (!pg.cover) sec += 1
    const title = pg.cover ? pg.title : sec + '. ' + pg.title
    if (pg.cover) {
      render_cover(slide, pg, th)
    } else if (pg.closing) {
      render_closing(slide, { title, items: pg.closing }, th)
    } else if (pg.stats) {
      render_stats(slide, { title, items: pg.stats, takeaway: pg.takeaway }, th)
    } else if (pg.process) {
      render_process(slide, { title, steps: pg.process, takeaway: pg.takeaway }, th)
    } else if (pg.table) {
      render_table(slide, { title, headers: pg.table.header, rows: pg.table.rows, formula: pg.formula }, th)
    } else if (pg.formula) {
      render_formula(slide, { title, formula: pg.formula, columns: pg.columns }, th)
    } else if (pg.compare) {
      render_compare(slide, { title, items: pg.compare }, th)
    } else if (pg.chart) {
      render_chart(slide, { title, chart: pg.chart, takeaway: pg.takeaway }, th)
    } else if (pg.figure !== undefined || pg.img !== undefined) {
      render_figure(slide, { title, img: pg.img, takeaway: pg.takeaway }, th)
    } else {
      render_content(slide, { title, bullets: pg.bullets }, th)
    }
    // Footer N/total
    slide.addText(`${i + 1}/${total}`, {
      x: 8.7, y: 5.25, w: 0.9, h: 0.25,
      fontFace: th.fontBody, fontSize: 9, color: th.footer, align: 'right'
    })
  })
  return pptx.write({ outputType: 'blob' })
}

function currentTheme() {
  const v = document.getElementById('sel-theme').value
  return THEMES[v] ? v : 'seriph'
}

// ===== Barra lateral: previsualización de estilos =====
const PREVIEW_BULLETS = [
  ['Contexto: ', 'energía', ' y recursos del tema'],
  ['La ', 'huella de carbono', ' como indicador clave'],
  ['Estrategias de ', 'ecodiseño', ' aplicadas'],
  ['Resultados y ', 'conclusiones', ' del análisis'],
]

function styleCard(k) {
  const th = THEMES[k]
  const btn = document.createElement('button')
  btn.type = 'button'
  btn.className = 'style-card'
  btn.setAttribute('data-theme', k)
  const vars = {
    bg: th.bg, ink: th.ink, accent: th.accent, a2: th.accent2, a4: th.accent4,
    card: th.cardBg, cardt: th.cardTitle, coverbg: th.coverBg,
    covert: th.coverTitle, covers: th.coverSub,
  }
  Object.keys(vars).forEach((key) => { btn.style.setProperty('--' + key, '#' + vars[key]) })
  btn.style.setProperty('--ft', "'" + th.fontTitle + "'")
  btn.style.setProperty('--fb', "'" + th.fontBody + "'")

  const name = document.createElement('span')
  name.className = 'style-name'
  name.textContent = th.name
  btn.appendChild(name)

  // Miniatura de portada (mismas proporciones que render_cover)
  const cover = document.createElement('div')
  cover.className = 'mini mini-cover'
  cover.style.background = 'var(--coverbg)'
  const blob1 = document.createElement('div'); blob1.className = 'm-blob'
  const blob2 = document.createElement('div'); blob2.className = 'm-blob m-blob2'
  const ct = document.createElement('div'); ct.className = 'm-title'; ct.textContent = 'Título del tema'
  const cs = document.createElement('div'); cs.className = 'm-sub'; cs.textContent = 'Subtítulo del tema'
  cover.append(blob1, blob2, ct, cs)
  btn.appendChild(cover)

  // Miniatura de contenido (título + barra de acento + bullets con keyword)
  const cont = document.createElement('div')
  cont.className = 'mini mini-content'
  cont.style.background = 'var(--bg)'
  const ht = document.createElement('div'); ht.className = 'm-h'; ht.textContent = '1. Encabezado de la slide'
  const bar = document.createElement('div'); bar.className = 'm-bar'
  const ul = document.createElement('div'); ul.className = 'm-bullets'
  PREVIEW_BULLETS.forEach((b) => {
    const li = document.createElement('div'); li.className = 'm-b'
    const dot = document.createElement('span'); dot.className = 'm-dot'
    li.append(dot, document.createTextNode(b[0]))
    const strong = document.createElement('b'); strong.textContent = b[1]
    li.append(strong, document.createTextNode(b[2]))
    ul.appendChild(li)
  })
  cont.append(ht, bar, ul)
  btn.appendChild(cont)

  btn.addEventListener('click', () => selectTheme(k))
  return btn
}

function syncStyleCards() {
  const cur = document.getElementById('sel-theme').value
  const cards = document.querySelectorAll('#style-previews .style-card')
  Array.prototype.forEach.call(cards, (b) => {
    const on = b.getAttribute('data-theme') === cur
    b.classList.toggle('selected', on)
    b.setAttribute('aria-pressed', on ? 'true' : 'false')
  })
}

function selectTheme(k) {
  const sel = document.getElementById('sel-theme')
  if (sel && sel.value !== k) {
    sel.value = k
    sel.dispatchEvent(new Event('change'))
  }
  syncStyleCards()
}

function renderStylePreviews() {
  const box = document.getElementById('style-previews')
  if (!box) return
  box.innerHTML = ''
  Object.keys(THEMES).forEach((k) => { box.appendChild(styleCard(k)) })
  syncStyleCards()
}

function downloadPptx(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  setTimeout(() => { URL.revokeObjectURL(url); a.remove() }, 1000)
}

function init() {
  const missing = [
    ['PptxGenJS', 'pptxgenjs'],
    ['mammoth', 'mammoth'],
    ['pdfjsLib', 'pdf.js'],
  ].filter((pair) => { return typeof window[pair[0]] === 'undefined' })
  const sel = document.getElementById('sel-provider')
  const key = document.getElementById('in-key')
  const model = document.getElementById('in-model')
  const buildModelSelect = function () {
    const prov = PROVIDERS[sel.value] || PROVIDERS.aistudio
    const plist = Array.isArray(prov.models) && prov.models.length ? prov.models : [prov.model]
    model.innerHTML = ''
    plist.forEach((m) => {
      const opt = document.createElement('option')
      opt.value = m
      opt.textContent = m
      model.appendChild(opt)
    })
    model.value = plist[0]
  }
  key.value = loadKey()
  buildModelSelect()
  sel.addEventListener('change', buildModelSelect)
  const themeSel = document.getElementById('sel-theme')
  Object.keys(THEMES).forEach((k) => {
    const opt = document.createElement('option')
    opt.value = k
    opt.textContent = THEMES[k].name
    themeSel.appendChild(opt)
  })
  themeSel.value = 'seriph'
  renderStylePreviews()
  themeSel.addEventListener('change', syncStyleCards)
  const estilosBtn = document.getElementById('btn-estilos')
  estilosBtn.addEventListener('click', () => {
    const open = document.getElementById('estilos').classList.toggle('collapsed')
    estilosBtn.setAttribute('aria-expanded', open ? 'false' : 'true')
  })
  key.addEventListener('change', () => { saveKey(key.value.trim()) })
  document.getElementById('btn-key-del').addEventListener('click', () => {
    key.value = ''
    clearKey()
    setStatus('listo', 'Clave borrada de este navegador.')
  })
  document.getElementById('btn-generate').disabled = false
  document.getElementById('btn-generate').addEventListener('click', onGenerate)
  document.getElementById('btn-download').disabled = true
  document.getElementById('btn-download').addEventListener('click', () => {
    if (!lastSlides)
      return
    if (typeof PptxGenJS === 'undefined') {
      setStatus('error', 'No se pudo cargar pptxgenjs (revisa tu conexión).')
      return
    }
    setStatus('armando', 'Armando PPTX…')
    buildPptx(lastSlides, { title: lastSlides[0].title, theme: currentTheme() }).then((blob) => {
      downloadPptx(blob, 'presentacion.pptx')
      setStatus('listo', 'PPTX descargado.')
    }).catch((e) => {
      setStatus('error', e.message)
    })
  })
  document.getElementById('btn-demo').addEventListener('click', () => {
    if (typeof PptxGenJS === 'undefined') {
      setStatus('error', 'No se pudo cargar pptxgenjs (revisa tu conexión).')
      return
    }
    SAMPLE_DECK.forEach((sl) => { if (sl.kind === 'figure' && !sl.img) sl.img = sampleFigureImg() })
    lastSlides = SAMPLE_DECK
    const themeKey = currentTheme()
    const th = THEMES[themeKey] || THEMES.seriph
    const pages = paginateSlides(SAMPLE_DECK)
    const inst = document.getElementById('in-instrucciones').value.trim()
    const extra = inst ? ' — Instrucciones detectadas: la demo usa un mazo fijo y no puede aplicarlas; se aplicarán al generar con tu clave.' : ''
    setStatus('listo', `Ejemplo cargado: ${pages.length} diapositivas (${THEMES[themeKey].name}). Pulsa Descargar .pptx.${extra}`)
    document.getElementById('btn-download').disabled = false
  })
  const themeBtn = document.getElementById('btn-theme')
  const applyTheme = function (dark) {
    document.body.classList.toggle('dark', dark)
    themeBtn.textContent = dark ? '☀️' : '🌙'
    try { localStorage.setItem('pptxgen.theme', dark ? 'dark' : 'light') } catch (e) {}
  }
  let dark = false
  try { dark = localStorage.getItem('pptxgen.theme') === 'dark' } catch (e) {}
  applyTheme(dark)
  themeBtn.addEventListener('click', () => { applyTheme(!document.body.classList.contains('dark')) })
  if (missing.length) {
    const names = missing.map((pair) => { return pair[1] }).join(', ')
    document.getElementById('btn-generate').disabled = true
    document.getElementById('btn-demo').disabled = true
    setStatus('error', `No se pudo cargar ${names} (revisa tu conexión).`)
    return
  }
  setStatus('listo', 'Listo.')
}

if (typeof window !== 'undefined') {
  init()
}
