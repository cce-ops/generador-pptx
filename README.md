# Generador PPTX (web estática)

Convierte apuntes en una presentación `.pptx` **editable en PowerPoint**, 100 % en el
navegador. Sin backend, sin cuentas, sin telemetría y **gratis de alojar**.

https://cce-ops.github.io/generador-pptx/

Pega texto, o sube un `.md` / `.docx` / `.pdf`, elige un estilo y descarga una
presentación 16:9 con portada, tablas, gráficos nativos y las propias figuras del
documento. Usas **tu propia clave de API** (BYOK): no pasa nada por un servidor intermedio.

**Probar sin clave:** el botón **Probar sin clave** genera una presentación de ejemplo
con los 10 tipos de slide, sin llamar a ninguna IA.

## Qué genera

- **16:9 (10" × 5.625")**, con fuentes ajustadas para que nada se desborde.
- **10 tipos de slide:** `cover`, `stats`, `process`, `table`, `formula`, `compare`,
  `content`, `closing`, `chart`, `figure`.
- **Títulos numerados** ("1. Contexto clave") excepto portada y cierre.
- **Títulos y bullets** con tamaño automático para llenar la slide sin solaparse ni desbordar.
- **Tablas** con cabecera de color y el pie `N/total` en cada slide.

### Gráficos nativos (no imágenes)

El tipo `chart` genera un **gráfico real e editable de PowerPoint** (puedes tocar los
datos allí). Cada variable lleva su **color propio**, con **nombre de eje vertical y
unidades**. Soporta barras, líneas, tarta y doughnut.

### Figuras del documento (`.docx` y `.pdf`)

Sube un documento con figuras: se **extraen las imágenes** (mín. 100×100 px, sin logos
repetidos, tope de 20), se insertan marcadores `[IMAGEN n]` junto al texto y la IA
decide cuáles se convierten en slides `figure`. Con **Gemini (visión)** la IA *ve* las
imágenes y escribe un pie real para cada una; con modelos solo-texto las asigna por
orden. Si el documento no tiene figuras, esos slides se rellenan como `content`.

## Estilos de presentación

Seis estilos, previsualizables en la barra lateral izquierda (plegable):

- **Seriph claro** — fondo blanco, serif, acento azul suave (por defecto).
- **Verde sostenible** — fondo blanco, acentos verdes.
- **Oscuro** — fondo oscuro, texto claro, acento menta.
- **Azul corporativo** — fondo blanco, azul marino.
- **Blanco y azul** — fondo blanco, azul `#4285F4`.
- **Verde skill** — bosque / lima / ámbar.

El estilo elegido se aplica igual a la generación con IA y al modo demo.

## Privacidad y modelos

- La **API key queda solo en el `localStorage` de tu navegador** (botón *Borrar clave* la
  elimina). No hay servidor propio: el navegador habla directo con el proveedor.
- **Google AI Studio (Gemini)** es el proveedor por defecto, con varios modelos y cadena
  de reintentos automática si uno falla.
- **OpenRouter** (modelos gratuitos) es la alternativa.
- Claves gratuitas: [AI Studio](https://aistudio.google.com/apikey) ·
  [OpenRouter](https://openrouter.ai/keys). Puedes restringir la clave por dominio.

## Desarrollo local

Sin build ni dependencias que instalar: solo un servidor estático sobre la carpeta.

```bash
npx serve .          # o: python -m http.server 8137
# abre http://localhost:8137
```

Tests (unit, sin navegador):

```bash
node --test tests/
```

## Despliegue gratis (GitHub Pages)

No hay paso de build ni variables de entorno: se publica tal cual.

1. **Settings → Pages → Source: Deploy from a branch**.
2. Rama `main`, carpeta `/ (root)`.
3. En unos segundos queda en `https://cce-ops.github.io/generador-pptx/`.

El fichero `.nojekyll` evita que GitHub Pages procese la carpeta con Jekyll.

También funciona en **Cloudflare Pages** (arrastra la carpeta o conecta el repo;
comando de build vacío).

## Notas técnicas

Librerías fijadas vía CDN jsDelivr (sin build), con **SRI** (`integrity` +
`crossorigin`) para que un fallo del CDN no ejecute código ajeno:

- `pptxgenjs@4.0.1` — genera el `.pptx`.
- `mammoth@1.13.0` — extrae texto de `.docx`.
- `pdfjs-dist@3.11.174` — extrae texto e imágenes de `.pdf` (UMD; la 4.x es solo ESM).

El worker de pdf.js (`pdf.worker.js`) va **autoalojado en el repo** (mismo origen):
así no depende del CDN y pdf.js usa su worker real. Además `getDocument` se llama
con `isEvalSupported: false` para cerrar el CVE-2024-4367 (el parser de fuentes de
pdf.js era vulnerable a ejecución de código con un PDF malicioso).

La IA devuelve una estructura JSON (`kind`, `title`, `bullets`…). Un normalizador
garantiza que el número de slides pedido se cumpla, que cada slide tenga título y que
los campos estén saneados antes de renderizar.
