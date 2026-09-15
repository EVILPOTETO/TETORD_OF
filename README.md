
## TETORD 2.7.1 — Imágenes y tablas PRO

- Herramientas contextuales funcionales para imágenes y tablas.
- Imágenes: ajuste de texto, rotación, recorte uniforme/por lados, posición/alineación y texto alternativo.
- Tablas: combinar/dividir celdas, alineación horizontal/vertical, color de celda, grosor de bordes, encabezado repetible al imprimir, filas/columnas y propiedades.
- Se conserva la edición HTML5 + CSS3 + JavaScript sin frameworks.

# Mi Word 1.1

Procesador de textos web propio construido con HTML5, CSS3 y JavaScript puro.

## Inicio
Al abrir Mi Word se muestra una pantalla inicial con:
- Documento en blanco
- Trabajo escolar
- Informe
- Carta
- Apuntes
- Currículum
- Documentos recientes guardados localmente
- Abrir documento

## Privacidad
Los documentos recientes y documentos locales se almacenan en el almacenamiento del navegador del dispositivo. No se usa un servidor externo para sincronización.

## Compatibilidad
Se recomienda un navegador moderno con soporte para `contenteditable`, `localStorage`, FileReader y, cuando esté disponible, File System Access API.

## Pruebas
La validación estática de la distribución pasa. La prueba interactiva completa en navegador no se realizó en el entorno de desarrollo porque Chromium está bloqueado.


## TETORD 2.0
TETORD (Teto + Word) añade herramientas de párrafo, ecuaciones Unicode, símbolos, notas al pie, tabla de contenido, marcadores, combinación/división de celdas, propiedades del documento y una barra de herramientas ampliada. Mantiene HTML5/CSS3/JavaScript vanilla y almacenamiento local.


## TETORD 4.3.0 — Slides PRO editor refresh

Slides recibe una interfaz de edición profesional con lienzo 16:9, miniaturas visuales, objetos posicionables, panel de propiedades, inserción de imágenes y modo presentación. Writer, Calc, Messenger y el resto de la suite se conservan como base existente.

## TETORD 4.2.0 — Final Build & Release Verification

Esta versión prepara el paquete de lanzamiento para Windows y conserva el enfoque offline-first.

### Build recomendado

```bash
npm install
npm run build:release
```

`build:release` ejecuta `preflight`, limpia `dist/`, verifica la allowlist de producción, genera los objetivos **NSIS** y **Portable** x64 y crea `dist/RELEASE_MANIFEST.json`. El instalador no incluye `server/`, `scripts/`, pruebas, `.git` ni otros artefactos de desarrollo.

### Archivos esperados

- `dist/TETORD-Setup-4.2.0-x64.exe` — instalador NSIS.
- `dist/TETORD-4.2.0-x64.exe` — versión Portable.

### Smoke test de Windows

Consulta `RELEASE_CHECKLIST.md` y ejecuta el checklist completo en una máquina Windows limpia antes de distribuir el instalador.

### Estado de verificación

El código y los tests estáticos pueden validarse en el entorno de desarrollo. La ejecución interactiva del instalador, persistencia real en Windows, permisos de `Documents/TETORD` y colaboración WebSocket entre clientes requieren una máquina Windows/runtime real.

## TETORD Desktop

Esta distribución incluye un empaquetado de escritorio con Electron. En Windows puedes ejecutar `npm install` y después `npm run dist` para generar el instalador y la versión portable.


## Versión actual

TETORD 4.3.0 — Final Build & Release Edition.


## 2.4.0
La edición 2.2 incorpora una cinta de opciones organizada por pestañas y un sistema de estilos basado en datos, manteniendo la arquitectura HTML5/CSS3/JavaScript puro.


### 2.4.0
Incluye controles contextuales para imágenes y tablas.

## 2.5.0 — Revisar
Añade comentarios anclados al texto, control de cambios (insertar/borrar
marcado, aceptar/rechazar), sinónimos con diccionario local, comparación
de documentos por texto plano y protección de documento a nivel de
aplicación. Ver `CHANGELOG.md` y `MiWord/PROJECT_STATE.md` para el
detalle y las limitaciones conocidas.


## TETORD 2.6.0 — Referencias

Incluye tabla de contenido actualizable, notas al final, fuentes y citas, bibliografía, títulos de figuras/tablas, referencias cruzadas, marcadores de índice e índice generado. Los datos de fuentes se almacenan localmente y las referencias se conservan dentro del HTML del documento.


## 2.8.0 — Paginación e impresión PRO
Incluye vista previa de impresión por hojas, saltos de página, configuración A4/Carta y orientación, márgenes, columnas, guionización y encabezados/pies con numeración.


## TETORD 2.9.0 — Correspondencia
La fase 2.9 incorpora combinación de correspondencia local: destinatarios desde CSV/TSV, campos combinados, vista previa por registro y generación de un documento HTML combinado. No requiere servidor.

## TETORD Office 3.0
TETORD 3.0 incorpora Writer, Calc, Slides y Messenger en una suite local. Messenger no usa servidor todavía.


## TETORD 3.1.0 — Writer PRO

Writer incorpora historial local de versiones, saltos de sección, pegado sin formato, estadísticas de selección y atajos de productividad.

## 3.5.1
Calc PRO incorpora múltiples hojas, fórmulas adicionales y formato de celdas. La identidad visual de TETORD Office diferencia Writer (rojo pastel), Calc (amarillo pastel) y Slides (azul pastel).


### 3.5.1 — TETORD Files
La suite incorpora un espacio común de **Archivos** para localizar y abrir recursos locales de Writer, Calc y Slides. Incluye búsqueda, filtros, recientes y accesos rápidos para crear contenido. Sigue siendo almacenamiento local; la nube, colaboración y sincronización quedan para fases posteriores.


## Messenger ONLINE

TETORD 3.8.0 incluye un backend WebSocket independiente en `server/`. El escritorio puede conectarse a `ws://127.0.0.1:8787`; si el servidor no está disponible, Messenger mantiene automáticamente el adaptador local.

### Servidor local

```bash
cd server
npm install
npm start
```

Opcionalmente define `TETORD_AUTH_TOKEN` para exigir un token compartido. Por defecto, el servidor local acepta el nick sin token para facilitar desarrollo. Los mensajes se persisten en `server/data/messenger.json`.

### Escritorio

```bash
npm start
npm run dist
```

El servidor no se ejecuta dentro del instalador: es un proceso independiente que puede desplegarse en otra máquina y exponerse como `ws://host:puerto` o, en producción, `wss://host:puerto`.
