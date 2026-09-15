# TETORD 4.3.1 — Slides Executive Layouts

- Slides: layouts for title/content, big stat, 3 cards, timeline, quote, comparison.
- Presentation editor keeps existing Writer/Calc/Messenger architecture intact.

# TETORD 4.3.0 — Slides PRO editor refresh

- Rediseño del editor de Slides con lienzo 16:9 real y espacio de trabajo de tres zonas.
- Miniaturas visuales de diapositivas y navegación más clara.
- Objetos editables de texto, formas e imágenes con selección, arrastre y panel de propiedades.
- Duplicación y eliminación de objetos, además de duplicación/eliminación de diapositivas.
- Inserción de imágenes persistentes en almacenamiento nativo.
- Modo presentación renderiza los objetos del lienzo respetando posición, tamaño y estilo.
- Compatibilidad de migración con presentaciones legacy basadas en título/cuerpo/imagen.

## 4.2.0 — Final Build & Release Verification
- Añadido `scripts/build-release.js` como gate de lanzamiento.
- `build:release` ejecuta `preflight`, limpia `dist/`, valida la allowlist de electron-builder y genera NSIS + Portable x64.
- Añadido `RELEASE_MANIFEST.json` con artefactos y estado del preflight.
- El servidor, scripts de QA y pruebas de concurrencia permanecen fuera del paquete Electron mediante una allowlist explícita.
- Añadido checklist de smoke tests para Windows y documentación de release.

# CHANGELOG

## 4.1.0 — Visual Polishing & Release Candidate Tests
- Added deterministic multi-client collaboration stress simulation.
- Added aggressive reconnect/retry coverage with operation deduplication.
- Hardened collaboration outbox, ACK tracking, and reconnect rejoin flow.
- Hardened Messenger room/DM deduplication across history and reconnect delivery.
- Added `npm run preflight` before distribution builds.

## TETORD 4.0.0 — Real-Time Collaboration & Platform Edition
- Autenticación de usuarios con JWT HS256 y contraseñas protegidas con scrypt.
- Presencia WebSocket por sala: online, ausente y offline.
- Rate limiting por IP, usuario y conexión WebSocket.
- Validación de payloads, límites de tamaño y sanitización de mensajes.
- API REST para registro/login y sincronización de documentos `.tetord`.
- Resolución básica de conflictos por timestamp de modificación.
- Nuevo módulo `MiWord/src/renderer/modules/cloud-sync.js` opcional y compatible con modo local/offline.
- Messenger WebSocket reutiliza el JWT guardado por Cloud Sync.

# TETORD 3.8.0 — Messenger ONLINE & Production Build

- Backend WebSocket independiente en `server/`.
- `WebSocketMessengerAdapter` con autenticación básica por nick/token.
- Reconexión automática y cola de mensajes pendientes.
- Fallback automático a `LocalMessengerAdapter`.
- Persistencia ligera de salas en JSON con escritura atómica.
- Documentación de servidor local y build de Windows.

# Changelog

## TETORD 3.7.0 — Files PRO & Real Document Engine

- Files PRO ahora inspecciona Documents/TETORD mediante IPC seguro.
- Writer, Calc y Slides usan documentos independientes `.tetord`.
- Metadatos reales de tamaño y modificación.
- CRUD nativo de archivos y carpetas.
- Messenger preparado con Adapter + LocalMessengerAdapter + WebSocketMessengerAdapter.
- Versionado centralizado en VERSION.

# TETORD 3.7.0 — Core Refactor & Native Storage

## 3.6.0
- Secure Electron IPC bridge for internal JSON and binary assets.
- Slides migrated from localStorage/Base64 persistence to disk JSON plus separate asset files.
- Safe DOM construction in Slides rendering and presentation mode.
- Added ES6 renderer module boundaries and AppModule base class.
- Removed legacy Python patch scripts.
- Added atomic Node.js version synchronization.

# TETORD 3.5.1 — Files PRO+

- Administrador local de archivos con carpetas, favoritos, búsqueda, filtros y ordenación.
- Acciones de renombrar, duplicar, mover y eliminar.
- Migas de navegación y persistencia local de metadatos.
- Integración con Writer, Calc y Slides conservada.
- Sin servidor: todas las operaciones siguen siendo locales.

## 3.4.0 — TETORD Files / Office Hub

- Nuevo módulo **Archivos** dentro de TETORD Office.
- Vista común de archivos locales de Writer, Calc y Slides.
- Búsqueda por nombre y filtros por aplicación.
- Acciones rápidas para crear Writer, Calc y Slides.
- Panel de archivos recientes en el inicio de Office.
- Apertura directa de documentos Writer guardados en el almacenamiento local de Mi Word.
- Metadatos de actividad local para Calc y Slides.
- Navegación unificada entre Inicio, Archivos, Writer, Calc, Slides y Messenger.
- Se mantiene el modelo local/offline: no hay nube ni sincronización de red en esta versión.


## 2.9.0 — Correspondencia
- Combinación de correspondencia local.
- Importación/pegado de CSV, TSV y texto delimitado.
- Destinatarios y campos dinámicos.
- Inserción de campos `{{Campo}}`.
- Vista previa por destinatario.
- Finalización en un HTML combinado con saltos de página.

## 2.8.0
- Paginación e impresión PRO.

## 2.7.1
- Barra contextual de imágenes y tablas corregida.

## 3.0.0 — TETORD Office
- Shell unificado con Writer, Calc, Slides y Messenger.
- Calc local: celdas, fórmulas básicas, SUM y CSV.
- Slides local: diapositivas, miniaturas y presentación.
- Messenger local persistente, preparado para backend futuro.


## 3.1.0 — Writer PRO
- Historial local de versiones por documento.
- Restauración de versiones con respaldo previo.
- Saltos de sección en página siguiente.
- Pegar texto sin formato.
- Estadísticas de texto seleccionado.
- Atajos de teclado para estilos, pegado sin formato, impresión y Guardar como.

## 3.3.0 — Calc PRO + identidad visual Office
- Calc evoluciona a módulo PRO: múltiples hojas, renombrado de hojas, formato de celdas, formatos numérico/moneda/porcentaje, alineación y fondo.
- Motor de fórmulas ampliado: SUM, AVERAGE, MIN, MAX, COUNT e IF básico.
- Importación CSV crea una hoja nueva; exportación usa la hoja activa.
- Identidad visual por aplicación: Writer rojo pastel (Teto), Calc amarillo pastel, Slides azul pastel.
- Persistencia local compatible con el formato Calc anterior.

## 3.3.0 — Slides PRO
- Slides evoluciona a editor con temas, fondos, duplicación e imágenes locales.
- Modo presentación navegable con teclado y botones.
- Identidad azul pastel oficial inspirada en Miku.
