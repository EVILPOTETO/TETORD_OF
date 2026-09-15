# TETORD Messenger Server

Backend WebSocket independiente de Electron.

## Requisitos
- Node.js 18+

## Instalación
```bash
npm install
```

## Inicio
```bash
npm start
```

Variables opcionales:
- `HOST` — por defecto `127.0.0.1`
- `PORT` — por defecto `8787`
- `TETORD_AUTH_TOKEN` — si se define, el cliente debe enviar el mismo token.
- `TETORD_SERVER_DATA` — directorio alternativo para `messenger.json`.

La persistencia se guarda en `data/messenger.json`. El servidor mantiene hasta 500 mensajes por sala y 100 mensajes de historial se entregan al conectar.
