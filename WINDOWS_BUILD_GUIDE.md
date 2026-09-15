# TETORD 4.2.0 — Build de Windows

## La forma facil

1. Extrae este ZIP completo en una carpeta.
2. Asegurate de tener **Node.js LTS** instalado en Windows.
3. Ejecuta `BUILD_TETORD_WINDOWS.bat` con doble clic.
4. El script instalara las dependencias y ejecutara `npm run build:release`.
5. Al terminar, abre la carpeta `dist`.

## Archivos esperados

- `TETORD-Setup-4.2.0-x64.exe` — instalador de Windows.
- `TETORD-4.2.0-x64.exe` — version portable.

## Si aparece un error

No cierres la ventana. Copia el mensaje de error o manda una captura y revisamos exactamente ese punto.

## Nota

El build de Windows debe ejecutarse en Windows. El paquete ya incluye la configuracion de release; este script solamente automatiza el proceso para evitar tener que escribir los comandos manualmente.
