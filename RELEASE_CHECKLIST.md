# TETORD 4.2.0 — Windows Release Smoke Test Checklist

Ejecutar después de `npm run build:release` en una máquina Windows limpia o VM. Marcar cada punto como PASS/FAIL y anotar evidencia.

## 0. Preparación

- [ ] Windows actualizado y usuario sin una instalación previa de TETORD, o instalación previa documentada.
- [ ] Conectividad a Internet disponible solo para las pruebas que realmente requieran Cloud/online.
- [ ] Guardar los dos artefactos de `dist/`: NSIS y Portable.
- [ ] Confirmar versión mostrada: **4.2.0**.

## 1. Instalador NSIS

- [ ] Ejecutar `TETORD-Setup-4.2.0-x64.exe`.
- [ ] Confirmar que el instalador permite elegir directorio.
- [ ] Confirmar acceso directo de escritorio y menú Inicio.
- [ ] Abrir TETORD desde el acceso directo.
- [ ] Confirmar que no aparece una consola de Node/servidor inesperada.
- [ ] Desinstalar desde Windows y confirmar que la desinstalación termina correctamente.

## 2. Portable

- [ ] Copiar `TETORD-4.2.0-x64.exe` a una carpeta independiente.
- [ ] Ejecutarlo sin instalación.
- [ ] Confirmar que abre Writer/Office correctamente.
- [ ] Cerrar y volver a abrir el Portable.

## 3. Writer — persistencia local

- [ ] Crear documento Writer nuevo.
- [ ] Escribir texto, aplicar formato básico y guardar como `.tetord`.
- [ ] Confirmar que el archivo aparece en `Documents\TETORD` del usuario.
- [ ] Cerrar TETORD.
- [ ] Abrir el documento desde Archivos.
- [ ] Confirmar que el contenido y metadatos básicos persisten.
- [ ] Renombrar, duplicar, mover y eliminar un documento de prueba.

## 4. Calc — persistencia y colaboración

- [ ] Crear una hoja Calc y guardar `.tetord`.
- [ ] Cerrar/reabrir y comprobar fórmulas y datos.
- [ ] Con dos clientes, abrir el mismo documento/sala colaborativa.
- [ ] Escribir cambios desde ambos clientes.
- [ ] Confirmar ACK y actualización remota sin bucle ni congelamiento.
- [ ] Verificar que los cursores/selecciones compartidos no provocan parpadeo continuo.

## 5. Slides

- [ ] Crear una presentación.
- [ ] Añadir/duplicar/eliminar una diapositiva.
- [ ] Guardar y reabrir.
- [ ] Confirmar que imágenes/assets no se pierden.
- [ ] Comprobar indicador de estado sin bloquear la edición.

## 6. Messenger y presencia

- [ ] Iniciar el servidor TETORD según la documentación de `server/`.
- [ ] Conectar dos clientes autenticados.
- [ ] Confirmar estado `online`.
- [ ] Desconectar uno y confirmar transición a `offline`/ausente según el flujo configurado.
- [ ] Enviar mensaje de sala.
- [ ] Enviar DM.
- [ ] Forzar reconexión y confirmar que el historial no contiene duplicados.
- [ ] Confirmar que un fallo de conexión muestra un Toast claro y mantiene modo local.

## 7. Cloud Sync

- [ ] Configurar el endpoint Cloud Sync de prueba.
- [ ] Iniciar sesión y confirmar token válido.
- [ ] Subir un `.tetord`.
- [ ] Descargarlo desde otro cliente.
- [ ] Modificar localmente y remotamente con timestamps diferentes.
- [ ] Confirmar la resolución de conflicto esperada y que no se sobrescribe silenciosamente una versión más reciente sin aviso.

## 8. Offline-first / red hostil

- [ ] Con TETORD funcionando, cortar Internet/servidor.
- [ ] Seguir escribiendo en Writer.
- [ ] Crear/guardar un documento local.
- [ ] Confirmar que la UI no se congela.
- [ ] Confirmar estado visual de desconexión/reconexión.
- [ ] Restaurar servidor/red.
- [ ] Confirmar reconexión y reenvío de operaciones pendientes.
- [ ] Confirmar ausencia de duplicados después de la resincronización.

## 9. Seguridad / paquete

- [ ] Inspeccionar el contenido del instalador/ASAR y confirmar que no contiene `server/tests`, `scripts`, `.git` ni archivos de desarrollo innecesarios.
- [ ] Confirmar que el servidor no se inicia automáticamente como parte del cliente.
- [ ] Confirmar que documentos fuera de `Documents\TETORD` no pueden ser accedidos mediante traversal.
- [ ] Confirmar que la aplicación sigue iniciando sin conexión.

## 10. Resultado final

**Build:** ____________________

**Windows:** ____________________

**Tester:** ____________________

**Fecha:** ____________________

**PASS / FAIL:** ____________________

**Notas / evidencia:**

____________________________________________________________

____________________________________________________________
