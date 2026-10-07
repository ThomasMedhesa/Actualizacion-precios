# Plan de continuación (pendiente a 30/09/2026)

Todo lo funcional ya está hecho y verificado:
11/11 secciones del self-test en verde · regresión `git HEAD` vs nuevo = 56/56 xlsx
idénticos · manual regenerado · `dist\ActualizarPrecios.exe` compilado y arrancado.

Orden sugerido para mañana:

## 1. Sincronizar la ayuda F1 (10 min, obligatorio)
`procesar_ofertas.py:134-159` (constante `AYUDA`) sigue describiendo la interfaz
anterior. Debe coincidir con la sección 3.7 del manual ya regenerada:

- [ ] "Doble clic en ... para editarlos" -> edita con **clic simple** (también doble clic/Intro)
- [ ] Anadir: filtro por texto, orden con clic en cabecera, plegar grupos con la flecha
- [ ] Anadir: botón `Limpiar log`, `Espacio` alterna casilla, `Intro` procesa
- [ ] Atajos: `Ctrl+Enter` -> `Intro`; documentar `Esc` (cierra editor/calendario) y `F1`
- [ ] Mencionar interruptor de tema claro/oscuro y que las preferencias van a `config.json`

## 2. Corregir el aviso engañoso (2 min)
`procesar_ofertas.py:1311-1314`: si no se puede guardar `config.json` el texto dice
"Se usará la carpeta Ofertas/ por defecto", pero en realidad solo avisa y la sesión
sigue con las carpetas ya cargadas. Reescribir el texto o outright quitarlo.

## 3. Comprobación visual del ejecutable (15 min, la única que no puedo hacer yo)
Ejecutar `dist\ActualizarPrecios.exe` con la carpeta `Ofertas\` real al lado y mirar:

- [ ] Tema claro (por defecto) y conmutador a oscuro; que la tabla y el log se adapten bien
- [ ] Editores en celda: comercializadora, Renovable, calendario
- [ ] Calendario: Aceptar / Limpiar / Cancelar (Cancelar **no** debe borrar la fecha)
- [ ] Filtrar, ordenar, plegar grupos, `Espacio` sobre una fila
- [ ] Procesar en segundo plano: log con colores, progreso y botón Cancelar

## 4. Cerrar el self-test (decisión)
Está en `%TEMP%\opencode\selftest_ui.py` (fuera del repo a propósito, así que
el commit `d925c88` no lo incluye).
Opciones: moverlo al repo como `test_interfaz.py`, o borrarlo.

## 5. Commit — YA HECHO (por el usuario, no por mí)
`d925c88 "Add customtkinter dependency and create config.json for application settings"`
contiene todo: `procesar_ofertas.py`, `llenar_planillas.py`, `generar_manual.py`,
`requirements.txt`, el manual, `build/*` y `dist\ActualizarPrecios.exe`.
La rama está sincronizada con `origin/main`: no hay nada que empujar.

Pendiente de decidir (tú):

- [ ] `dist/config.json` se commiteó y contiene **rutas absolutas de esta máquina**
      (`dir_ofertas`, `dir_salida`). Es configuración de runtime, no código.
      ¿Sacarlo del repo con `.gitignore` + `git rm --cached`? Ahora mismo solo
      importa para tener una copia del `.exe`, porque al arrancar se regenera solo.
- [ ] No hay `.gitignore` en el proyecto y `build/`, `dist/` y `__pycache__/`
      están versionados. Si se quiere dejar de versionarlos, es un `git rm -r --cached`
      más un `.gitignore` con `build/`, `dist/`, `__pycache__/`, `config.json`, `*.pyc`.
      (Ojo: si se deja de versionar `dist/ActualizarPrecios.exe`, hay que subirlo
      aparte cuando se genere una versión nueva.)

Comprobaciones si se vuelve a tocar el código:

```
python -m py_compile procesar_ofertas.py llenar_planillas.py
python selftest_ui.py
python generar_manual.py
```

## Notas menores (opcionales)
- Cabecera del calendario con azul fijo `#3B8ED0` en modo oscuro.
- `ToolTip` usa `"Segoe UI"` fijo en lugar de `self.familia`.