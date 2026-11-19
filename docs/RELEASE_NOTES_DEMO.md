# Release Notes

## v2.1.0 — Blind receiving, UI polish (current)

- **Montaje a ciegas:** `ComparisonSummary.IsBlindReceiving`, banner ámbar, warning `[MONTAJE A CIEGAS]`, checklist adaptado (NCR suspendido, `Bultos` pendiente). Tests: `BlindReceivingTests`, `MainViewModelTests`.
- **Examinar...:** la barra de archivos usa 9 columnas independientes; el botón *Ejecutar Conciliación 3-Way* ya no tapa el segundo *Examinar...*.
- **Iconos vectoriales y tooltips en español** en todos los controles (sin emojis).
- **Selección de DataGrid:** fondo celeste claro con texto oscuro; la columna *Empaque Canónico* ya no queda azul sobre azul.
- **Empaquetado:** `empaquetar_standalone.bat` genera `Acr.Desktop.exe` autocontenido (~62 MB) y `web-edition.zip`.
- **Documentación** reescrita y con índice (`docs/README.md`).

## v2.0.0 — Paridad total C# ↔ Web

- Ocho pestañas WPF equivalentes a `web/src/components`.
- **IMP-207 → 288:** exclusión de etiquetas de resumen al leer el ID; bloqueo del primer ID válido; respaldo por nombre de archivo.
- **IMP-207 Layout #4:** `UNIDADES`/`UNID` ignorados como bultos complemento; rescate matemático.
- Hojas `ESQUEMA`/`ALISTO` excluidas al detectar partes; detección QC vs Cierre por firma, libros intercambiados e híbridos (IMP-218); sin tope de 500 filas (IMP-187).
- **Reparto BOM proporcional** (mayor residuo) — `BomAllocationService`.
- **`SapClipboardSanitizer`** (`1E2`, `(48)`, `$-50.00`, delimitadores `\t | ; ,`).
- Edición web React/TypeScript incorporada en `web/`.
- Proyectos en `net9.0` (SDK 9.0.x disponible); 217 → 221 pruebas.

## v1.x — Edición Python/Streamlit (legado)

Motor Python, interfaz Streamlit, CLI y empaquetado PyInstaller. Ver historial de git y [WIKI.md](WIKI.md) para el comportamiento actual.
