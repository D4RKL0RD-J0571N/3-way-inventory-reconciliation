# ACR (Asistente de Cierre de Recepción): Resumen de trabajo y prompt de continuación

## 1. Resumen del trabajo realizado

Repo: `D4RKL0RD-J0571N/3-way-inventory-reconciliation`, rama `repository_audit_planning`.
Objetivo: aplicación de escritorio personal en C# (.NET 10 LTS, WPF) para cierres de recepción de contenedores de importación, conciliación QC vs proveedor, control diario y base de conocimiento SAP offline.

| Fase | Estado | Detalle |
|---|---|---|
| Fase 0: Auditoría | PASS | `docs/PHASE_0_AUDIT_REPORT.md`. 58 columnas (A..BF) en fila 12 de `CIERRE RECEP`. 6 etapas de ciclo de vida. 5 libros históricos corruptos preexistentes (IMP-241, 244, 256, 360, 90). Autor en `'DATOS GENERALES'!B33`. |
| Fase 1: `Acr.Core` | PASS | Modelos y servicios de dominio, `SapKnowledge` (modelos), 104 tests. |
| Pista E2E fase A | PASS | `docs/e2e/E2E_TEST_SPECIFICATIONS.md` (254 casos) y 47 fixtures en `tests/e2e_fixtures/`. |
| Fase 2-A: Infraestructura | PASS (compuerta formal) | `QcReader`, `WorkbookSignatureScanner`, `HistoricalClosureScanner`, `OperationalSafetyGuard`. |
| Fase 2-B: Infraestructura | Implementada, compuerta NO cerrada | `ClosureExporter`, `FormulaIntegrityAuditor`, `DailyLogRepository/Service`, `AtomicFileTransaction`, `ReceivingPipeline`, `Reporting/`. 175/175 tests pasan. |
| Fases 3 a 6 | No iniciadas | Paridad de regresión, WPF, SAP KB, empaquetado. |

Commits: `e8c601b` (Fase 0 + 1) y el commit de Fase 2 con este documento.

### Defectos pendientes de la verificación adversaria de M2-B
Se reportaron 4. El worker de remediación pudo haber corregido algunos (el conteo subió de 165 a 175 tests). **Verifica cada uno con un test antes de darlo por resuelto:**
1. **Alto.** `JsonProcessingReportStore.BuildReportFileName`: `ReportId` sin sanitizar (path traversal, subdirectorios, ADS). Sin tope de longitud en tokens (`IOException` con `ContainerId` de 300 caracteres).
2. **Alto.** `OperationHistoryStore.Append`: sin sincronización por archivo. Con 10 hilos hubo excepciones y pérdida de escrituras. Usar lock o `SemaphoreSlim` por ruta canónica.
3. **Medio.** `DailyLogRepository.WriteDataRow`: no borra la columna C cuando `UpdatedDate` es null.
4. **Medio.** `JsonProcessingReportStore.LoadAll`: falla si hay otro `.json` en la carpeta (ej. `operation_history.json`). Filtrar por `report_*.json`.

## 2. Entorno
- SDK .NET 10 instalado en `%LocalAppData%\Microsoft\dotnet`. Antes de cada comando:
  `$env:PATH = "$env:LOCALAPPDATA\Microsoft\dotnet;" + $env:PATH; $env:DOTNET_ROOT="$env:LOCALAPPDATA\Microsoft\dotnet"`
- Solución: `src/dotnet/Acr.sln` con `Acr.Core`, `Acr.Infrastructure`, `Acr.Tests` (falta `Acr.Desktop`).
- Línea base Python (autoritativa hasta cerrar paridad): `src/` y `tests/test_real_dataset_audit.py` (58/58).
- Datos de producción (SOLO LECTURA): `D:\Development\Software Development\3 Way Inventory Reconciliation\data` y `...\Control_Diario_Operaciones_Jostin.xlsx`.
- Estado de teamwork: el coordinador se detuvo por cuota agotada. Notas internas en `%USERPROFILE%\.gemini\antigravity\scratch\teamwork\607b2f1d-d517-4470-b4e8-1178ccb12f9e\.agents\teamwork\` (`PROJECT.md`, `sub_orch_m2_b\GATE_STATUS.md`, `teamwork_challenger_m2b_r1_2\handoff.md`).

## 3. Prompt para otro LLM

```
Continúas el proyecto ACR (Asistente de Cierre de Recepción) en el repo
D4RKL0RD-J0571N/3-way-inventory-reconciliation, rama repository_audit_planning.
Lee primero docs/HANDOFF_CONTINUATION_PROMPT.md, docs/PHASE_0_AUDIT_REPORT.md y
docs/e2e/E2E_TEST_SPECIFICATIONS.md.

Producto: herramienta personal de escritorio (C#, .NET 10 LTS, WPF, CommunityToolkit.Mvvm)
para cierres de recepción de contenedores de importación. Sin login, sin servidores,
sin microservicios, sin LLM ni vector DB, ejecutable portable. Tema corporativo sobrio,
sin emojis. ACR nunca ejecuta ni modifica transacciones SAP.

Reglas: PascalCase, funciones < 50 líneas, XML docs en APIs públicas, Acr.Core sin
dependencias externas. Los datos de producción y Control_Diario_Operaciones_Jostin.xlsx
son SOLO LECTURA; las pruebas usan copias temporales. Nunca compenses entre materiales
distintos (Cardinal Discrepancy Rule). Separa identidad de negocio (contenedor, PO, fecha)
de procedencia (ruta, hash). Antes de cada commit, dotnet test debe pasar sin fallos.

Estado: Fase 0, Fase 1 y Fase 2-A cerradas. Fase 2-B implementada (175 tests verdes) con
4 defectos por verificar o corregir (lista en la sección 1 del documento de traspaso).

Tareas, en orden:
1. Verifica y corrige los 4 defectos de M2-B, cada uno con un test de regresión.
2. Fase 3: ejecuta el pipeline C# sobre los 51 QC y 77 cierres reales (solo lectura) y
   compara con la línea base Python. Produce un reporte de excepciones; los 5 libros
   corruptos preexistentes se clasifican como excepciones, no como errores del exportador.
3. Fase 4: crea src/dotnet/Acr.Desktop con 3 vistas (Conciliación, Control Diario,
   Sincronización Histórica) y ViewModels con tests. Español e inglés.
4. Fase 5: base de conocimiento SAP offline (índice determinista, enciclopedia ME80FN,
   SE16N, MB51, VL31N, RFUI, MIGO, ME21N/22N/23N, ZDescarga, EINA, MAST, STPO, guías de
   diagnóstico enlazadas desde las discrepancias).
5. Fase 6: publicar ACR.exe portable (self-contained, single-file), smoke test y push.

Haz commit y push al cerrar cada fase. Trabaja en pasos pequeños y verifica con tests reales.
```
