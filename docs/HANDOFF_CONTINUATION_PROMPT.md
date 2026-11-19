# Handoff — estado del proyecto y continuación

Repo: `D4RKL0RD-J0571N/3-way-inventory-reconciliation`, rama `feat/web-react-and-csharp-parity`.

## Estado

| Área | Estado |
|---|---|
| `Acr.Core`, `Acr.Infrastructure` | Completos; incluyen BOM, sanitizador SAP, blindaje IMP-207, detección por firma, montaje a ciegas |
| `Acr.Desktop` (WPF, 8 pestañas) | Completo; iconos vectoriales y tooltips |
| `Acr.Tests` | 221 pruebas, 0 fallos (`dotnet test src/dotnet/Acr.sln`) |
| `web/` | Compila (`npm run build`); referencia visual |
| Empaquetado | `empaquetar_standalone.bat` |
| Python legado | Sin los fixes recientes; solo CLI/lotes |

## Decisiones vigentes

- Regla Cardinal, `IsZeroTotal`, empaque canónico y macro-safe `.xlsm` (ver README).
- Destino `net9.0` por el SDK instalado; migrar a `net10.0` cuando esté el SDK.
- Sin emojis en la UI; iconos `StreamGeometry`.
- `data/` y `publish/` no se versionan.

## Pendientes sugeridos

1. Prueba manual del exe portable con importaciones reales y con el montaje a ciegas (`data/IMP-100_MONTAJE_A_CIEGAS`, solo local).
2. Alinear el texto del aviso de montaje a ciegas entre `web/` (modo pausa sin QC, ya existente) y WPF.
3. Migrar a .NET 10; revisar `CoreIsolationTests`.
4. Soporte `.xls` binario (hoy no soportado) y lectura de PDF/MSG si se requiere.
5. Revisar los documentos marcados *Historical* en `docs/README.md` y archivarlos si ya no aportan.

## Comandos

```powershell
dotnet build src/dotnet/Acr.sln -c Release
dotnet test  src/dotnet/Acr.sln
cd web; npm install; npm run build
.\empaquetar_standalone.bat
```
