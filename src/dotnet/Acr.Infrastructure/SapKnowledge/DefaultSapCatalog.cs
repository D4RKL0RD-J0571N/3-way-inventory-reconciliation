using Acr.Core.Models;
using Acr.Core.SapKnowledge;

namespace Acr.Infrastructure.SapKnowledge;

/// <summary>
/// Offline catalog containing default SAP encyclopedic knowledge entries and diagnostic guides.
/// </summary>
public static class DefaultSapCatalog
{
    private static readonly Lazy<IReadOnlyList<TransactionEntry>> LazyTransactions = new(BuildTransactions);
    private static readonly Lazy<IReadOnlyList<DiagnosticGuide>> LazyGuides = new(BuildGuides);

    /// <summary>Default list of initial SAP transaction and table encyclopedia entries.</summary>
    public static IReadOnlyList<TransactionEntry> Transactions => LazyTransactions.Value;

    /// <summary>Default list of guided troubleshooting procedures.</summary>
    public static IReadOnlyList<DiagnosticGuide> DiagnosticGuides => LazyGuides.Value;

    private static IReadOnlyList<TransactionEntry> BuildTransactions()
    {
        return new List<TransactionEntry>
        {
            new()
            {
                Code = "ME80FN",
                Title = "General Evaluations for Purchase Orders",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Historial completo de pedidos de compra, entregas entrantes e historial de movimientos por material o proveedor.",
                WhenToUse = "Investigar discrepancias de cantidades facturadas vs recibidas o revisar el historial de recepciones anteriores del PO.",
                NavigationPath = "SAP Menu -> Logística -> Gestión de materiales -> Compras -> Pedido -> Evaluaciones -> ME80FN",
                Inputs = ["Número de Pedido (PO 43...)", "Material", "Centro (C001)", "Fechas de documento"],
                Outputs = ["Historial de pedidos", "Entregas asociadas", "Historial de modificaciones"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "SE16N",
                Title = "General Table Display",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Consulta directa de tablas de base de datos SAP (EKKO, EKPO, EINA, MAST, STPO) para verificación de datos maestros y compras.",
                WhenToUse = "Verificar registros de info de compras (EINA), listas de materiales (MAST/STPO) o parámetros técnicos del pedido.",
                NavigationPath = "Transacción directa /nSE16N",
                Inputs = ["Nombre de tabla", "Filtros de campo (EBELN, MATNR, WERKS)"],
                Outputs = ["Filas de base de datos crudas"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "MB51",
                Title = "Material Document List",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Listado de documentos de material (movimientos 101, 102, 103, 105, 501) por material, centro o almacén.",
                WhenToUse = "Rastrear entradas de mercancías exactas contabilizadas para un contenedor o verificar si una entrega ya fue ingresada.",
                NavigationPath = "Logística -> Gestión de materiales -> Gestión de stocks -> Entorno -> Información de stocks -> Documentos de material",
                Inputs = ["Material", "Centro C001", "Clase de movimiento (101)", "Fecha de contabilización"],
                Outputs = ["Documentos de material", "Cantidades", "Unidad de medida", "Usuario"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "VL31N",
                Title = "Create Inbound Delivery",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Creación de la entrega entrante previa a la recepción física.",
                WhenToUse = "Generar el documento de entrega entrante a partir de un pedido de compra para habilitar la descarga en almacén.",
                NavigationPath = "Logística -> Logistics Execution -> Entrada de mercancías -> Para entrega entrante -> Crear -> Documento individual",
                Inputs = ["Número de pedido (PO)", "Proveedor", "Fecha de entrega"],
                Outputs = ["Número de Entrega Entrante (18...)"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "RFUI",
                Title = "Radio Frequency Mobile UI",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Interfaz para terminales de radiofrecuencia en almacén (EWM/WM) para confirmación de descarga, conteo y asignación de HU.",
                WhenToUse = "Confirmar tareas de descarga por contenedor y verificar unidades de manipulación (Handling Units - HU).",
                NavigationPath = "Transacción directa /nRFUI o terminal RF",
                Inputs = ["Número de Tarea de Almacén", "Handling Unit (HU)", "Ubicación de destino"],
                Outputs = ["Confirmación de tarea", "Estado de stock"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "MIGO",
                Title = "Goods Movement",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Contabilización de entrada de mercancías (movimiento 101) referida a Pedido o Entrega Entrante.",
                WhenToUse = "Ejecutar la recepción formal de inventario tras finalizar la conciliación y revisión de calidad.",
                NavigationPath = "Logística -> Gestión de materiales -> Gestión de stocks -> Movimiento de mercancías -> MIGO",
                Inputs = ["Operación: Entrada de mercancías (A01)", "Referencia: Pedido (R01) o Entrega (R08)", "Cantidades verificadas"],
                Outputs = ["Documento de material generado", "Contabilización de stock en almacén"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "ME21N",
                Title = "Create Purchase Order",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Creación de pedidos de compra.",
                WhenToUse = "Consultar estructura estándar de pedidos.",
                NavigationPath = "Compras -> Pedido -> Crear -> ME21N",
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "ME22N",
                Title = "Change Purchase Order",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Modificación de pedidos de compra.",
                WhenToUse = "Verificar historial de cambios o adiciones de posiciones al pedido.",
                NavigationPath = "Compras -> Pedido -> Modificar -> ME22N",
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "ME23N",
                Title = "Display Purchase Order",
                ObjectKind = SapObjectKind.StandardTransaction,
                Purpose = "Visualización detallada de pedidos de compra (PO 43...).",
                WhenToUse = "Consultar el detalle de posiciones, curva de tallas/colores, cantidades pedidas y estatus de entrega.",
                NavigationPath = "Compras -> Pedido -> Visualizar -> ME23N",
                Inputs = ["Número de pedido (EBELN)"],
                Outputs = ["Detalle de posiciones", "Precios", "Condiciones de entrega"],
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "ZDescarga",
                Title = "Custom Container Unloading Transaction",
                ObjectKind = SapObjectKind.ZTransaction,
                Purpose = "Transacción Z interna para el control operativo de descarga de contenedores de importación.",
                WhenToUse = "Registrar inicio, fin y novedades de la descarga física en andén.",
                NavigationPath = "Menú Z interno /nZDescarga",
                Inputs = ["Número de contenedor", "Andén", "Cuadrilla"],
                Outputs = ["Reporte de tiempos y pallets descargados"],
                EpistemicTier = EpistemicTier.CompanyVerified
            },
            new()
            {
                Code = "EINA",
                Title = "Purchasing Info Record - General Data",
                ObjectKind = SapObjectKind.SapTable,
                Purpose = "Tabla que almacena datos generales del registro de info de compras (relación material-proveedor).",
                WhenToUse = "Consultar códigos de material del proveedor, empaques y tolerancias de exceso/defecto.",
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "MAST",
                Title = "Material to BOM Link",
                ObjectKind = SapObjectKind.SapTable,
                Purpose = "Tabla que vincula materiales con listas de materiales (BOM).",
                WhenToUse = "Investigar si un material compuesto o prepack tiene desgloses en subcomponentes.",
                EpistemicTier = EpistemicTier.OfficialSapDoc
            },
            new()
            {
                Code = "STPO",
                Title = "BOM Items",
                ObjectKind = SapObjectKind.SapTable,
                Purpose = "Posiciones de lista de materiales que detallan los componentes y cantidades de un conjunto.",
                WhenToUse = "Validar curvas de tallas o componentes hijos en sets preempacados.",
                EpistemicTier = EpistemicTier.OfficialSapDoc
            }
        };
    }

    private static IReadOnlyList<DiagnosticGuide> BuildGuides()
    {
        return new List<DiagnosticGuide>
        {
            new()
            {
                GuideId = "DIAG-01-FALTANTE",
                Title = "Investigación de Material Faltante (Shortage)",
                ProblemSummaryAndAffectedObjects = new DiagnosticProblemSummary(
                    "Diferencia negativa entre las unidades físicas recibidas vs unidades declaradas en factura.",
                    ["PO_43", "MATERIAL", "CONTAINER"]),
                TargetDiscrepancyTypes = [DiscrepancyType.Shortage],
                TransactionCodes = ["ME80FN", "ME23N", "MB51"],
                InvestigationSteps =
                [
                    "1. Verificar si existen bultos complementos o cajas en otra tarima no reportadas en el conteo inicial.",
                    "2. En ME23N / ME80FN, revisar si el proveedor facturó la cantidad completa o si hubo aviso de embarque parcial (ASN).",
                    "3. Consultar MB51 con movimiento 101 para asegurar que no se haya ingresado previamente por otra entrega.",
                    "4. Si la diferencia persiste tras revisión física en andén, clasificar formalmente como FALTANTE y adjuntar en el cierre."
                ]
            },
            new()
            {
                GuideId = "DIAG-02-SOBRANTE",
                Title = "Investigación de Material Sobrante (Overage)",
                ProblemSummaryAndAffectedObjects = new DiagnosticProblemSummary(
                    "Diferencia positiva con unidades físicas excedentes respecto a la factura del proveedor.",
                    ["PO_43", "MATERIAL", "CONTAINER"]),
                TargetDiscrepancyTypes = [DiscrepancyType.Overage],
                TransactionCodes = ["ME23N", "SE16N"],
                InvestigationSteps =
                [
                    "1. Confirmar en ME23N la pestaña 'Entrega' para verificar la tolerancia de exceso de suministro (Overdelivery Tolerance).",
                    "2. Verificar si el material sobrante corresponde a un 'Revuelto' compensatorio de otro material de la misma familia.",
                    "3. En caso de ser exceso genuino fuera de tolerancia, notificar a Compras para ampliación de pedido o separación en almacén."
                ]
            },
            new()
            {
                GuideId = "DIAG-03-REVUELTO",
                Title = "Investigación y Confirmación de Material Revuelto (Cross-Pack)",
                ProblemSummaryAndAffectedObjects = new DiagnosticProblemSummary(
                    "Par compensatorio entre materiales de la misma curva donde uno falta y el otro sobra en igual magnitud.",
                    ["PO_43", "MATERIAL", "CONTAINER"]),
                TargetDiscrepancyTypes = [DiscrepancyType.RevueltoCandidate],
                TransactionCodes = ["SE16N", "ME23N"],
                InvestigationSteps =
                [
                    "1. Verificar que la suma neta de unidades entre el par coincida (neto cero o cercano a cero).",
                    "2. Consultar EINA/STPO para comprobar si ambos materiales comparten la misma orden de empaque o proveedor.",
                    "3. Documentar ambos materiales en el Cierre con la etiqueta REVUELTO y su contraparte de referencia sin alterar el pedido original."
                ]
            }
        };
    }
}
