using Acr.Core.Models;
using Acr.Core.SapKnowledge;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;
using Acr.Infrastructure.Safety;
using Acr.Infrastructure.SapKnowledge;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace Acr.Desktop.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly OperationalSafetyGuard _safetyGuard;

    // ==========================================
    // RECONCILIATION & CLOSURE PIPELINE
    // ==========================================
    [ObservableProperty]
    private string _qcFilePath = "";

    [ObservableProperty]
    private string _closureTemplatePath = "";

    [ObservableProperty]
    private string _exportOutputPath = "";

    [ObservableProperty]
    private string _statusMessage = "Listo para iniciar operaciones.";

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private bool _hideZeroTotalItems = true;

    [ObservableProperty]
    private bool _isBlindReceiving;

    [ObservableProperty]
    private ReceivingPipelineResult? _lastReconciliationResult;

    [ObservableProperty]
    private ObservableCollection<DiscrepancyItem> _discrepancies = new();

    [ObservableProperty]
    private ObservableCollection<DiscrepancyItem> _filteredDiscrepancies = new();

    [ObservableProperty]
    private DiscrepancyItem? _selectedDiscrepancy;

    // ==========================================
    // TAB 1: CUADRE & BULTOS + CALCULATOR
    // ==========================================
    [ObservableProperty]
    private bool _strictCustomsMode;

    [ObservableProperty]
    private ObservableCollection<BultosPresentationItem> _bultosItems = new();

    [ObservableProperty]
    private ObservableCollection<BultosPresentationItem> _filteredBultosItems = new();

    [ObservableProperty]
    private BultosPresentationItem? _selectedBultoItem;

    [ObservableProperty]
    private int _totalBultosDescarga;

    [ObservableProperty]
    private int _totalBultosPl;

    [ObservableProperty]
    private int _difBultos;

    [ObservableProperty]
    private int _totalBultosEnteros;

    [ObservableProperty]
    private int _totalBultosComplementos;

    // Interactive carton calculator
    [ObservableProperty]
    private int _calcUnits = 420;

    [ObservableProperty]
    private int _calcCurve = 12;

    [ObservableProperty]
    private int _calcStdBox = 48;

    [ObservableProperty]
    private bool _calcHasOdds = true;

    [ObservableProperty]
    private bool _calcOddModeIsLoose;

    [ObservableProperty]
    private int _calcOddUnits = 36;

    [ObservableProperty]
    private int _calcTotalBoxes;

    [ObservableProperty]
    private int _calcBoxesFull;

    [ObservableProperty]
    private int _calcBoxesComp;

    [ObservableProperty]
    private double _calcResPrk;

    [ObservableProperty]
    private bool _calcCurveOk;

    [ObservableProperty]
    private string _calcPkgStr = "";

    // ==========================================
    // TAB 2: FICHA 360° MATERIAL
    // ==========================================
    [ObservableProperty]
    private DiscrepancyItem? _selectedInspectorItem;

    [ObservableProperty]
    private string _inspectorSearchQuery = "";

    [ObservableProperty]
    private string _inspectorStatusFilter = "ALL";

    [ObservableProperty]
    private ObservableCollection<DiscrepancyItem> _filteredInspectorItems = new();

    // ==========================================
    // TAB 3: TABLA SEMÁFORO
    // ==========================================
    [ObservableProperty]
    private string _semaforoSearchQuery = "";

    [ObservableProperty]
    private string _semaforoStatusFilter = "ALL";

    // ==========================================
    // TAB 4: EXPLOSIÓN CURVAS (BOM)
    // ==========================================
    [ObservableProperty]
    private ObservableCollection<DiscrepancyItem> _prepacksWithBom = new();

    [ObservableProperty]
    private DiscrepancyItem? _selectedBomPrepack;

    // ==========================================
    // TAB 5: VALIDADOR SAP (PORTAPAPELES)
    // ==========================================
    [ObservableProperty]
    private string _sapRawPasteText = "";

    [ObservableProperty]
    private SanitizedSapTable? _sanitizedTable;

    [ObservableProperty]
    private ObservableCollection<string> _sapSanitizedHeaders = new();

    [ObservableProperty]
    private ObservableCollection<IReadOnlyList<string>> _sapSanitizedRows = new();

    [ObservableProperty]
    private string _sapPasteError = "";

    [ObservableProperty]
    private string _sapDelimiterLabel = "Tabulación";

    // ==========================================
    // TAB 6: COPIAR A EXCEL & EXPORTAR
    // ==========================================
    [ObservableProperty]
    private int _excelExportSubTab;

    [ObservableProperty]
    private string _tira1AsAtText = "";

    [ObservableProperty]
    private string _tira2AwAyText = "";

    [ObservableProperty]
    private string _xlsmVbaSha256Result = "";

    [ObservableProperty]
    private bool? _xlsmSha256Preserved;

    // ==========================================
    // TAB 7: CHECKLIST DE EXPEDIENTE FÍSICO
    // ==========================================
    [ObservableProperty]
    private bool _checklist1Portada = true;

    [ObservableProperty]
    private bool _checklist2DatosGenerales = true;

    [ObservableProperty]
    private bool _checklist3CierreRecep = true;

    [ObservableProperty]
    private bool _checklist4NotaCredito = true;

    [ObservableProperty]
    private bool _checklist5Mb51 = true;

    [ObservableProperty]
    private bool _checklist6Bultos = true;

    [ObservableProperty]
    private bool _checklist7Dua = true;

    [ObservableProperty]
    private string _checklist1Text = "";

    [ObservableProperty]
    private string _checklist2Text = "";

    [ObservableProperty]
    private string _checklist3Text = "";

    [ObservableProperty]
    private string _checklist4Text = "";

    [ObservableProperty]
    private string _checklist5Text = "";

    [ObservableProperty]
    private string _checklist6Text = "";

    [ObservableProperty]
    private string _checklist7Text = "";

    [ObservableProperty]
    private int _checklistCompletedCount = 7;

    [ObservableProperty]
    private int _checklistProgress = 100;

    [ObservableProperty]
    private bool _checklistIsComplete = true;

    // ==========================================
    // TAB 8: SAP & WIKI
    // ==========================================
    [ObservableProperty]
    private int _wikiSubTab;

    [ObservableProperty]
    private string _sapSearchQuery = "";

    [ObservableProperty]
    private ObservableCollection<TransactionEntry> _filteredSapTransactions = new();

    [ObservableProperty]
    private TransactionEntry? _selectedSapTransaction;

    [ObservableProperty]
    private ObservableCollection<DiagnosticGuide> _diagnosticGuides = new();

    [ObservableProperty]
    private DiagnosticGuide? _selectedDiagnosticGuide;

    [ObservableProperty]
    private string _sopManualText = "";

    // ==========================================
    // EXECUTIVE KPI STRIP
    // ==========================================
    [ObservableProperty]
    private string _kpiContainerId = "-";

    [ObservableProperty]
    private string _kpiPoNumber = "-";

    [ObservableProperty]
    private string _kpiPedidoNumber = "-";

    [ObservableProperty]
    private int _kpiTotalStyles;

    [ObservableProperty]
    private int _kpiCleanStyles;

    [ObservableProperty]
    private int _kpiDiscrepantStyles;

    [ObservableProperty]
    private int _kpiDeclaredUnits;

    [ObservableProperty]
    private int _kpiPhysicalUnits;

    [ObservableProperty]
    private int _kpiDeltaUnits;

    [ObservableProperty]
    private int _kpiDeclaredCartons;

    [ObservableProperty]
    private int _kpiPhysicalCartons;

    [ObservableProperty]
    private int _kpiDeltaCartons;

    [ObservableProperty]
    private int _kpiRevueltoPairs;

    [ObservableProperty]
    private int _kpiDamagedUnits;

    // ==========================================
    // DAILY LOG 6-STAGES (BACKWARD COMPAT)
    // ==========================================
    [ObservableProperty]
    private string _dailyLogFilePath = @"D:\Development\Software Development\3 Way Inventory Reconciliation\Control_Diario_Operaciones_Jostin.xlsx";

    [ObservableProperty]
    private string _containerId = "";

    [ObservableProperty]
    private string _po43Number = "";

    [ObservableProperty]
    private string _pedido41Number = "";

    [ObservableProperty]
    private string _entrega18Number = "";

    [ObservableProperty]
    private string _ewmTaskNumber = "";

    [ObservableProperty]
    private string _handlingUnit = "";

    [ObservableProperty]
    private Saludo _selectedSaludo = Saludo.BuenosDias;

    [ObservableProperty]
    private CustomsDuaStage _stage1Customs = CustomsDuaStage.DuaAndDocsReady;

    [ObservableProperty]
    private SapIntakeStage _stage2SapIntake = SapIntakeStage.ReadyAndNotified;

    [ObservableProperty]
    private BlindSetupStage _stage3BlindSetup = BlindSetupStage.ReadyComplete;

    [ObservableProperty]
    private QualityControlStage _stage4QualityControl = QualityControlStage.QcApprovedZDescarga;

    [ObservableProperty]
    private ReceptionClosureStage _stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff;

    [ObservableProperty]
    private DispatchClosureStage _stage6DispatchClosure = DispatchClosureStage.DispatchClosureReady;

    [ObservableProperty]
    private string _computedColumnO = "1. En Cola Conteo a Ciegas";

    [ObservableProperty]
    private string _computedColumnR = "-";

    [ObservableProperty]
    private string _clipboardSummary = "";

    public IReadOnlyList<Saludo> SaludoOptions { get; } = Enum.GetValues<Saludo>();
    public IReadOnlyList<CustomsDuaStage> Stage1Options { get; } = Enum.GetValues<CustomsDuaStage>();
    public IReadOnlyList<SapIntakeStage> Stage2Options { get; } = Enum.GetValues<SapIntakeStage>();
    public IReadOnlyList<BlindSetupStage> Stage3Options { get; } = Enum.GetValues<BlindSetupStage>();
    public IReadOnlyList<QualityControlStage> Stage4Options { get; } = Enum.GetValues<QualityControlStage>();
    public IReadOnlyList<ReceptionClosureStage> Stage5Options { get; } = Enum.GetValues<ReceptionClosureStage>();
    public IReadOnlyList<DispatchClosureStage> Stage6Options { get; } = Enum.GetValues<DispatchClosureStage>();

    [ObservableProperty]
    private int _selectedTabIndex;

    public MainViewModel()
    {
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string baseDir = Path.Combine(localAppData, "ACR_ReceivingAssistant");
        Directory.CreateDirectory(baseDir);

        _safetyGuard = new OperationalSafetyGuard(
            OperationalSafetyMode.Production,
            baseDir,
            _dailyLogFilePath);

        foreach (var t in DefaultSapCatalog.Transactions)
        {
            FilteredSapTransactions.Add(t);
        }

        foreach (var g in DefaultSapCatalog.DiagnosticGuides)
        {
            DiagnosticGuides.Add(g);
        }

        InitializeSopManualText();
        RecalculateCartonCalculator();
        UpdateChecklistDynamicLabels();
        RecalculateDailyLogFormulas();
    }

    // ==========================================
    // CALCULATOR REACTIVITY
    // ==========================================
    partial void OnCalcUnitsChanged(int value) => RecalculateCartonCalculator();
    partial void OnCalcCurveChanged(int value) => RecalculateCartonCalculator();
    partial void OnCalcStdBoxChanged(int value) => RecalculateCartonCalculator();
    partial void OnCalcHasOddsChanged(bool value) => RecalculateCartonCalculator();
    partial void OnCalcOddModeIsLooseChanged(bool value) => RecalculateCartonCalculator();
    partial void OnCalcOddUnitsChanged(int value) => RecalculateCartonCalculator();

    public void RecalculateCartonCalculator()
    {
        int safeStd = Math.Max(1, CalcStdBox);
        int safeCurve = Math.Max(1, CalcCurve);
        int boxesFull = 0;
        int boxesComp = 0;
        int oddFinal = CalcHasOdds ? CalcOddUnits : 0;
        string pkgStr = "";

        if (CalcHasOdds && oddFinal > 0 && CalcUnits > oddFinal)
        {
            int fullBoxUnits = CalcUnits - oddFinal;
            boxesFull = fullBoxUnits / safeStd;
            int rem = fullBoxUnits % safeStd;
            if (rem > 0) oddFinal += rem;

            if (CalcOddModeIsLoose)
            {
                boxesComp = 0;
                pkgStr = $"{boxesFull}*{safeStd} +{oddFinal}";
            }
            else
            {
                boxesComp = 1;
                pkgStr = $"{boxesFull}*{safeStd} 1*{oddFinal}";
            }
        }
        else
        {
            boxesFull = CalcUnits / safeStd;
            int rem = CalcUnits % safeStd;
            if (rem > 0)
            {
                if (CalcOddModeIsLoose)
                {
                    boxesComp = 0;
                    pkgStr = $"{boxesFull}*{safeStd} +{rem}";
                }
                else
                {
                    boxesComp = 1;
                    pkgStr = $"{boxesFull}*{safeStd} 1*{rem}";
                }
            }
            else
            {
                boxesComp = 0;
                pkgStr = safeStd.ToString();
            }
        }

        CalcTotalBoxes = boxesFull + boxesComp;
        CalcBoxesFull = boxesFull;
        CalcBoxesComp = boxesComp;
        CalcResPrk = (double)CalcUnits / safeCurve;
        CalcCurveOk = CalcUnits % safeCurve == 0;
        CalcPkgStr = pkgStr;
    }

    // ==========================================
    // FILTERING REACTIVITY
    // ==========================================
    partial void OnHideZeroTotalItemsChanged(bool value)
    {
        ApplyDiscrepancyFilter();
        ApplyBultosFilter();
        ApplyInspectorFilter();
    }

    partial void OnSemaforoSearchQueryChanged(string value) => ApplyDiscrepancyFilter();
    partial void OnSemaforoStatusFilterChanged(string value) => ApplyDiscrepancyFilter();

    partial void OnInspectorSearchQueryChanged(string value) => ApplyInspectorFilter();
    partial void OnInspectorStatusFilterChanged(string value) => ApplyInspectorFilter();

    private void ApplyDiscrepancyFilter()
    {
        FilteredDiscrepancies.Clear();
        string q = (SemaforoSearchQuery ?? "").Trim().ToLowerInvariant();
        string status = (SemaforoStatusFilter ?? "ALL").Trim();

        foreach (var d in Discrepancies)
        {
            if (HideZeroTotalItems && d.IsZeroTotal)
                continue;

            if (status != "ALL" && !d.DiscrepancyTypeDescription.Contains(status, StringComparison.OrdinalIgnoreCase))
                continue;

            if (!string.IsNullOrWhiteSpace(q) &&
                !d.MaterialCode.ToLowerInvariant().Contains(q) &&
                !(d.VendorMaterial ?? "").ToLowerInvariant().Contains(q) &&
                !(d.Description ?? "").ToLowerInvariant().Contains(q))
            {
                continue;
            }

            FilteredDiscrepancies.Add(d);
        }
    }

    private void ApplyBultosFilter()
    {
        FilteredBultosItems.Clear();
        foreach (var b in BultosItems)
        {
            if (HideZeroTotalItems && b.IsZeroTotal)
                continue;
            FilteredBultosItems.Add(b);
        }
    }

    private void ApplyInspectorFilter()
    {
        FilteredInspectorItems.Clear();
        string q = (InspectorSearchQuery ?? "").Trim().ToLowerInvariant();
        string status = (InspectorStatusFilter ?? "ALL").Trim();

        foreach (var d in Discrepancies)
        {
            if (HideZeroTotalItems && d.IsZeroTotal)
                continue;

            if (status != "ALL" && !d.DiscrepancyTypeDescription.Contains(status, StringComparison.OrdinalIgnoreCase))
                continue;

            if (!string.IsNullOrWhiteSpace(q) &&
                !d.MaterialCode.ToLowerInvariant().Contains(q) &&
                !(d.VendorMaterial ?? "").ToLowerInvariant().Contains(q) &&
                !(d.Description ?? "").ToLowerInvariant().Contains(q))
            {
                continue;
            }

            FilteredInspectorItems.Add(d);
        }

        if (SelectedInspectorItem == null || !FilteredInspectorItems.Contains(SelectedInspectorItem))
        {
            SelectedInspectorItem = FilteredInspectorItems.FirstOrDefault();
        }
    }

    // ==========================================
    // CHECKLIST REACTIVITY
    // ==========================================
    partial void OnChecklist1PortadaChanged(bool value) => RecalculateChecklistProgress();
    partial void OnChecklist2DatosGeneralesChanged(bool value) => RecalculateChecklistProgress();
    partial void OnChecklist3CierreRecepChanged(bool value) => RecalculateChecklistProgress();
    partial void OnChecklist4NotaCreditoChanged(bool value) => RecalculateChecklistProgress();
    partial void OnChecklist5Mb51Changed(bool value) => RecalculateChecklistProgress();
    partial void OnChecklist6BultosChanged(bool value) => RecalculateChecklistProgress();
    partial void OnChecklist7DuaChanged(bool value) => RecalculateChecklistProgress();

    private void RecalculateChecklistProgress()
    {
        int count = 0;
        if (Checklist1Portada) count++;
        if (Checklist2DatosGenerales) count++;
        if (Checklist3CierreRecep) count++;
        if (Checklist4NotaCredito) count++;
        if (Checklist5Mb51) count++;
        if (Checklist6Bultos) count++;
        if (Checklist7Dua) count++;

        ChecklistCompletedCount = count;
        ChecklistProgress = (int)Math.Round((double)count / 7 * 100);
        ChecklistIsComplete = count == 7;
    }

    public void UpdateChecklistDynamicLabels()
    {
        string prov = LastReconciliationResult?.DeclaredItems.FirstOrDefault()?.VendorMaterial ?? "PROVEEDOR INTERNACIONAL";
        string impId = KpiContainerId != "-" ? KpiContainerId : "IMPORTACIÓN ACTIVA";

        Checklist1Text = $"1. Portada de Proveedor (Hoja 'PORTADA') — Carátula oficial con nombre del proveedor ({prov}) y código {impId}.";
        Checklist2Text = $"2. Datos Generales (Hoja 'DATOS GENERALES') — Resumen con POs ({KpiDeclaredUnits:N0} UN Factura 41) y totales.";
        Checklist3Text = $"3. Carátula del Cierre Oficial (Hoja 'CIERRE RECEP') — Balance general con unidades recibidas ({KpiPhysicalUnits:N0} UN) y bultos ({KpiPhysicalCartons:N0}).";

        decimal totalNcr = Discrepancies.Sum(d => d.NcrAmount);
        Checklist4Text = totalNcr == 0m
            ? "4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: LIMPIO ($0.00)."
            : $"4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: RECLAMO POR ${totalNcr:N2} USD.";

        Checklist5Text = "5. Respaldo en ERP de ingreso pedido 41 (Mov 101 en MB51) — Impresión de MB51, clases 101-102.";
        Checklist6Text = "6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos') — Desglose físico de cartones enteros y complementos.";
        Checklist7Text = "7. DUA (Documentación Única Aduanera) — Copia aduanera sellada con peso bruto y valor CIF.";

        if (IsBlindReceiving)
        {
            Checklist3Text = $"3. Carátula del Cierre Oficial (Hoja 'CIERRE RECEP') — Montaje a ciegas preliminar ({KpiDeclaredUnits:N0} UN Factura | Físico: 0 UN).";
            Checklist4Text = "4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: SUSPENDIDO (MONTAJE A CIEGAS, PENDIENTE CONTEO).";
            Checklist6Text = "6. Lista de Empaque de Bultos y Unidades (Hoja 'Bultos') — Estado: PENDIENTE (MONTAJE A CIEGAS, SIN CONTEO FÍSICO).";
            Checklist6Bultos = false;
        }

        RecalculateChecklistProgress();
    }

    // ==========================================
    // PIPELINE EXECUTION & DATA BINDING
    // ==========================================
    [RelayCommand]
    public void BrowseQcFile()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Archivos de Control de Calidad (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Todos los archivos (*.*)|*.*",
            Title = "Seleccionar libro de Control de Calidad (QC)"
        };
        if (dlg.ShowDialog() == true)
        {
            QcFilePath = dlg.FileName;
            StatusMessage = $"QC seleccionado: {Path.GetFileName(dlg.FileName)}";
        }
    }

    [RelayCommand]
    public void BrowseClosureTemplate()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Plantillas de Cierre (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Todos los archivos (*.*)|*.*",
            Title = "Seleccionar plantilla corporativa de Cierre de Recepción"
        };
        if (dlg.ShowDialog() == true)
        {
            ClosureTemplatePath = dlg.FileName;
            StatusMessage = $"Plantilla seleccionada: {Path.GetFileName(dlg.FileName)}";
        }
    }

    [RelayCommand]
    public void ExecuteReconciliation()
    {
        bool hasQc = !string.IsNullOrWhiteSpace(QcFilePath) && File.Exists(QcFilePath);
        bool hasClosure = !string.IsNullOrWhiteSpace(ClosureTemplatePath) && File.Exists(ClosureTemplatePath);

        if (!hasClosure && !hasQc)
        {
            StatusMessage = "Seleccione al menos una plantilla de Cierre (y opcionalmente el archivo de Control de Calidad QC).";
            return;
        }

        IsBusy = true;
        StatusMessage = "Ejecutando conciliación 3-Way en streaming...";

        try
        {
            var options = new ReceivingPipelineOptions
            {
                QcFilePath = QcFilePath,
                ClosureTemplatePath = ClosureTemplatePath,
                SafetyGuard = _safetyGuard,
                ExcludeZeroTotalFromStripsAndBultos = HideZeroTotalItems
            };

            ReceivingPipelineResult result = ReceivingPipeline.Execute(options);
            LastReconciliationResult = result;

            PopulateResults(result);
            UpdateChecklistDynamicLabels();

            if (result.Comparison.IsBlindReceiving)
            {
                StatusMessage = $"AVISO: Montaje a ciegas detectado (Facturado: {result.Comparison.TotalDeclaredUnits:N0} UN | Físico: 0 UN). Cierre preliminar en espera de descarga física.";
            }
            else
            {
                StatusMessage = $"Conciliación exitosa: Contenedor {result.ContainerId} | {result.Comparison.ActiveStylesCount} estilos activos procesados.";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error durante la conciliación: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void PopulateResults(ReceivingPipelineResult result)
    {
        Tira1AsAtText = result.Tira1AsAtTsv;
        Tira2AwAyText = result.Tira2AwAyTsv;

        Discrepancies.Clear();
        PrepacksWithBom.Clear();
        foreach (var item in result.Comparison.Items)
        {
            Discrepancies.Add(item);
            if (item.Children.Count > 0)
            {
                PrepacksWithBom.Add(item);
            }
        }
        SelectedBomPrepack = PrepacksWithBom.FirstOrDefault();

        BultosItems.Clear();
        foreach (var b in result.BultosItems)
        {
            BultosItems.Add(b);
        }

        TotalBultosDescarga = BultosItems.Sum(b => b.PhysicalCartons);
        TotalBultosPl = BultosItems.Sum(b => b.DeclaredCartons);
        DifBultos = TotalBultosDescarga - TotalBultosPl;
        TotalBultosEnteros = BultosItems.Sum(b => b.FullCartons);
        TotalBultosComplementos = BultosItems.Sum(b => b.ComplementCartons);

        ApplyDiscrepancyFilter();
        ApplyBultosFilter();
        ApplyInspectorFilter();

        // Update Executive KPI Strip
        KpiContainerId = result.ContainerId;
        KpiPoNumber = result.DeclaredItems.FirstOrDefault(d => !string.IsNullOrWhiteSpace(d.OrderPo43))?.OrderPo43 ?? "-";
        KpiPedidoNumber = result.Comparison.Items.FirstOrDefault()?.OrderPo43 ?? "-";
        KpiTotalStyles = result.Comparison.ActiveStylesCount;
        KpiCleanStyles = result.Comparison.CleanMaterialsCount;
        KpiDiscrepantStyles = result.Comparison.DiscrepantMaterialsCount;
        KpiDeclaredUnits = result.Comparison.TotalDeclaredUnits;
        KpiPhysicalUnits = result.Comparison.TotalPhysicalUnits;
        KpiDeltaUnits = result.Comparison.NetUnitDifference;
        KpiDeclaredCartons = result.Comparison.TotalDeclaredCartons;
        KpiPhysicalCartons = result.Comparison.TotalPhysicalCartons;
        KpiDeltaCartons = result.Comparison.NetCartonDifference;
        KpiRevueltoPairs = result.Comparison.RevueltoPairsCount;

        IsBlindReceiving = result.Comparison.IsBlindReceiving;
        KpiDamagedUnits = result.Comparison.TotalDamagedUnits;

        ContainerId = result.ContainerId;
        Po43Number = KpiPoNumber;
    }

    // ==========================================
    // VALIDADOR SAP (PORTAPAPELES)
    // ==========================================
    [RelayCommand]
    public void LoadSapSample()
    {
        SapRawPasteText = "Material\tDescripcion\tCantidad\tBultos\n" +
                          "000000001455102001\tCAMISETA ALGODON BASICA\t600.00\t10\n" +
                          "1.455102002E+09\tJEANS SLIM FIT INDIGO\t1,200\t20\n" +
                          "1455102003.0\tBLUSA POPELINA BLANCA\t(48)\t1\n" +
                          "1E2\tETIQUETA ALFANUMERICA\t$-50.00\t0";
        StatusMessage = "Datos de muestra de SAP cargados.";
    }

    [RelayCommand]
    public void SanitizeSapClipboard()
    {
        SapPasteError = "";
        if (string.IsNullOrWhiteSpace(SapRawPasteText))
        {
            SapPasteError = "El cuadro de texto está vacío. Pegue datos de SAP antes de validar.";
            return;
        }

        try
        {
            SanitizedTable = SapClipboardSanitizer.SanitizeSapTable(SapRawPasteText);
            SapSanitizedHeaders.Clear();
            foreach (var h in SanitizedTable.Headers)
                SapSanitizedHeaders.Add(h);

            SapSanitizedRows.Clear();
            foreach (var r in SanitizedTable.Rows)
                SapSanitizedRows.Add(r);

            SapDelimiterLabel = SanitizedTable.DelimiterLabel;
            StatusMessage = $"Tabla SAP sanitizada: {SanitizedTable.RowCount} filas limpias ({SanitizedTable.DelimiterLabel}).";
        }
        catch (Exception ex)
        {
            SapPasteError = $"Error al procesar: {ex.Message}";
        }
    }

    [RelayCommand]
    public void CopyCleanTsv()
    {
        if (SanitizedTable == null || string.IsNullOrWhiteSpace(SanitizedTable.CleanTsv))
        {
            StatusMessage = "No hay tabla sanitizada para copiar.";
            return;
        }

        try
        {
            Clipboard.SetText(SanitizedTable.CleanTsv);
            StatusMessage = "TSV limpio copiado al portapapeles. Listo para pegar en Excel o SAP.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error al copiar: {ex.Message}";
        }
    }

    // ==========================================
    // EXPORT & CLIPBOARD ACTIONS
    // ==========================================
    [RelayCommand]
    public void CopyTira1AsAt()
    {
        if (string.IsNullOrWhiteSpace(Tira1AsAtText))
        {
            StatusMessage = "No hay datos de Tira 1 para copiar. Ejecute la conciliación primero.";
            return;
        }

        try
        {
            Clipboard.SetText(Tira1AsAtText);
            StatusMessage = "Tira 1 (AS:AT - Bultos y Unidades Físicas) copiada al portapapeles.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error al copiar al portapapeles: {ex.Message}";
        }
    }

    [RelayCommand]
    public void CopyTira2AwAy()
    {
        if (string.IsNullOrWhiteSpace(Tira2AwAyText))
        {
            StatusMessage = "No hay datos de Tira 2 para copiar. Ejecute la conciliación primero.";
            return;
        }

        try
        {
            Clipboard.SetText(Tira2AwAyText);
            StatusMessage = "Tira 2 (AW:AY - Packing List y Empaque) copiada al portapapeles.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error al copiar al portapapeles: {ex.Message}";
        }
    }

    [RelayCommand]
    public void ExportMacroSafeXlsm()
    {
        if (LastReconciliationResult == null)
        {
            StatusMessage = "Ejecute la conciliación primero antes de exportar.";
            return;
        }

        var dlg = new SaveFileDialog
        {
            Filter = "Libro de Excel habilitado para macros (*.xlsm)|*.xlsm",
            FileName = $"{KpiContainerId}_Cierre_Oficial_Cuadrado.xlsm",
            Title = "Guardar Cierre Oficial Macro-Safe (.xlsm)"
        };

        if (dlg.ShowDialog() == true)
        {
            try
            {
                var exportRes = ClosureExporter.Export(
                    ClosureTemplatePath,
                    dlg.FileName,
                    LastReconciliationResult.Comparison,
                    LastReconciliationResult.QcReport,
                    _safetyGuard);

                XlsmVbaSha256Result = exportRes.Validation.OutputHashes?.VbaProjectSha256 ?? "N/A";
                XlsmSha256Preserved = exportRes.Validation.VbaProjectHashPreserved;

                StatusMessage = $"Exportación Macro-Safe completada: {Path.GetFileName(dlg.FileName)} (SHA-256 VBA Preservado: {XlsmSha256Preserved}).";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error al exportar: {ex.Message}";
            }
        }
    }

    [RelayCommand]
    public void OpenDiagnosticGuideForDiscrepancy(DiscrepancyItem item)
    {
        if (item is null) return;

        DiagnosticGuide? matchedGuide = null;
        if (item.DiscrepancyType is DiscrepancyType.Shortage)
        {
            matchedGuide = DiagnosticGuides.FirstOrDefault(g => g.GuideId.Contains("FALTANTE", StringComparison.OrdinalIgnoreCase));
        }
        else if (item.DiscrepancyType is DiscrepancyType.Overage)
        {
            matchedGuide = DiagnosticGuides.FirstOrDefault(g => g.GuideId.Contains("SOBRANTE", StringComparison.OrdinalIgnoreCase));
        }
        else if (item.DiscrepancyType is DiscrepancyType.RevueltoCandidate)
        {
            matchedGuide = DiagnosticGuides.FirstOrDefault(g => g.GuideId.Contains("REVUELTO", StringComparison.OrdinalIgnoreCase));
        }

        if (matchedGuide is not null)
        {
            SelectedDiagnosticGuide = matchedGuide;
            SelectedTabIndex = 7; // Jump directly to SAP & Wiki Tab (Index 7)
            WikiSubTab = 1;      // Jump to Guías tab inside Wiki
            StatusMessage = $"Abriendo guía de diagnóstico para discrepancia '{item.DiscrepancyType}' del material {item.MaterialCode}.";
        }
        else
        {
            StatusMessage = $"No se encontró una guía específica para el tipo de discrepancia '{item.DiscrepancyType}'.";
        }
    }

    [RelayCommand]
    private void SearchSap()
    {
        FilteredSapTransactions.Clear();
        string q = (SapSearchQuery ?? "").Trim().ToLowerInvariant();

        foreach (var t in DefaultSapCatalog.Transactions)
        {
            if (string.IsNullOrWhiteSpace(q) ||
                t.Code.ToLowerInvariant().Contains(q) ||
                t.Title.ToLowerInvariant().Contains(q) ||
                t.Purpose.ToLowerInvariant().Contains(q))
            {
                FilteredSapTransactions.Add(t);
            }
        }
    }

    // ==========================================
    // DAILY LOG & FORMULAS (BACKWARD COMPAT)
    // ==========================================
    partial void OnStage1CustomsChanged(CustomsDuaStage value) => RecalculateDailyLogFormulas();
    partial void OnStage2SapIntakeChanged(SapIntakeStage value) => RecalculateDailyLogFormulas();
    partial void OnStage3BlindSetupChanged(BlindSetupStage value) => RecalculateDailyLogFormulas();
    partial void OnStage4QualityControlChanged(QualityControlStage value) => RecalculateDailyLogFormulas();
    partial void OnStage5ReceptionClosureChanged(ReceptionClosureStage value) => RecalculateDailyLogFormulas();
    partial void OnStage6DispatchClosureChanged(DispatchClosureStage value) => RecalculateDailyLogFormulas();
    partial void OnSelectedSaludoChanged(Saludo value) => RecalculateDailyLogFormulas();
    partial void OnContainerIdChanged(string value) => RecalculateDailyLogFormulas();
    partial void OnEntrega18NumberChanged(string value) => RecalculateDailyLogFormulas();

    public void RecalculateDailyLogFormulas()
    {
        var entry = new DailyLogEntry
        {
            ContainerId = ContainerId,
            Po43Numbers = string.IsNullOrWhiteSpace(Po43Number) ? Array.Empty<string>() : new[] { Po43Number },
            Pedido41 = string.IsNullOrWhiteSpace(Pedido41Number) ? null : Pedido41Number,
            EntregaVl31n = string.IsNullOrWhiteSpace(Entrega18Number) ? null : Entrega18Number,
            Greeting = SelectedSaludo,
            Stage1Customs = Stage1Customs,
            Stage2SapIntake = Stage2SapIntake,
            Stage3BlindSetup = Stage3BlindSetup,
            Stage4QualityControl = Stage4QualityControl,
            Stage5ReceptionClosure = Stage5ReceptionClosure,
            Stage6DispatchClosure = Stage6DispatchClosure
        };

        MasterContainerState state = DailyLogLifecycleValidator.ComputeMasterState(entry);
        ComputedColumnO = DailyLogCatalogs.ToExcelString(state);

        if (!string.IsNullOrWhiteSpace(Entrega18Number) && Entrega18Number.Trim() != "-")
        {
            string saludoText = DailyLogCatalogs.ToExcelString(SelectedSaludo);
            if (saludoText == "-") saludoText = "Buenos días";
            string rawContainer = string.IsNullOrWhiteSpace(ContainerId) ? "[CONTENEDOR]" : ContainerId.Replace("IMP-", "", StringComparison.OrdinalIgnoreCase).Trim();
            ComputedColumnR = $"{saludoText}, Se informa que el ingreso a nivel de sistema del Contenedor {rawContainer}, quedo listo.";
        }
        else
        {
            ComputedColumnR = "-";
        }
    }

    [RelayCommand]
    public void GenerateSapClipboardNotification()
    {
        RecalculateDailyLogFormulas();

        if (string.IsNullOrWhiteSpace(ContainerId))
        {
            StatusMessage = "Indique al menos el número de contenedor.";
            return;
        }

        ClipboardSummary = $"*RECEPCIÓN CONTENEDOR {ContainerId}*\n" +
                           $"PO: {Po43Number}\n" +
                           $"Pedido: {Pedido41Number}\n" +
                           $"Entrega: {Entrega18Number}\n" +
                           $"Tarea EWM: {EwmTaskNumber}\n" +
                           $"HU: {HandlingUnit}\n" +
                           $"Etapa Actual (Col O): {ComputedColumnO}\n" +
                           $"Notificación SAP (Col R): {ComputedColumnR}";

        try
        {
            Clipboard.SetText(ClipboardSummary);
            StatusMessage = "Resumen y notificación SAP copiados al portapapeles.";
        }
        catch
        {
            StatusMessage = "Resumen generado con éxito.";
        }
    }

    private void InitializeSopManualText()
    {
        SopManualText =
@"MANUAL OPERATIVO Y PROTOCOLO DE CONCILIACIÓN DE RECEPCIONES (SOP ALMACEMA C001)

1. REGLA CARDINAL DE DISCREPANCIAS
- Las unidades faltantes y las unidades sobrantes NUNCA se compensan entre sí algebraicamente.
- Un faltante de 60 unidades y un sobrante de 48 unidades generan un faltante bruto de 60 UN y un sobrante de 48 UN (no un faltante neto de 12 UN).
- La Nota de Crédito (NCR) se calcula exclusivamente sobre faltantes o mercancía con daño físico real.

2. REGLA DEL BULTO CERO Y MATERIALES SIN MOVIMIENTO
- Cuando un material tiene 0 unidades pedidas, 0 recibidas y 0 bultos en ambas fuentes, se clasifica como 'SIN MOVIMIENTO (0 TOTAL)'.
- Al generar las tiras para pegado en Excel (AS:AT y AW:AY) y la hoja 'Bultos', los materiales en cero se omiten por defecto para evitar desfases de celdas.

3. EMPAQUE CANÓNICO Y UNIDADES SUELTAS
- Formato estándar compuesto: 'NxM KxP' (ej. '8*48 1*36' para 8 bultos enteros de 48 y 1 bulto complementario de 36).
- Si existen unidades complementarias pero no hay bulto adicional (Layout #4 o muelle), se representa como '8*48 +36' (unidades sueltas sin deducir un cartón fantasma).

4. PROTOCOLO DE CANDIDATOS A REVUELTO
- Dos materiales con discrepancias opuestas de idéntica magnitud se evalúan por:
  a) Mismo proveedor o estilo similar.
  b) Misma categoría y discrepancia opuesta en bultos (Confirmed).
  c) Discrepancia opuesta en bultos o misma categoría (Probable).
  d) Magnitud idéntica de unidades sin otros factores (Possible).

5. EXPORTACIÓN MACRO-SAFE (.XLSM)
- Toda exportación a archivo .xlsm preserva el bloque binario xl/vbaProject.bin y valida su hash SHA-256 antes y después de escribir datos.";
    }
}
