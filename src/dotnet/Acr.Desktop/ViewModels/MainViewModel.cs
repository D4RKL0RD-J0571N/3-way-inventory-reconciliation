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

    // --- RECONCILIATION & CLOSURE ---
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
    private ReceivingPipelineResult? _lastReconciliationResult;

    [ObservableProperty]
    private ObservableCollection<DiscrepancyItem> _discrepancies = new();

    [ObservableProperty]
    private DiscrepancyItem? _selectedDiscrepancy;

    // --- DAILY LOG 6-STAGES & FORMULAS ---
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

    // Catalog lists for UI ComboBoxes
    public IReadOnlyList<Saludo> SaludoOptions { get; } = Enum.GetValues<Saludo>();
    public IReadOnlyList<CustomsDuaStage> Stage1Options { get; } = Enum.GetValues<CustomsDuaStage>();
    public IReadOnlyList<SapIntakeStage> Stage2Options { get; } = Enum.GetValues<SapIntakeStage>();
    public IReadOnlyList<BlindSetupStage> Stage3Options { get; } = Enum.GetValues<BlindSetupStage>();
    public IReadOnlyList<QualityControlStage> Stage4Options { get; } = Enum.GetValues<QualityControlStage>();
    public IReadOnlyList<ReceptionClosureStage> Stage5Options { get; } = Enum.GetValues<ReceptionClosureStage>();
    public IReadOnlyList<DispatchClosureStage> Stage6Options { get; } = Enum.GetValues<DispatchClosureStage>();

    // --- SAP KNOWLEDGE ---
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

        RecalculateDailyLogFormulas();
    }

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

        // Col R formula reproduction: IF(F<>"-" and not empty, Q & ", Se informa que el ingreso a nivel de sistema del Contenedor " & SUBSTITUTE(B, "IMP-", "") & ", quedo listo.", "-")
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
    public void BrowseQcFile()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Seleccionar Archivo de Control de Calidad (QC)",
            Filter = "Archivos de Excel habilitados para macros (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Todos los archivos (*.*)|*.*",
            CheckFileExists = true
        };

        if (dlg.ShowDialog() == true)
        {
            QcFilePath = dlg.FileName;
        }
    }

    [RelayCommand]
    public void BrowseClosureTemplate()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Seleccionar Plantilla de Cierre Corporativo",
            Filter = "Archivos de Cierre (*.xlsm)|*.xlsm|Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*",
            CheckFileExists = true
        };

        if (dlg.ShowDialog() == true)
        {
            ClosureTemplatePath = dlg.FileName;
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
            SelectedTabIndex = 1; // Jump directly to SAP Knowledge Tab (Tab 2)
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
        string q = SapSearchQuery.Trim().ToLowerInvariant();

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

    [RelayCommand]
    public async Task RunReconciliationAsync()
    {
        if (string.IsNullOrWhiteSpace(QcFilePath) || !File.Exists(QcFilePath))
        {
            StatusMessage = "Error: Seleccione un archivo de Control de Calidad válido.";
            return;
        }

        if (string.IsNullOrWhiteSpace(ClosureTemplatePath) || !File.Exists(ClosureTemplatePath))
        {
            StatusMessage = "Error: Seleccione una plantilla de Cierre válida.";
            return;
        }

        IsBusy = true;
        StatusMessage = "Procesando reconciliación y validando estructura...";

        try
        {
            string outDir = Path.GetDirectoryName(ClosureTemplatePath) ?? Environment.CurrentDirectory;
            string exportPath = Path.Combine(outDir, $"CIERRE_{Path.GetFileNameWithoutExtension(QcFilePath)}.xlsm");

            var options = new ReceivingPipelineOptions
            {
                QcFilePath = QcFilePath,
                ClosureTemplatePath = ClosureTemplatePath,
                ExportOutputPath = exportPath,
                SafetyGuard = _safetyGuard
            };

            ReceivingPipelineResult result = await Task.Run(() => ReceivingPipeline.Execute(options));
            LastReconciliationResult = result;

            Discrepancies.Clear();
            foreach (var item in result.Comparison.Items)
            {
                Discrepancies.Add(item);
            }

            StatusMessage = $"Reconciliación exitosa. Total líneas: {result.Comparison.Items.Count}, Discrepancias: {result.Comparison.DiscrepantMaterialsCount}. Guardado en: {exportPath}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error durante la reconciliación: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
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
}
