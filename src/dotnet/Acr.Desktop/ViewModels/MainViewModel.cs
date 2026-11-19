using Acr.Core.Models;
using Acr.Core.SapKnowledge;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;
using Acr.Infrastructure.Safety;
using Acr.Infrastructure.SapKnowledge;
using Acr.Infrastructure.Sync;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace Acr.Desktop.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly OperationalSafetyGuard _safetyGuard;

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

    // Daily Log properties
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
    private string _clipboardSummary = "";

    // SAP Knowledge
    [ObservableProperty]
    private string _sapSearchQuery = "";

    [ObservableProperty]
    private ObservableCollection<TransactionEntry> _filteredSapTransactions = new();

    [ObservableProperty]
    private TransactionEntry? _selectedSapTransaction;

    [ObservableProperty]
    private ObservableCollection<DiagnosticGuide> _diagnosticGuides = new();

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
    private async Task RunReconciliationAsync()
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
    private void GenerateSapClipboardNotification()
    {
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
                           $"Estado: Descarga finalizada / Listo para revisión.";

        try
        {
            Clipboard.SetText(ClipboardSummary);
            StatusMessage = "Resumen copiado al portapapeles para notificación SAP.";
        }
        catch
        {
            StatusMessage = "Resumen generado (no se pudo acceder directamente al portapapeles).";
        }
    }
}
