using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using Acr.Core.Models;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;

namespace Acr.Desktop;

/// <summary>Operator review surface for matching supplier list rows and accepting declared fields.</summary>
public sealed class SupplierPackingListReviewWindow : Window
{
    private readonly DataGrid _grid;
    private readonly ObservableCollection<PackingListReviewRow> _rows;
    private readonly string _sourceFile;
    private readonly HashSet<string> _validTargetCodes;

    public IReadOnlyDictionary<string, SupplierPackingListOverride>? AcceptedOverrides { get; private set; }

    public SupplierPackingListReviewWindow(
        string sourceFile,
        IReadOnlyList<SupplierPackingListRow> supplierRows,
        IReadOnlyList<DeclaredItem> declaredItems,
        IReadOnlyList<QcInspectionItem> qcItems)
    {
        _sourceFile = sourceFile;
        Title = "Revisar Packing List del proveedor";
        Width = 1200;
        Height = 600;
        MinWidth = 900;
        MinHeight = 420;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.CanResize;

        var declaredByCode = declaredItems
            .GroupBy(x => NormalizeMaterialCode(x.MaterialCode), StringComparer.OrdinalIgnoreCase)
            .Where(x => x.Key.Length > 0)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
        _validTargetCodes = declaredByCode.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var declaredByVendorCode = declaredItems
            .Where(x => !string.IsNullOrWhiteSpace(x.VendorMaterial))
            .GroupBy(x => NormalizeMaterialCode(x.VendorMaterial!), StringComparer.OrdinalIgnoreCase)
            .Where(x => x.Key.Length > 0)
            .ToDictionary(
                x => x.Key,
                x => x.GroupBy(item => NormalizeMaterialCode(item.MaterialCode), StringComparer.OrdinalIgnoreCase)
                    .Select(items => items.First())
                    .ToList(),
                StringComparer.OrdinalIgnoreCase);
        var qcByCode = qcItems.Where(x => !x.IsChildSku)
            .GroupBy(x => NormalizeMaterialCode(x.MaterialCode), StringComparer.OrdinalIgnoreCase)
            .Where(x => x.Key.Length > 0)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

        _rows = new ObservableCollection<PackingListReviewRow>(supplierRows
            .GroupBy(x => NormalizeMaterialCode(x.MaterialCode), StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                string key = group.Key;
                DeclaredItem? target = null;
                string matchType = "";
                IReadOnlyList<DeclaredItem> vendorCandidates = Array.Empty<DeclaredItem>();
                bool matched = declaredByCode.TryGetValue(key, out target);
                if (matched)
                {
                    matchType = "Código SAP";
                }
                else if (declaredByVendorCode.TryGetValue(key, out List<DeclaredItem>? mappedItems))
                {
                    vendorCandidates = mappedItems;
                    if (mappedItems.Count == 1)
                    {
                        target = mappedItems[0];
                        matched = true;
                        matchType = "MAT. PROVEEDOR";
                    }
                    else
                    {
                        matchType = "MAT. PROVEEDOR ambiguo";
                    }
                }
                string qcMatchKey = target is not null ? NormalizeMaterialCode(target.MaterialCode) : key;
                qcByCode.TryGetValue(qcMatchKey, out QcInspectionItem? qc);
                return new PackingListReviewRow
                {
                    SupplierCode = group.First().MaterialCode,
                    Description = group.First().Description,
                    MatchType = matchType,
                    SourceLocation = string.Join(", ", group.Select(x => x.SourceLocation)),
                    SourceRows = group.ToList(),
                    Units = group.Any(x => x.Units.HasValue) ? group.Sum(x => x.Units ?? 0) : null,
                    Cartons = group.Any(x => x.Cartons.HasValue) ? group.Sum(x => x.Cartons ?? 0) : null,
                    ConfidencePercent = (int)Math.Round(group.Min(x => x.Confidence) * 100),
                    TargetMaterialCode = matched ? target!.MaterialCode : "",
                    SuggestedMaterialCode = matched ? "" : vendorCandidates.Count > 1
                        ? string.Join(", ", vendorCandidates.Select(x => x.MaterialCode))
                        : FindDescriptionCandidate(group.First().Description, declaredItems),
                    ClosureUnits = target?.PackingListUnits ?? target?.DeclaredUnits,
                    ClosureCartons = target?.DeclaredCartons,
                    QcUnits = qc?.CalculationResult?.CalculatedUnits,
                    QcCartons = qc?.CalculationResult?.CalculatedCartons,
                    AcceptUnits = matched && group.Any(x => x.Units.HasValue) && group.Where(x => x.Units.HasValue).All(x => x.Confidence >= 0.85),
                    AcceptCartons = matched && group.Any(x => x.Cartons.HasValue) && group.Where(x => x.Cartons.HasValue).All(x => x.Confidence >= 0.85)
                };
            }).ToList());

        var root = new DockPanel { Margin = new Thickness(12) };
        var title = new TextBlock
        {
            Text = $"Documento: {Path.GetFileName(sourceFile)} — Revise la coincidencia y marque los campos que desea usar.",
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 10),
            TextWrapping = TextWrapping.Wrap,
            ToolTip = "Archivo proveedor que se está revisando. Los valores solo se aplican después de confirmar."
        };
        DockPanel.SetDock(title, Dock.Top);
        root.Children.Add(title);

        var bulkSelection = new WrapPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
        AddBulkButton(bulkSelection, "Marcar todas las unidades", "Selecciona las unidades disponibles en todas las filas, incluso las que tengan confianza OCR baja. Revise esos valores antes de aplicar.", () => SetAll(row => row.AcceptUnits = row.Units.HasValue));
        AddBulkButton(bulkSelection, "Marcar todos los bultos", "Selecciona los bultos disponibles en todas las filas, incluso las que tengan confianza OCR baja. Revise esos valores antes de aplicar.", () => SetAll(row => row.AcceptCartons = row.Cartons.HasValue));
        AddBulkButton(bulkSelection, "Marcar todo", "Selecciona unidades y bultos disponibles en todas las filas, incluso valores OCR de baja confianza. Revise antes de aplicar.", () => SetAll(row =>
        {
            row.AcceptUnits = row.Units.HasValue;
            row.AcceptCartons = row.Cartons.HasValue;
        }));
        AddBulkButton(bulkSelection, "Desmarcar todo", "Quita la selección de unidades y bultos en todas las filas.", () => SetAll(row =>
        {
            row.AcceptUnits = false;
            row.AcceptCartons = false;
        }));
        DockPanel.SetDock(bulkSelection, Dock.Top);
        root.Children.Add(bulkSelection);

        _grid = new DataGrid
        {
            ItemsSource = _rows,
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            IsReadOnly = false,
            SelectionUnit = DataGridSelectionUnit.FullRow,
            FrozenColumnCount = 2,
            Margin = new Thickness(0, 0, 0, 10),
            ToolTip = "Revise los valores detectados, compare con el cierre y QC, y seleccione qué campos importar."
        };
        AddTextColumn("Código proveedor", nameof(PackingListReviewRow.SupplierCode), 125, true, "Código o referencia del artículo detectado en el documento del proveedor.");
        AddTextColumn("Descripción", nameof(PackingListReviewRow.Description), 190, true, "Descripción del producto detectada en el documento del proveedor.");
        AddTextColumn("Tipo coincidencia", nameof(PackingListReviewRow.MatchType), 115, true, "Indica si el código coincidió con el material SAP o con MAT. PROVEEDOR de la plantilla.");
        AddTextColumn("Ubicación fuente", nameof(PackingListReviewRow.SourceLocation), 155, true, "Página, hoja y fila de donde se extrajeron los valores.");
        AddTextColumn("Unidades PL", nameof(PackingListReviewRow.Units), 85, false, "Unidades declaradas por el proveedor. Puede corregir este valor antes de importarlo.");
        AddTextColumn("Bultos PL", nameof(PackingListReviewRow.Cartons), 75, false, "Bultos declarados por el proveedor. Puede corregir este valor antes de importarlo.");
        AddTextColumn("Unidades Cierre", nameof(PackingListReviewRow.ClosureUnits), 95, true, "Unidades de Packing List que ya están en la plantilla de cierre.");
        AddTextColumn("Bultos Cierre", nameof(PackingListReviewRow.ClosureCartons), 90, true, "Bultos de Packing List que ya están en la plantilla de cierre.");
        AddTextColumn("Unidades QC", nameof(PackingListReviewRow.QcUnits), 85, true, "Unidades físicas calculadas desde Control de Calidad. Si no se cargó QC, queda vacío.");
        AddTextColumn("Bultos QC", nameof(PackingListReviewRow.QcCartons), 75, true, "Bultos físicos calculados desde Control de Calidad. Si no se cargó QC, queda vacío.");
        AddTextColumn("OCR %", nameof(PackingListReviewRow.ConfidencePercent), 65, true, "Confianza promedio del reconocimiento OCR. Los valores menores a 85% quedan desmarcados.");
        AddTextColumn("PL − QC unidades", nameof(PackingListReviewRow.UnitsDifferenceFromQc), 105, true, "Diferencia entre unidades del proveedor y unidades físicas de QC.");
        AddTextColumn("PL − QC bultos", nameof(PackingListReviewRow.CartonsDifferenceFromQc), 95, true, "Diferencia entre bultos del proveedor y bultos físicos de QC.");
        AddTextColumn("Sugerencia", nameof(PackingListReviewRow.SuggestedMaterialCode), 110, true, "Coincidencia sugerida por descripción. Verifique el código antes de usarlo.");
        AddTextColumn("Material de cierre (escriba para resolver)", nameof(PackingListReviewRow.TargetMaterialCode), 210, false, "Código de material existente en la plantilla al que se aplicarán los valores seleccionados.");
        AddCheckColumn("Usar unidades", nameof(PackingListReviewRow.AcceptUnits), 85, "Marque para reemplazar las unidades de Packing List de este material con el valor del proveedor.");
        AddCheckColumn("Usar bultos", nameof(PackingListReviewRow.AcceptCartons), 80, "Marque para reemplazar los bultos de Packing List de este material con el valor del proveedor.");
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var apply = new Button { Content = "Aplicar valores seleccionados", MinWidth = 190, Padding = new Thickness(12, 7, 12, 7), Margin = new Thickness(0, 0, 8, 0), IsDefault = true,
            ToolTip = "Confirma los campos seleccionados y los usa en esta conciliación y en la próxima exportación." };
        apply.Click += (_, _) => ApplyAndClose();
        var cancel = new Button { Content = "Cancelar", MinWidth = 90, Padding = new Thickness(12, 7, 12, 7), IsCancel = true,
            ToolTip = "Cierra la revisión y descarta los valores del proveedor para esta conciliación." };
        buttons.Children.Add(apply);
        buttons.Children.Add(cancel);
        DockPanel.SetDock(buttons, Dock.Bottom);
        root.Children.Add(buttons);
        root.Children.Add(_grid);
        Content = root;
    }

    private void AddBulkButton(Panel panel, string label, string tooltip, Action action)
    {
        var button = new Button
        {
            Content = label,
            Padding = new Thickness(9, 5, 9, 5),
            Margin = new Thickness(0, 0, 6, 0),
            ToolTip = tooltip
        };
        button.Click += (_, _) => action();
        panel.Children.Add(button);
    }

    private void SetAll(Action<PackingListReviewRow> update)
    {
        _grid.CommitEdit(DataGridEditingUnit.Cell, true);
        _grid.CommitEdit(DataGridEditingUnit.Row, true);
        foreach (PackingListReviewRow row in _rows)
            update(row);
    }

    private void AddTextColumn(string header, string property, double width, bool readOnly, string tooltip)
    {
        var headerStyle = new Style(typeof(DataGridColumnHeader));
        headerStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, tooltip));
        var cellStyle = new Style(typeof(DataGridCell));
        cellStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,
            new Binding(property) { StringFormat = $"{header}: {{0}}" }));
        _grid.Columns.Add(new DataGridTextColumn
        {
            Header = header,
            Binding = new Binding(property) { UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
            Width = new DataGridLength(width),
            IsReadOnly = readOnly,
            HeaderStyle = headerStyle,
            CellStyle = cellStyle
        });
    }

    private void AddCheckColumn(string header, string property, double width, string tooltip)
    {
        var headerStyle = new Style(typeof(DataGridColumnHeader));
        headerStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, tooltip));
        var cellStyle = new Style(typeof(DataGridCell));
        cellStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, tooltip));
        _grid.Columns.Add(new DataGridCheckBoxColumn
        {
            Header = header,
            Binding = new Binding(property) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
            Width = new DataGridLength(width),
            HeaderStyle = headerStyle,
            CellStyle = cellStyle
        });
    }

    private void ApplyAndClose()
    {
        _grid.CommitEdit(DataGridEditingUnit.Cell, true);
        _grid.CommitEdit(DataGridEditingUnit.Row, true);

        bool hasInvalidAcceptedTarget = _rows.Any(row =>
            (row.AcceptUnits || row.AcceptCartons) &&
            !_validTargetCodes.Contains(NormalizeMaterialCode(row.TargetMaterialCode)));
        if (hasInvalidAcceptedTarget)
        {
            MessageBox.Show(this,
                "Cada fila marcada para importar debe tener un Material de cierre que exista en la plantilla. Revise el código destino o desmarque los campos.",
                "Material destino no válido",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var chosen = _rows
            .Where(row => !string.IsNullOrWhiteSpace(row.TargetMaterialCode) &&
                          _validTargetCodes.Contains(NormalizeMaterialCode(row.TargetMaterialCode)) &&
                          ((row.AcceptUnits && row.Units.HasValue) || (row.AcceptCartons && row.Cartons.HasValue)))
            .GroupBy(row => NormalizeMaterialCode(row.TargetMaterialCode), StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Key.Length > 0);

        var overrides = new Dictionary<string, SupplierPackingListOverride>(StringComparer.OrdinalIgnoreCase);
        foreach (var group in chosen)
        {
            int? units = group.Any(row => row.AcceptUnits && row.Units.HasValue)
                ? group.Where(row => row.AcceptUnits && row.Units.HasValue).Sum(row => row.Units!.Value)
                : null;
            int? cartons = group.Any(row => row.AcceptCartons && row.Cartons.HasValue)
                ? group.Where(row => row.AcceptCartons && row.Cartons.HasValue).Sum(row => row.Cartons!.Value)
                : null;
            string unitsLocation = string.Join(", ", group
                .Where(row => row.AcceptUnits && row.Units.HasValue)
                .SelectMany(row => row.SourceRows.Where(source => source.Units.HasValue).Select(source => source.SourceLocation))
                .Distinct());
            string cartonsLocation = string.Join(", ", group
                .Where(row => row.AcceptCartons && row.Cartons.HasValue)
                .SelectMany(row => row.SourceRows.Where(source => source.Cartons.HasValue).Select(source => source.SourceLocation))
                .Distinct());
            overrides[group.Key] = new SupplierPackingListOverride(units, cartons, _sourceFile, unitsLocation, cartonsLocation);
        }

        AcceptedOverrides = overrides;
        DialogResult = true;
    }

    private static string FindDescriptionCandidate(string description, IReadOnlyList<DeclaredItem> declaredItems)
    {
        string source = NormalizeWords(description);
        if (source.Length == 0) return "";
        var sourceWords = source.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet(StringComparer.Ordinal);
        string bestCode = "";
        double bestScore = 0;
        foreach (DeclaredItem item in declaredItems)
        {
            string target = NormalizeWords(item.Description ?? "");
            var targetWords = target.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet(StringComparer.Ordinal);
            if (targetWords.Count == 0) continue;
            double score = sourceWords.Intersect(targetWords).Count() / (double)sourceWords.Union(targetWords).Count();
            if (score > bestScore)
            {
                bestScore = score;
                bestCode = item.MaterialCode;
            }
        }
        return bestScore >= 0.35 ? bestCode : "";
    }

    private static string NormalizeWords(string value)
    {
        string decomposed = value.Normalize(System.Text.NormalizationForm.FormD);
        var chars = decomposed.Where(c => System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) !=
                                         System.Globalization.UnicodeCategory.NonSpacingMark)
            .Select(c => char.IsLetterOrDigit(c) ? char.ToUpperInvariant(c) : ' ');
        return System.Text.RegularExpressions.Regex.Replace(new string(chars.ToArray()), @"\s+", " ").Trim();
    }

    private static string NormalizeMaterialCode(string value)
    {
        string cleaned = value.Trim();
        if (cleaned.EndsWith(".0", StringComparison.Ordinal) &&
            double.TryParse(cleaned, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double numeric) &&
            Math.Abs(numeric - Math.Round(numeric)) < 0.000001)
            return Math.Round(numeric).ToString("0", System.Globalization.CultureInfo.InvariantCulture);
        return cleaned;
    }
}

public sealed partial class PackingListReviewRow : ObservableObject
{
    [ObservableProperty] private string _supplierCode = "";
    [ObservableProperty] private string _description = "";
    [ObservableProperty] private string _matchType = "";
    [ObservableProperty] private string _sourceLocation = "";
    [ObservableProperty] private string _targetMaterialCode = "";
    [ObservableProperty] private string _suggestedMaterialCode = "";
    [ObservableProperty] private int? _units;
    [ObservableProperty] private int? _cartons;
    [ObservableProperty] private int? _closureUnits;
    [ObservableProperty] private int? _closureCartons;
    [ObservableProperty] private int? _qcUnits;
    [ObservableProperty] private int? _qcCartons;
    [ObservableProperty] private int _confidencePercent = 100;
    public int? UnitsDifferenceFromQc => Units.HasValue && QcUnits.HasValue ? Units.Value - QcUnits.Value : null;
    public int? CartonsDifferenceFromQc => Cartons.HasValue && QcCartons.HasValue ? Cartons.Value - QcCartons.Value : null;
    [ObservableProperty] private bool _acceptUnits;
    [ObservableProperty] private bool _acceptCartons;
    [ObservableProperty] private IReadOnlyList<SupplierPackingListRow> _sourceRows = Array.Empty<SupplierPackingListRow>();

    partial void OnUnitsChanged(int? value) => OnPropertyChanged(nameof(UnitsDifferenceFromQc));
    partial void OnCartonsChanged(int? value) => OnPropertyChanged(nameof(CartonsDifferenceFromQc));
    partial void OnQcUnitsChanged(int? value) => OnPropertyChanged(nameof(UnitsDifferenceFromQc));
    partial void OnQcCartonsChanged(int? value) => OnPropertyChanged(nameof(CartonsDifferenceFromQc));
}
