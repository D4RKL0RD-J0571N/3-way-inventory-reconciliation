using Acr.Core.Models;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Acr.Tests.Infrastructure;

public sealed class BlindReceivingTests : IDisposable
{
    private readonly string _tempDir;

    public BlindReceivingTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"BlindReceivingTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, true);
        }
        catch { }
    }

    [Fact]
    public void ComparisonSummary_WithDeclaredUnitsAndZeroPhysicalUnits_SetsIsBlindReceivingTrue()
    {
        var summary = new ComparisonSummary
        {
            TotalDeclaredUnits = 1200,
            TotalPhysicalUnits = 0,
            NetUnitDifference = -1200,
            TotalDeclaredCartons = 50,
            TotalPhysicalCartons = 0,
            NetCartonDifference = -50,
            TotalShortageUnits = 1200,
            TotalOverageUnits = 0,
            TotalDamagedUnits = 0,
            CleanMaterialsCount = 0,
            DiscrepantMaterialsCount = 10,
            RevueltoPairsCount = 0,
            Items = Array.Empty<DiscrepancyItem>()
        };

        Assert.True(summary.IsBlindReceiving);
    }

    [Fact]
    public void ReceivingPipeline_BlindReceiving_EmitsWarningAndDetectsBlindState()
    {
        // 1. Create Synthetic Cierre Template with only invoice Pedido 41 data (500 units)
        string cierrePath = Path.Combine(_tempDir, "IMP-999_Cierre.xlsx");
        using (var doc = SpreadsheetDocument.Create(cierrePath, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet();

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CIERRE RECEP" });

            var sd = wsPart.Worksheet.AppendChild(new SheetData());

            // Header row 1: Container
            var row1 = sd.AppendChild(new Row { RowIndex = 1 });
            row1.AppendChild(CreateCell("A1", "CONTENEDOR:", CellValues.String));
            row1.AppendChild(CreateCell("B1", "IMP-999", CellValues.String));

            // Header row 12: Column headers
            var row12 = sd.AppendChild(new Row { RowIndex = 12 });
            row12.AppendChild(CreateCell("E12", "MATERIAL", CellValues.String));
            row12.AppendChild(CreateCell("G12", "CANTIDAD PEDIDO 41", CellValues.String));
            row12.AppendChild(CreateCell("AW12", "BULTOS SEGÚN PACKING LIST", CellValues.String));

            // Data row 13: 500 units declared, 0 cartons declared (or empty)
            var row13 = sd.AppendChild(new Row { RowIndex = 13 });
            row13.AppendChild(CreateCell("E13", "MAT-BLIND-01", CellValues.String));
            row13.AppendChild(CreateCell("G13", "500", CellValues.Number));
            row13.AppendChild(CreateCell("AW13", "0", CellValues.Number));

            wbPart.Workbook.Save();
        }

        // 2. Create Empty QC Workbook (0 counted physical units)
        string qcPath = Path.Combine(_tempDir, "IMP-999_QC.xlsx");
        using (var doc = SpreadsheetDocument.Create(qcPath, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var sheets = wbPart.Workbook.AppendChild(new Sheets());

            var wsPart1 = wbPart.AddNewPart<WorksheetPart>();
            wsPart1.Worksheet = new Worksheet();
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart1), SheetId = 1, Name = "CONTROL DATA" });
            var sd1 = wsPart1.Worksheet.AppendChild(new SheetData());

            // Header Row 6
            var row6 = sd1.AppendChild(new Row { RowIndex = 6 });
            row6.AppendChild(CreateCell("C6", "MATERIAL", CellValues.String));
            row6.AppendChild(CreateCell("D6", "CANTIDAD", CellValues.String));

            // Empty data or row with 0
            var row7 = sd1.AppendChild(new Row { RowIndex = 7 });
            row7.AppendChild(CreateCell("C7", "MAT-BLIND-01", CellValues.String));
            row7.AppendChild(CreateCell("D7", "0", CellValues.Number));

            wbPart.Workbook.Save();
        }

        // 3. Execute Pipeline
        var options = new ReceivingPipelineOptions
        {
            QcFilePath = qcPath,
            ClosureTemplatePath = cierrePath
        };

        ReceivingPipelineResult result = ReceivingPipeline.Execute(options);

        // Verification:
        Assert.True(result.Comparison.IsBlindReceiving, "Debe clasificarse como IsBlindReceiving");
        Assert.Equal(500, result.Comparison.TotalDeclaredUnits);
        Assert.Equal(0, result.Comparison.TotalPhysicalUnits);
        Assert.Contains(result.Warnings, w => w.Contains("[MONTAJE A CIEGAS]"));
    }

    [Fact]
    public void ReceivingPipeline_RealImp100BlindFiles_DetectsBlindReceivingCleanly()
    {
        string blindDir = @"D:\Development\Software Development\3 Way Inventory Reconciliation\data\IMP-100_MONTAJE_A_CIEGAS";
        string cierrePath = Path.Combine(blindDir, "IMP-100_CIERRE_MONTAJE_A_CIEGAS.xlsm");
        string qcPath = Path.Combine(blindDir, "IMP-100_QC_MONTAJE_A_CIEGAS.xlsm");

        if (!File.Exists(cierrePath) || !File.Exists(qcPath))
            return; // Skip if external data path is not mounted

        var options = new ReceivingPipelineOptions
        {
            QcFilePath = qcPath,
            ClosureTemplatePath = cierrePath
        };

        ReceivingPipelineResult result = ReceivingPipeline.Execute(options);

        Assert.True(result.Comparison.IsBlindReceiving);
        Assert.True(result.Comparison.TotalDeclaredUnits > 0);
        Assert.Equal(0, result.Comparison.TotalPhysicalUnits);
        Assert.Contains(result.Warnings, w => w.Contains("[MONTAJE A CIEGAS]"));
    }

    [Fact]
    public void ReceivingPipeline_ClosureOnlyWithoutQcFile_RunsInBlindReceivingModeWithoutThrowing()
    {
        string cierrePath = Path.Combine(_tempDir, "IMP-998_CierreOnly.xlsx");
        using (var doc = SpreadsheetDocument.Create(cierrePath, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet();

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CIERRE RECEP" });

            var sd = wsPart.Worksheet.AppendChild(new SheetData());
            var row12 = sd.AppendChild(new Row { RowIndex = 12 });
            row12.AppendChild(CreateCell("E12", "MATERIAL", CellValues.String));
            row12.AppendChild(CreateCell("G12", "CANTIDAD PEDIDO 41", CellValues.String));

            var row13 = sd.AppendChild(new Row { RowIndex = 13 });
            row13.AppendChild(CreateCell("E13", "MAT-NO-QC", CellValues.String));
            row13.AppendChild(CreateCell("G13", "300", CellValues.Number));

            wbPart.Workbook.Save();
        }

        ReceivingPipelineResult result = ReceivingPipeline.Execute(new ReceivingPipelineOptions
        {
            QcFilePath = "",
            ClosureTemplatePath = cierrePath
        });

        Assert.Equal("IMP-998", result.ContainerId);
        Assert.True(result.Comparison.IsBlindReceiving);
        Assert.Equal(300, result.Comparison.TotalDeclaredUnits);
        Assert.Equal(0, result.Comparison.TotalPhysicalUnits);
        Assert.Contains(result.Warnings, w => w.Contains("[MONTAJE A CIEGAS]"));
    }

    private static Cell CreateCell(string reference, string text, CellValues dataType)
    {
        return new Cell
        {
            CellReference = reference,
            DataType = dataType,
            CellValue = new CellValue(text)
        };
    }
}
