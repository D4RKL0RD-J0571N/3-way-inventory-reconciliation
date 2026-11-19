using System.Text.RegularExpressions;
using Acr.Core.Models;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Acr.Tests.Infrastructure;

public sealed class Imp207HardeningTests : IDisposable
{
    private readonly string _tempDir;

    public Imp207HardeningTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"Imp207Tests_{Guid.NewGuid():N}");
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
    public void ClosureTemplateReader_Imp207WithRow22FaltanteImportacion288_DoesNotOverwriteContainerId()
    {
        // Create synthetic closure workbook named IMP-207_Cierre.xlsx
        string filePath = Path.Combine(_tempDir, "IMP-207_Cierre.xlsx");
        using (var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet();

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CIERRE RECEP" });

            var sd = wsPart.Worksheet.AppendChild(new SheetData());

            // Header row 2: "Faltante Importación" = "288"
            var row2 = sd.AppendChild(new Row { RowIndex = 2 });
            row2.AppendChild(CreateCell("A2", "Faltante Importación", CellValues.String));
            row2.AppendChild(CreateCell("B2", "288", CellValues.String));

            // Header row 12: Column headers
            var row12 = sd.AppendChild(new Row { RowIndex = 12 });
            row12.AppendChild(CreateCell("E12", "MATERIAL", CellValues.String));
            row12.AppendChild(CreateCell("G12", "CANTIDAD PEDIDO 41", CellValues.String));
            row12.AppendChild(CreateCell("AW12", "BULTOS SEGÚN PACKING LIST", CellValues.String));

            // Data row 13
            var row13 = sd.AppendChild(new Row { RowIndex = 13 });
            row13.AppendChild(CreateCell("E13", "MAT-100", CellValues.String));
            row13.AppendChild(CreateCell("G13", "50", CellValues.Number));
            row13.AppendChild(CreateCell("AW13", "5", CellValues.Number));

            wbPart.Workbook.Save();
        }

        ClosureTemplateData data = ClosureTemplateReader.ReadDeclaredItems(filePath);

        // ContainerId MUST be IMP-207 from filename fallback and NEVER "288"!
        Assert.Equal("IMP-207", data.ContainerId);
        Assert.Single(data.DeclaredItems);
        Assert.Equal("MAT-100", data.DeclaredItems[0].MaterialCode);
    }

    [Fact]
    public void QcReader_Layout4AndEsquemaAlisto_EmitsLayout4WarningAndExcludesEsquemaAlistoFromParts()
    {
        string filePath = Path.Combine(_tempDir, "IMP-207_QC.xlsm");
        using (var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var sheets = wbPart.Workbook.AppendChild(new Sheets());

            // Sheet 1: CONTROL DATA
            var wsPart1 = wbPart.AddNewPart<WorksheetPart>();
            wsPart1.Worksheet = new Worksheet();
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart1), SheetId = 1, Name = "CONTROL DATA" });
            var sd1 = wsPart1.Worksheet.AppendChild(new SheetData());

            // Header Row 6: Notice Bultos Complemento is omitted or labeled "UNIDADES BULTOS COMPLEMENTOS" (Layout #4)
            var row6 = sd1.AppendChild(new Row { RowIndex = 6 });
            row6.AppendChild(CreateCell("C6", "MATERIAL", CellValues.String));
            row6.AppendChild(CreateCell("D6", "CANTIDAD", CellValues.String));
            row6.AppendChild(CreateCell("G6", "BULTOS ENTEROS", CellValues.String));
            row6.AppendChild(CreateCell("H6", "UNIDADES POR BULTO", CellValues.String));
            // Unidades bultos complementos without Bultos complementos column
            row6.AppendChild(CreateCell("I6", "UNIDADES BULTOS COMPLEMENTOS", CellValues.String));
            row6.AppendChild(CreateCell("J6", "TOTAL DE BULTOS", CellValues.String));
            row6.AppendChild(CreateCell("K6", "TOTAL DE UNIDADES", CellValues.String));

            // Data row 7
            var row7 = sd1.AppendChild(new Row { RowIndex = 7 });
            row7.AppendChild(CreateCell("C7", "MAT-101", CellValues.String));
            row7.AppendChild(CreateCell("D7", "420", CellValues.Number));
            row7.AppendChild(CreateCell("G7", "8", CellValues.Number));
            row7.AppendChild(CreateCell("H7", "48", CellValues.Number));
            row7.AppendChild(CreateCell("I7", "36", CellValues.Number));
            row7.AppendChild(CreateCell("J7", "8", CellValues.Number));
            row7.AppendChild(CreateCell("K7", "420", CellValues.Number));

            // Sheet 2: "ESQUEMA ALISTO PARTE B" (MUST BE IGNORED)
            var wsPart2 = wbPart.AddNewPart<WorksheetPart>();
            wsPart2.Worksheet = new Worksheet();
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart2), SheetId = 2, Name = "ESQUEMA ALISTO PARTE B" });

            // Sheet 3: "ESQUEMA ALISTO PARTE C" (MUST BE IGNORED)
            var wsPart3 = wbPart.AddNewPart<WorksheetPart>();
            wsPart3.Worksheet = new Worksheet();
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart3), SheetId = 3, Name = "ESQUEMA ALISTO PARTE C" });

            wbPart.Workbook.Save();
        }

        QcInspectionReport qcReport = QcReader.Read(filePath);

        // 1. Layout #4 mathematical rescue warning emitted
        Assert.Contains(qcReport.Warnings, w => w.Contains("Layout #4"));

        // 2. ESQUEMA ALISTO sheets are strictly ignored from detected parts
        Assert.Empty(qcReport.DetectedParts);

        // 3. Mathematical rescue: 36 units in complement rescued
        var item = qcReport.Items.Single(i => i.MaterialCode == "MAT-101");
        Assert.Equal(420, item.CalculationResult!.CalculatedUnits);
        Assert.Contains(QuantityEngineFlags.FlagQcComplementWithoutCartonCount, item.CalculationResult.Flags);
    }

    private static Cell CreateCell(string cellRef, string value, CellValues dataType) =>
        new() { CellReference = cellRef, DataType = dataType, CellValue = new CellValue(value) };
}
