using Acr.Core.Services;
using Xunit;

namespace Acr.Tests.Core;

public sealed class SapClipboardSanitizerTests
{
    [Theory]
    [InlineData("1.455102001E+09", "1455102001")]
    [InlineData("1.455102002e+09", "1455102002")]
    [InlineData("1455102003.0", "1455102003")]
    [InlineData("1455102003.00", "1455102003")]
    [InlineData("000000001455102001", "1455102001")]
    [InlineData("1E2", "1E2")] // Alphanumeric SKU must NEVER be corrupted to 100
    [InlineData("2E5-BLK", "2E5-BLK")]
    [InlineData("  MAT-9988_01  ", "MAT-9988_01")]
    public void CleanMaterialCode_SanitizesExpectedFormats(string input, string expected)
    {
        string actual = SapClipboardSanitizer.CleanMaterialCode(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("(48)", -48.0)]
    [InlineData("$-50.00", -50.0)]
    [InlineData("-$50.00", -50.0)]
    [InlineData("₡-150.50", -150.50)]
    [InlineData("1,200", 1200.0)]
    [InlineData("1.200,50", 1200.50)]
    [InlineData("600.00", 600.0)]
    [InlineData("-", 0.0)]
    [InlineData("nan", 0.0)]
    [InlineData("#N/A", 0.0)]
    public void CleanNumber_CorrectlyParsesAccountingSignsAndCurrencies(string input, decimal expected)
    {
        decimal actual = SapClipboardSanitizer.CleanNumber(input);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void SanitizeSapTable_TabDelimitedRawDump_ProducesCleanGridAndTsv()
    {
        string raw = "Material\tDescripcion\tCantidad\tBultos\n" +
                     "000000001455102001\tCAMISETA ALGODON BASICA\t600.00\t10\n" +
                     "1.455102002E+09\tJEANS SLIM FIT INDIGO\t1,200\t20\n" +
                     "1455102003.0\tBLUSA POPELINA BLANCA\t(48)\t1\n" +
                     "1E2\tETIQUETA ALFANUMERICA\t$-50.00\t0";

        SanitizedSapTable table = SapClipboardSanitizer.SanitizeSapTable(raw);

        Assert.Equal("Tabulación", table.DelimiterLabel);
        Assert.Equal(4, table.Headers.Count);
        Assert.Equal(4, table.RowCount);

        // Row 1: stripped zeroes
        Assert.Equal("1455102001", table.Rows[0][0]);
        Assert.Equal("600", table.Rows[0][2]);

        // Row 2: scientific notation decoded
        Assert.Equal("1455102002", table.Rows[1][0]);
        Assert.Equal("1200", table.Rows[1][2]);

        // Row 3: accounting parentheses negative
        Assert.Equal("1455102003", table.Rows[2][0]);
        Assert.Equal("-48", table.Rows[2][2]);

        // Row 4: alphanumeric preserved, dollar negative preserved
        Assert.Equal("1E2", table.Rows[3][0]);
        Assert.Equal("-50", table.Rows[3][2]);

        Assert.Contains("1455102001\tCAMISETA ALGODON BASICA\t600\t10", table.CleanTsv);
    }

    [Fact]
    public void SanitizeSapTable_PipeDelimitedRawDump_HandlesPipesAndEmptyBoundaryCells()
    {
        string raw = "| Material | Descripcion | Cantidad |\n" +
                     "| 000000001455102001 | CAMISETA | 100 |\n" +
                     "| 1E2 | ETIQUETA | -25 |";

        SanitizedSapTable table = SapClipboardSanitizer.SanitizeSapTable(raw);

        Assert.Equal("Tubo (|)", table.DelimiterLabel);
        Assert.Equal(3, table.Headers.Count);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("1455102001", table.Rows[0][0]);
        Assert.Equal("1E2", table.Rows[1][0]);
    }
}
