namespace Acr.Core.Models;

/// <summary>
/// Modelo de presentación y proyección oficial para la hoja 'Bultos' y análisis de empaque.
/// Basado en la especificación de 10 columnas enlazadas 1 a 1 con 'CIERRE RECEP'
/// y desglose de empaque de muelle (CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md).
/// </summary>
public sealed record BultosPresentationItem
{
    /// <summary>Fila correspondiente en CIERRE RECEP (ej. 13+).</summary>
    public int CierreRow { get; init; }

    /// <summary>Fila correspondiente en la hoja Bultos (CierreRow - 10).</summary>
    public int BultosRow => CierreRow >= 13 ? CierreRow - 10 : 0;

    /// <summary>Col B (2): Categoría del producto (ej. DAMAS, CABALLEROS, ESCOLAR).</summary>
    public string Category { get; init; } = "";

    /// <summary>Col C (3): Bultos físicos descargados en muelle (Z-Carga, Col AS).</summary>
    public int PhysicalCartons { get; init; }

    /// <summary>Col D (4): Unidades físicas descargadas en muelle (Z-Carga, Col AT).</summary>
    public int PhysicalUnits { get; init; }

    /// <summary>Col E (5): Código de material SAP (Padre / Estilo).</summary>
    public string MaterialCode { get; init; } = "";

    /// <summary>Col F (6): Unidades finales recibidas (Col H).</summary>
    public int ReceivedUnits { get; init; }

    /// <summary>Col G (7): Ratio de unidades por prepack (Curva, Col BB).</summary>
    public int Curve { get; init; }

    /// <summary>Col H (8): Unidad de medida (PRK o UN, Col BC).</summary>
    public string Uom { get; init; } = "UN";

    /// <summary>Col I (9): Descripción completa del material (Col AD).</summary>
    public string Description { get; init; } = "";

    /// <summary>Col J (10): Expresión o unidades de empaque (Col AY, ej. "60", "39*60 1*78").</summary>
    public string PackagingExpression { get; init; } = "";

    /// <summary>Bultos Enteros (cajas completas estándar).</summary>
    public int FullCartons { get; init; }

    /// <summary>Unidades por bulto entero.</summary>
    public int UnitsPerFullCarton { get; init; }

    /// <summary>Bultos Complemento o Picos (cajas incompletas con saldo).</summary>
    public int ComplementCartons { get; init; }

    /// <summary>Unidades en bultos complemento / picos.</summary>
    public int UnitsInComplement { get; init; }

    /// <summary>Bultos teóricos de Packing List (Col AW).</summary>
    public int DeclaredCartons { get; init; }

    /// <summary>Diferencia de Bultos (Físicos Muelle - Packing List).</summary>
    public int DifferenceCartons => PhysicalCartons - DeclaredCartons;

    /// <summary>Color para resaltar diferencias de bultos (Verde si 0, Rojo si negativo, Ámbar si positivo).</summary>
    public string DifferenceCartonsColor => DifferenceCartons switch
    {
        0 => "#16A34A",
        < 0 => "#DC2626",
        > 0 => "#D97706"
    };

    /// <summary>Color de fondo para la fila de bultos si hay descuadre.</summary>
    public string RowBackgroundColor => DifferenceCartons == 0 ? "#FFFFFF" : "#FFFBEB";

    /// <summary>Observaciones del muelle o notas de empaque.</summary>
    public string Observation { get; init; } = "";
}
