using BasicExcel.Util;

namespace BasicExcel;

public class XlWorkbook
{
    public string Creator { get; set; } = "";
    public string LastModifiedBy { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime ModifiedAt { get; set; } = DateTime.Now;
    public List<XlSheet> Sheets { get; set; } = [];
    public XlSheet? ActiveSheet { get; set; } = null;
    /// <summary>
    ///     Default style for all cells in the workbook, unless overridden by <see cref="XlSheet.Style"/>, <see
    ///     cref="XlColumn.Style"/> or a cell style.</summary>
    public XlStyle Style { get; set; } = new()
    {
        Format = XlFmt.General,
        Font = "Aptos Narrow",
        Size = 11,
        Color = "",
        Bold = false,
        Italic = false,
        FillColor = "",
        Horz = XlHorz.Auto,
        Vert = XlVert.Bottom,
        Wrap = false,
        BrLeft = XlBorder.None,
        BrLeftColor = "",
        BrRight = XlBorder.None,
        BrRightColor = "",
        BrTop = XlBorder.None,
        BrTopColor = "",
        BrBot = XlBorder.None,
        BrBotColor = "",
    };

    public void Save(Stream stream, bool leaveOpen = false)
    {
        using var writer = new XlWriter(this, stream, leaveOpen);
        writer.Write();
    }

    public void Save(string filePath)
    {
        using var stream = File.Open(filePath, FileMode.Create, FileAccess.Write, FileShare.Read);
        using var writer = new XlWriter(this, stream);
        writer.Write();
    }
}

public class XlSheet
{
    /// <summary>Name of the sheet. Automatically deduplicated if necessary. Defaults to "Sheet".</summary>
    public string Name { get; set; } = "Sheet";
    public AutoDictionary<int, XlColumn> Columns { get; } = new(_ => new());
    public int? FreezeRows { get; set; } = null;
    public int? FreezeCols { get; set; } = null;
    public XlStyle? Style { get; set; } = new();

    public Action<XlSheetWriter> WriteSheet { get; set; } = (XlSheetWriter w) => { };
}

public class XlColumn
{
    public double? Width { get; set; } = null;
    public XlStyle? Style { get; set; } = new();
}

public record class XlStyle
{
    public string? Format { get; set; }

    public string? Font { get; set; }
    public double? Size { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    /// <summary>Text color as RGB or ARGB hex, or empty string for theme default, or null to inherit.</summary>
    public string? Color { get; set; }
    /// <summary>Solid fill color as RGB or ARGB hex, or empty string for fill pattern "none", or null to inherit.</summary>
    public string? FillColor { get; set; }

    public XlHorz? Horz { get; set; }
    public XlVert? Vert { get; set; }
    public bool? Wrap { get; set; }

    public XlBorder? BrLeft { get; set; }
    /// <summary>Border color as RGB or ARGB hex, or empty string for "auto", or null to inherit.</summary>
    public string? BrLeftColor { get; set; }
    public XlBorder? BrRight { get; set; }
    /// <summary>Border color as RGB or ARGB hex, or empty string for "auto", or null to inherit.</summary>
    public string? BrRightColor { get; set; }
    public XlBorder? BrTop { get; set; }
    /// <summary>Border color as RGB or ARGB hex, or empty string for "auto", or null to inherit.</summary>
    public string? BrTopColor { get; set; }
    public XlBorder? BrBot { get; set; }
    /// <summary>Border color as RGB or ARGB hex, or empty string for "auto", or null to inherit.</summary>
    public string? BrBotColor { get; set; }

    /// <summary>Inherit non-null properties from the parent style, modifying this style in-place.</summary>
    public XlStyle Inherit(XlStyle? parent)
    {
        if (parent == null) return this;
        Format ??= parent.Format;
        Font ??= parent.Font;
        Size ??= parent.Size;
        Bold ??= parent.Bold;
        Italic ??= parent.Italic;
        Color ??= parent.Color;
        FillColor ??= parent.FillColor;
        Horz ??= parent.Horz;
        Vert ??= parent.Vert;
        Wrap ??= parent.Wrap;
        BrLeft ??= parent.BrLeft;
        BrLeftColor ??= parent.BrLeftColor;
        BrRight ??= parent.BrRight;
        BrRightColor ??= parent.BrRightColor;
        BrTop ??= parent.BrTop;
        BrTopColor ??= parent.BrTopColor;
        BrBot ??= parent.BrBot;
        BrBotColor ??= parent.BrBotColor;
        return this;
    }

    /// <summary>A helper for modifying this style in place using a fluent style.</summary>
    public XlStyleMod Mod() => new XlStyleMod(this);
    /// <summary>A helper for creating a new style from scratch using fluent style. Same as <c>new XlStyle().Mod()</c>.</summary>
    public static XlStyleMod New() => new XlStyleMod(new XlStyle());
    /// <summary>A helper for creating a new style by cloning another.</summary>
    public static XlStyleMod New(XlStyle? inherit) => new XlStyleMod(new XlStyle().Inherit(inherit));
}

public enum XlHorz { Auto = 0, Left, Center, Right } // do not rename - .ToString written to output files
public enum XlVert { Bottom = 0, Center, Top } // do not rename - .ToString written to output files
public enum XlBorder { None = 0, Hair, Thin, Medium, Thick, Dot, Dash, MediumDash, DashDot, MediumDashDot, DashDotDot, MediumDashDotDot, SlantDashDot, Double } // do not reorder - see XlWriter lookup array

public static class XlFmt
{
    // built-in number formats
    public const string General = "General";
    public const string NumberWhole = "0";
    public const string NumberFrac = "0.00";
    public const string NumberWholeThouSep = "#,##0";
    public const string NumberFracThouSep = "#,##0.00";
    public const string PercentWhole = "0%";
    public const string PercentFrac = "0.00%";
    public const string Scientific = "0.00E+00";
    /// <summary>Locale-specific date-only format, eg "31/01/2024" in UK.</summary>
    public const string LocaleDate = "<LD>"; // don't use a specific string here like d/m/yyyy because that prevents using an actual literal "d/m/yyyy" as a format, remapping it to a locale-specific format
    /// <summary>Locale-specific date-time format, eg "31/01/2024 21:59" in UK.</summary>
    public const string LocaleDateTime = "<LDT>";
    public const string Text = "@";

    // some non-builtin helpers
    public const string AccountingGbp = """_-[$£-809]* #,##0.00_-;\-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-""";

    public static int StandardNumberFormatId(string numberFormat)
    {
        return numberFormat switch
        {
            General => 0,
            NumberWhole => 1,
            NumberFrac => 2,
            NumberWholeThouSep => 3,
            NumberFracThouSep => 4,
            PercentWhole => 9,
            PercentFrac => 10,
            Scientific => 11,
            LocaleDate => 14,
            LocaleDateTime => 22,
            Text => 49,
            _ => -1,
        };
    }
}
