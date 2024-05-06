namespace BasicExcel;

public class XlStyleMod
{
    private XlStyle style;
    public XlStyleMod(XlStyle style) => this.style = style;

    public static implicit operator XlStyle(XlStyleMod mod) => mod.style;

    public XlStyleMod Fmt(string fmt)
    {
        style.Format = fmt;
        return this;
    }

    public XlStyleMod Font(string font, double? size = null, bool? bold = null, bool? italic = null)
    {
        style.Font = font;
        if (size != null) style.Size = size;
        if (bold != null) style.Bold = bold;
        if (italic != null) style.Italic = italic;
        return this;
    }

    public XlStyleMod Font(double size, bool? bold = null, bool? italic = null)
    {
        style.Size = size;
        if (bold != null) style.Bold = bold;
        if (italic != null) style.Italic = italic;
        return this;
    }

    public XlStyleMod Font(bool bold, bool? italic = null)
    {
        style.Bold = bold;
        if (italic != null) style.Italic = italic;
        return this;
    }

    public XlStyleMod Color(string textColor, string? fillColor = null)
    {
        style.Color = textColor;
        if (fillColor != null) style.FillColor = fillColor;
        return this;
    }

    public XlStyleMod Fill(string fillColor)
    {
        style.FillColor = fillColor;
        return this;
    }

    public XlStyleMod Align(XlHorz horz, XlVert? vert = null, bool? wrap = null)
    {
        style.Horz = horz;
        if (vert != null) style.Vert = vert;
        if (wrap != null) style.Wrap = wrap;
        return this;
    }

    public XlStyleMod Align(XlVert vert, bool? wrap = null)
    {
        style.Vert = vert;
        if (wrap != null) style.Wrap = wrap;
        return this;
    }

    public XlStyleMod Border(XlBorder type, string? color = null)
    {
        style.BrLeft = style.BrRight = style.BrTop = style.BrBot = type;
        if (color != null) style.BrLeftColor = style.BrRightColor = style.BrTopColor = style.BrBotColor = color;
        return this;
    }

    public XlStyleMod BorderL(XlBorder type, string? color = null)
    {
        style.BrLeft = type;
        if (color != null) style.BrLeftColor = color;
        return this;
    }

    public XlStyleMod BorderR(XlBorder type, string? color = null)
    {
        style.BrRight = type;
        if (color != null) style.BrRightColor = color;
        return this;
    }

    public XlStyleMod BorderT(XlBorder type, string? color = null)
    {
        style.BrTop = type;
        if (color != null) style.BrTopColor = color;
        return this;
    }

    public XlStyleMod BorderB(XlBorder type, string? color = null)
    {
        style.BrBot = type;
        if (color != null) style.BrBotColor = color;
        return this;
    }

    public XlStyleMod BorderLR(XlBorder type, string? color = null)
    {
        style.BrLeft = style.BrRight = type;
        if (color != null) style.BrLeftColor = style.BrRightColor = color;
        return this;
    }

    public XlStyleMod BorderTB(XlBorder type, string? color = null)
    {
        style.BrTop = style.BrBot = type;
        if (color != null) style.BrTopColor = style.BrBotColor = color;
        return this;
    }

    public XlStyleMod Inherit(XlStyle? s)
    {
        style.Inherit(s);
        return this;
    }
}
