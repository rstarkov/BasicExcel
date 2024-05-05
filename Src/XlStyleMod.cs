namespace BasicExcel;

public class XlStyleMod(XlStyle style)
{
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
        style.BrLeft = type;
        style.BrRight = type;
        style.BrTop = type;
        style.BrBot = type;
        if (color != null)
            style.BrLeftColor = style.BrRightColor = style.BrTopColor = style.BrBotColor = color;
        return this;
    }
}
