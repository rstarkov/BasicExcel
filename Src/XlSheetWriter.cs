using System.Security;

namespace BasicExcel;

public class XlSheetWriter
{
    private XlWriter _xlWriter;
    private XlSheet _sheet;
    private StreamWriter _stream;
    private bool _rowStarted = false;

    internal XlSheetWriter(XlWriter writer, XlSheet sheet, StreamWriter stream)
    {
        _xlWriter = writer;
        _sheet = sheet;
        _stream = stream;
    }

    public int Row { get; private set; } = 1;
    public int Col { get; private set; } = 1;

    internal void Finalise()
    {
        if (_rowStarted)
            EndRow();
    }

    public void StartRow(int row, XlStyle? rowStyle = null)
    {
        // todo: skip rows
        StartRow(rowStyle);
    }

    public void StartRow(XlStyle? rowStyle = null)
    {
        if (_rowStarted)
            EndRow();
        _rowStarted = true;
        _stream.Write($"    <row");
        int styleId = _xlWriter.MapStyle(rowStyle, _sheet.Style);
        if (styleId != 0)
            _stream.Write($" s=\"{styleId}\"");
        _stream.Write(">");
    }

    private void EndRow()
    {
        if (!_rowStarted) throw new Exception();
        _stream.WriteLine("</row>");
        _rowStarted = false;
        Row++;
        Col = 1;
    }

    public void AddCell(string value, XlStyle? style = null) => addCell(value, "str", style);
    public void AddCell(int value, XlStyle? style = null) => addCell(value.ToString(), null, style);
    public void AddCell(double value, XlStyle? style = null) => addCell(value.ToString(), null, style);
    public void AddCell(decimal value, XlStyle? style = null) => addCell(value.ToString(), null, style);
    public void AddCell(bool value, XlStyle? style = null) => addCell(value ? "1" : "0", "b", style);

    private void addCell(string rawvalue, string? type, XlStyle? style)
    {
        if (!_rowStarted)
            StartRow();
        _stream.Write("<c");
        if (type != null)
            _stream.Write($" t=\"{type}\"");
        int styleId = _xlWriter.MapStyle(style, _sheet.Style);
        if (styleId != 0)
            _stream.Write($" s=\"{styleId}\"");
        _stream.Write("><v>");
        _stream.Write(SecurityElement.Escape(rawvalue));
        _stream.Write("</v></c>");
    }
}
