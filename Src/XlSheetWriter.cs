using System.Security;

namespace BasicExcel;

public class XlSheetWriter
{
    private XlWriter _xlWriter;
    private StreamWriter _stream;
    private XlSheet _sheet;
    private XlStyle _parentStyle;
    private bool _rowStarted = false;
    private XlStyle? _rowStyle; // full style with sheet+wb styling

    /// <param name="parentStyle">A merged sheet + workbook default style.</param>
    internal XlSheetWriter(XlWriter writer, StreamWriter stream, XlSheet sheet, XlStyle parentStyle)
    {
        _xlWriter = writer;
        _stream = stream;
        _sheet = sheet;
        _parentStyle = parentStyle;
        _rowStyle = _parentStyle;
    }

    public int Row { get; private set; } = 1;
    public int Col { get; private set; } = 1;

    internal void Finalise()
    {
        if (_rowStarted)
            EndRow();
    }

    public void StartRow(int row, XlStyle? rowStyle = null, double? height = null)
    {
        if (_rowStarted)
            EndRow();
        if (Row > row) throw new Exception("Can't start a row out of order");
        while (Row < row)
        {
            _stream.WriteLine($"    <row></row>");
            Row++;
        }
        StartRow(rowStyle, height);
    }

    public void StartRow(XlStyle? rowStyle = null, double? height = null)
    {
        if (_rowStarted)
            EndRow();
        _rowStarted = true;
        _rowStyle = XlStyle.New(rowStyle).Inherit(_parentStyle);
        _stream.Write($"    <row");
        if (height != null)
            _stream.Write($" ht=\"{height}\" customHeight=\"1\"");
        int styleId = _xlWriter.MapStyle(_rowStyle);
        if (styleId != 0)
            _stream.Write($" s=\"{styleId}\" customFormat=\"1\"");
        _stream.Write(">");
    }

    private void EndRow()
    {
        if (!_rowStarted) throw new Exception();
        _stream.WriteLine("</row>");
        _rowStarted = false;
        _rowStyle = _parentStyle;
        Row++;
        Col = 1;
    }

    public void AddCell(string value, XlStyle? style = null) => addCell(value, "str", style);
    public void AddCell(int value, XlStyle? style = null) => addCell(value.ToString(), null, style);
    public void AddCell(double value, XlStyle? style = null) => addCell(value.ToString(), null, style);
    public void AddCell(decimal value, XlStyle? style = null) => addCell(value.ToString(), null, style);
    public void AddCell(bool value, XlStyle? style = null) => addCell(value ? "1" : "0", "b", style);
    public void AddCell(DateTime value, XlStyle? style = null) => addCell((value - new DateTime(1899, 12, 30)).TotalDays.ToString(), null, style); // 1 day off before Feb 1900, don't care

    public void AddCell(int col, string value, XlStyle? style = null) { moveTo(col); addCell(value, "str", style); }
    public void AddCell(int col, int value, XlStyle? style = null) { moveTo(col); addCell(value.ToString(), null, style); }
    public void AddCell(int col, double value, XlStyle? style = null) { moveTo(col); addCell(value.ToString(), null, style); }
    public void AddCell(int col, decimal value, XlStyle? style = null) { moveTo(col); addCell(value.ToString(), null, style); }
    public void AddCell(int col, bool value, XlStyle? style = null) { moveTo(col); addCell(value ? "1" : "0", "b", style); }
    public void AddCell(int col, DateTime value, XlStyle? style = null) { moveTo(col); AddCell(value, style); }

    public void AddCell(int row, int col, string value, XlStyle? style = null) { moveTo(row, col); addCell(value, "str", style); }
    public void AddCell(int row, int col, int value, XlStyle? style = null) { moveTo(row, col); addCell(value.ToString(), null, style); }
    public void AddCell(int row, int col, double value, XlStyle? style = null) { moveTo(row, col); addCell(value.ToString(), null, style); }
    public void AddCell(int row, int col, decimal value, XlStyle? style = null) { moveTo(row, col); addCell(value.ToString(), null, style); }
    public void AddCell(int row, int col, bool value, XlStyle? style = null) { moveTo(row, col); addCell(value ? "1" : "0", "b", style); }
    public void AddCell(int row, int col, DateTime value, XlStyle? style = null) { moveTo(row, col); AddCell(value, style); }

    private void addCell(string rawvalue, string? type, XlStyle? style)
    {
        if (!_rowStarted)
            StartRow();
        _stream.Write("<c");
        if (type != null)
            _stream.Write($" t=\"{type}\"");
        var colStyle = _sheet.Columns.TryGetValue(Col, out var c) ? c.Style : null;
        int styleId = _xlWriter.MapStyle(XlStyle.New(style).Inherit(colStyle).Inherit(_rowStyle));
        if (styleId != 0)
            _stream.Write($" s=\"{styleId}\"");
        _stream.Write("><v>");
        _stream.Write(SecurityElement.Escape(rawvalue));
        _stream.Write("</v></c>");
        Col++;
    }

    private void moveTo(int col)
    {
        if (!_rowStarted)
            StartRow();
        if (Col > col) throw new Exception("Can't move to a column out of order");
        while (Col < col)
        {
            var colStyle = _sheet.Columns.TryGetValue(Col, out var c) ? c.Style : null;
            int styleId = _xlWriter.MapStyle(XlStyle.New(colStyle).Inherit(_rowStyle));
            if (styleId != 0)
                _stream.Write($"<c s=\"{styleId}\"></c>");
            else
                _stream.Write("<c></c>");
            Col++;
        }
    }

    private void moveTo(int row, int col)
    {
        if (!_rowStarted || Row != row)
            StartRow(row);
        moveTo(col);
    }
}
