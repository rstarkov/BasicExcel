using System.Security;

namespace BasicExcel;

public class XlSheetWriter
{
    private XlWriter _xlWriter;
    private StreamWriter _stream;
    private XlSheet _sheet;
    private XlStyle _parentStyle;
    private bool _rowStarted = false; // only ever false when the sheet is blank
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
            _stream.WriteLine("</row>");
    }

    public void StartRow(int row, XlStyle? style = null, double? height = null)
    {
        if (_rowStarted)
        {
            _stream.WriteLine("</row>");
            Row++;
        }
        if (row < Row) throw new Exception("Can't start a row out of order");
        if (row == Row)
            _stream.Write("    <row");
        else
            _stream.Write($"    <row r=\"{row}\"");
        _rowStarted = true;
        Row = row;
        Col = 1;
        if (height != null)
            _stream.Write($" ht=\"{height}\" customHeight=\"1\"");
        _rowStyle = XlStyle.New(style).Inherit(_parentStyle)!;
        int styleId = _xlWriter.MapStyle(_rowStyle);
        if (styleId != 0)
            _stream.Write($" s=\"{styleId}\" customFormat=\"1\"");
        _stream.Write(">");
    }

    public void StartRow(XlStyle? style = null, double? height = null)
    {
        StartRow(_rowStarted ? Row + 1 : Row, style, height);
    }

    public void AddCell(string value, XlStyle? style = null) => AddCell(Row, Col, value, style);
    public void AddCell(int value, XlStyle? style = null) => AddCell(Row, Col, value, style);
    public void AddCell(double value, XlStyle? style = null) => AddCell(Row, Col, value, style);
    public void AddCell(decimal value, XlStyle? style = null) => AddCell(Row, Col, value, style);
    public void AddCell(bool value, XlStyle? style = null) => AddCell(Row, Col, value, style);
    public void AddCell(DateTime value, XlStyle? style = null) => AddCell(Row, Col, value, style);

    public void AddCell(int col, string value, XlStyle? style = null) => AddCell(Row, col, value, style);
    public void AddCell(int col, int value, XlStyle? style = null) => AddCell(Row, col, value, style);
    public void AddCell(int col, double value, XlStyle? style = null) => AddCell(Row, col, value, style);
    public void AddCell(int col, decimal value, XlStyle? style = null) => AddCell(Row, col, value, style);
    public void AddCell(int col, bool value, XlStyle? style = null) => AddCell(Row, col, value, style);
    public void AddCell(int col, DateTime value, XlStyle? style = null) => AddCell(Row, col, value, style);

    public void AddCell(int row, int col, string value, XlStyle? style = null) { addCell(row, col, value, "str", style); }
    public void AddCell(int row, int col, int value, XlStyle? style = null) { addCell(row, col, value.ToString(), null, style); }
    public void AddCell(int row, int col, double value, XlStyle? style = null) { addCell(row, col, value.ToString(), null, style); }
    public void AddCell(int row, int col, decimal value, XlStyle? style = null) { addCell(row, col, value.ToString(), null, style); }
    public void AddCell(int row, int col, bool value, XlStyle? style = null) { addCell(row, col, value ? "1" : "0", "b", style); }
    public void AddCell(int row, int col, DateTime value, XlStyle? style = null) { addCell(row, col, (value - new DateTime(1899, 12, 30)).TotalDays.ToString(), null, style); } // 1 day off before Feb 1900, don't care

    private void addCell(int row, int col, string rawvalue, string? type, XlStyle? style)
    {
        if (!_rowStarted || row != Row)
            StartRow(row);
        if (col < Col) throw new Exception("Can't write a cell out of order");
        if (col == Col)
            _stream.Write("<c");
        else
            _stream.Write($"<c r=\"{XlUtil.CellRef(row, col)}\"");
        if (type != null)
            _stream.Write($" t=\"{type}\"");
        var colStyle = _sheet.Columns.TryGetValue(col, out var c) ? c.Style : null;
        int styleId = _xlWriter.MapStyle(XlStyle.New(style).Inherit(colStyle).Inherit(_rowStyle));
        if (styleId != 0)
            _stream.Write($" s=\"{styleId}\"");
        _stream.Write("><v>");
        _stream.Write(SecurityElement.Escape(rawvalue));
        _stream.Write("</v></c>");
        Col = col + 1;
    }
}
