namespace BasicExcel;

public class XlSheetWriter
{
    private XlWriter _xlWriter;
    private XlSheet _sheet;

    internal XlSheetWriter(XlWriter writer, XlSheet sheet)
    {
        _xlWriter = writer;
        _sheet = sheet;
    }

    public int Row { get; private set; } = 1;
    public int Col { get; private set; } = 1;

    public void StartRow(int row, XlStyle? rowStyle = null)
    {
        int styleId = _xlWriter.MapStyle(rowStyle, _sheet.Style);
        throw new NotImplementedException();
    }

    public void StartRow(XlStyle? rowStyle = null)
    {
        throw new NotImplementedException();
    }

    public void AddCell(string value, XlStyle? style = null)
    {
        throw new NotImplementedException();
    }
}
