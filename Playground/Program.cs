using System.IO.Compression;
using System.Xml.Linq;

namespace BasicExcel.Playground;

internal class Program
{
    static void Main(string[] args)
    {
        CreateTestXmls();
        ReformatFile(@"empty.xlsx", @"empty.xlsx.zip");
        ReformatFile(@"basic.xlsx", @"basic.xlsx.zip");
    }

    static void CreateTestXmls()
    {
        var wb1 = new XlWorkbook();
        wb1.Save("empty.xlsx");

        var wb2 = new XlWorkbook();
        wb2.Sheets.Add(new XlSheet());
        wb2.Sheets[0].Name = "Test sheet";
        wb2.Sheets[0].WriteSheet = sw =>
        {
            sw.AddCell("Foo");
            sw.AddCell("Bar");
            sw.AddCell(123);
            sw.AddCell(Math.PI);
            sw.AddCell(true);
            sw.AddCell(false);
            sw.StartRow(3);
            sw.AddCell("Row 3");
            sw.AddCell("Row 3");
            sw.AddCell(5, "Row 3 Col 5");
            sw.AddCell(7, 4, "Row 7 Col 4");
        };
        wb2.Save("basic.xlsx");

        var wb3 = new XlWorkbook();
        wb3.Sheets.Add(new XlSheet());
        wb3.Sheets[0].Columns[2].Style.Mod().Color("EEEEEE", "228811").BorderL(XlBorder.Thick, "0000FF");
        wb3.Sheets[0].Columns[3].Width = 20;
        wb3.Sheets[0].WriteSheet = sw =>
        {
            sw.AddCell(2, 5, "Row 2 Col 5", new XlStyle().Mod().Color("AA4422"));
            sw.StartRow(4, new XlStyle().Mod().Fill("FFDD22").Font(bold: true).Align(XlHorz.Center));
            sw.AddCell("foo");
            sw.AddCell("foobar");
            sw.AddCell(6, 1, "foo", new XlStyle().Mod().Align(XlVert.Center));
            sw.AddCell("foobar", new XlStyle().Mod().Font(20));
            sw.AddCell("foobar абвгд", new XlStyle().Mod().Font("Segoe UI", italic: true));
            sw.AddCell(5, "foo", new XlStyle().Mod().Align(XlHorz.Right).BorderLR(XlBorder.Dash, "FF8080").BorderT(XlBorder.Double).BorderB(XlBorder.Thin));
        };
        wb3.Save("styles.xlsx");
    }

    static void ReformatFile(string inputPath, string outputPath)
    {
        using var inputZip = ZipFile.OpenRead(inputPath);
        using var outputZip = new ZipArchive(File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.Read), ZipArchiveMode.Create);

        foreach (var oe in inputZip.Entries)
        {
            var ne = outputZip.CreateEntry(oe.FullName, CompressionLevel.SmallestSize);
            using var ns = ne.Open();
            using var os = oe.Open();
            var xml = XDocument.Load(os);
            xml.Save(ns, SaveOptions.None);
        }
    }
}
