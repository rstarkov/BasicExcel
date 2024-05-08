using System.Diagnostics;
using System.IO.Compression;
using System.Xml.Linq;

namespace BasicExcel.Playground;

internal class Program
{
    static void Main(string[] args)
    {
        //SwiftExcelLikePerfTest();
        CreateTestXmls();
        foreach (var file in new DirectoryInfo(".").EnumerateFiles("*.xlsx"))
            if (file.Name != "perftest.xlsx")
                ReformatFile(file.FullName, file.FullName + ".zip");
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
        wb3.Sheets[0].Columns[2].Style!.Mod().Color("EEEEEE", "228811").BorderL(XlBorder.Thick, "0000FF");
        wb3.Sheets[0].Columns[3].Width = 20;
        wb3.Sheets[0].WriteSheet = sw =>
        {
            sw.AddCell(2, 5, "Row 2 Col 5", XlStyle.New().Color("AA4422"));
            sw.StartRow(4, XlStyle.New().Fill("FFDD22").Font(bold: true).Align(XlHorz.Center));
            sw.AddCell("foo");
            sw.AddCell("foobar");
            sw.AddCell(6, "bar");
            sw.AddCell(6, 1, "foo", XlStyle.New().Align(XlVert.Center));
            sw.AddCell("foobar", XlStyle.New().Font(20));
            sw.AddCell("foobar абвгд", XlStyle.New().Font("Segoe UI", italic: true));
            sw.AddCell(5, "foo", XlStyle.New().Align(XlHorz.Right).BorderLR(XlBorder.Dash, "FF8080").BorderT(XlBorder.Double).BorderB(XlBorder.Thin));
        };
        wb3.Save("styles.xlsx");

        var start = DateTime.UtcNow;
        var wb4 = new XlWorkbook();
        wb4.Sheets.Add(new XlSheet());
        wb4.Sheets[0].FreezeRows = 1;
        wb4.Sheets[0].Columns[1].Width = 11;
        wb4.Sheets[0].Columns[1].Style!.Mod().Fmt(XlFmt.LocaleDate);
        wb4.Sheets[0].Columns[2].Width = 30;
        wb4.Sheets[0].Columns[2].Style!.Mod().Align(XlHorz.Center);
        wb4.Sheets[0].Columns[3].Width = 12;
        wb4.Sheets[0].Columns[3].Style!.Mod().Fmt(XlFmt.AccountingGbp);
        wb4.Sheets[0].Columns[4].Width = 15;
        wb4.Sheets[0].Columns[4].Style!.Mod().Fmt(XlFmt.LocaleDateTime);
        wb4.Sheets[0].Columns[5].Style!.Mod().Align(XlHorz.Left);
        wb4.Sheets[0].WriteSheet = sw =>
        {
            sw.StartRow(XlStyle.New().Color("FFFFFF", "008800").BorderB(XlBorder.Medium).Align(XlVert.Center), height: 32);
            sw.AddCell("Date");
            sw.AddCell("Centered");
            sw.AddCell("Total");
            sw.AddCell(5, "Left-aligned num");
            for (int i = 0; i < 20; i++)
            {
                sw.StartRow();
                sw.AddCell(DateTime.Today.AddDays(-i));
                sw.AddCell("Foobar");
                sw.AddCell(Random.Shared.Next(0, 2000_00) / 100m);
                sw.AddCell(DateTime.Now.AddDays(-i));
                sw.AddCell(Random.Shared.Next(0, 999));
            }
        };
        wb4.Sheets.Add(new XlSheet { Name = "Freeze 1 col", FreezeCols = 1 });
        wb4.Sheets.Add(new XlSheet { Name = "Freeze 2 rows", FreezeRows = 2 });
        wb4.Sheets.Add(new XlSheet { Name = "Freeze 2 cols", FreezeCols = 2 });
        wb4.Sheets.Add(new XlSheet { Name = "Freeze 1r1c", FreezeRows = 1, FreezeCols = 1 });
        wb4.Sheets.Add(new XlSheet { Name = "Freeze 2r2c", FreezeRows = 2, FreezeCols = 2 });
        wb4.Save("formats.xlsx");
        Console.WriteLine($"{(DateTime.UtcNow - start).TotalMilliseconds:0}ms"); // 30k cells = 165ms

        var wb5 = new XlWorkbook();
        wb5.Style.Mod().Color("555500").Font("Courier New", 15, italic: true).Align(XlHorz.Center).Fmt(XlFmt.AccountingGbp);
        wb5.Sheets.Add(new XlSheet { Name = "Book style" });
        wb5.Sheets[0].WriteSheet = sw =>
        {
            for (int i = 0; i < 10; i++)
            {
                sw.StartRow();
                sw.AddCell("foo");
                sw.AddCell(0.4725);
                sw.AddCell(7, "baz");
            }
            sw.AddCell(20, 7, "end");
        };
        wb5.Save("defaultbook.xlsx");

        var wb6 = new XlWorkbook();
        wb6.Sheets.Add(new XlSheet { Name = "Sheet style" });
        wb6.Sheets[0].Style!.Mod().Color("00FF00", "0000FF").Font("Arial", 14, bold: true).Align(XlHorz.Right, XlVert.Top).Border(XlBorder.Double, "00FF00").Fmt(XlFmt.PercentFrac);
        wb6.Sheets[0].WriteSheet = sw =>
        {
            for (int i = 0; i < 10; i++)
            {
                sw.StartRow();
                sw.AddCell("foo");
                sw.AddCell(0.4725);
                sw.AddCell(7, "baz");
            }
            sw.AddCell(20, 7, "end");
        };
        wb6.Save("defaultsheet.xlsx");
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

    static void SwiftExcelLikePerfTest()
    {
        var p = Process.GetCurrentProcess();
        var baseline = p.PeakPagedMemorySize64;
        var start = DateTime.UtcNow;
        var wb = new XlWorkbook();
        wb.Sheets.Add(new XlSheet());
        wb.Sheets[0].WriteSheet = sw =>
        {
            for (var row = 1; row <= 100000; row++)
                for (var col = 1; col <= 100; col++)
                    sw.AddCell(row, col, $"row:{row}-col:{col}");
        };
        wb.Save("perftest.xlsx");
        Console.WriteLine($"{(DateTime.UtcNow - start).TotalMilliseconds:0}ms");
        p.Refresh();
        Console.WriteLine($"peak paged={p.PeakPagedMemorySize64 - baseline:#,0} + baseline={baseline:#,0}");
    }
}
