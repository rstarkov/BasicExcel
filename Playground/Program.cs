using System.IO.Compression;
using System.Xml.Linq;

namespace BasicExcel.Playground;

internal class Program
{
    static void Main(string[] args)
    {
        CreateTestXmls();
        ReformatFile(@"empty.xlsx", @"empty.xlsx.zip");
    }

    static void CreateTestXmls()
    {
        var wb1 = new XlWorkbook();
        wb1.Save("empty.xlsx");
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
