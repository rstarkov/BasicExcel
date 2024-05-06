using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

namespace BasicExcel;

internal class XlWriter : IDisposable
{
    private XlWorkbook _wb;
    private ZipArchive _zip;

    public XlWriter(XlWorkbook workbook, Stream stream, bool leaveOpen = false)
    {
        _wb = workbook;
        _zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen);
    }

    public void Dispose()
    {
        _zip?.Dispose();
        _zip = null!;
    }

    private void writeFile(string path, string content)
    {
        var entry = _zip.CreateEntry(path, CompressionLevel.SmallestSize);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    public void Write()
    {
        initialiseStyles();
        writeSheets(); // also collects styles that need to be saved
        writeStyles();
        writeMisc();
        writeWorkbook();
    }

    private void writeSheets()
    {
        if (_wb.Sheets.Count == 0)
            _wb.Sheets.Add(new XlSheet());
        if (_wb.Sheets.Count != _wb.Sheets.Distinct().Count())
            throw new InvalidOperationException("Multiple instances of the same sheet are not supported.");
        foreach (var dupes in _wb.Sheets.GroupBy(s => s.Name).Where(g => g.Count() > 1))
        {
            var i = 1;
            foreach (var sheet in dupes)
            {
                while (_wb.Sheets.Any(s => s.Name == sheet.Name + i))
                    i++;
                sheet.Name = sheet.Name + i;
            }
        }

        for (int si = 0; si < _wb.Sheets.Count; si++)
            writeSheet(_wb.Sheets[si], si);
    }

    private void writeSheet(XlSheet s, int si)
    {
        var entry = _zip.CreateEntry($"xl/worksheets/sheet{si + 1}.xml", CompressionLevel.SmallestSize);
        using var writer = new StreamWriter(entry.Open());
        writer.WriteLine(
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">
              <dimension ref="A1"/>
            """);
        var viewXml = new XElement("sheetView", new XAttribute("workbookViewId", "0"));
        if (_wb.ActiveSheet == s)
            viewXml.Add(new XAttribute("tabSelected", "1"));
        if (s.FreezeRows != null || s.FreezeCols != null)
        {
            var topleftCell = XlUtil.CellRef((s.FreezeRows ?? 0) + 1, (s.FreezeCols ?? 0) + 1);
            var activePane = s.FreezeRows == null ? "topRight" : s.FreezeCols == null ? "bottomLeft" : "bottomRight";
            XElement paneXml = new XElement("pane", new XAttribute("state", "frozen"), new XAttribute("topLeftCell", topleftCell), new XAttribute("activePane", activePane));
            if (s.FreezeCols != null)
                paneXml.Add(new XAttribute("xSplit", s.FreezeCols.Value));
            if (s.FreezeRows != null)
                paneXml.Add(new XAttribute("ySplit", s.FreezeRows.Value));
            viewXml.Add(paneXml);
            viewXml.Add(new XElement("selection", new XAttribute("pane", activePane), new XAttribute("activeCell", topleftCell), new XAttribute("sqref", topleftCell)));
        }
        writer.Write("  ");
        writer.Write(new XElement("sheetViews", viewXml).ToString(SaveOptions.DisableFormatting));
        writer.WriteLine(
            """

              <sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35" />
            """);

        var fullSheetStyle = XlStyle.New(s.Style).Inherit(_wb.Style);
        if (s.Columns.Count > 0)
        {
            writer.WriteLine("  <cols>");
            foreach (var kvp in s.Columns.OrderBy(kvp => kvp.Key))
            {
                writer.Write($"    <col min=\"{kvp.Key}\" max=\"{kvp.Key}\" width=\"{kvp.Value.Width ?? 8.7265625:0.###}\"");
                if (kvp.Value.Width != null) // width is mandatory; without it the style has no effect. "bestFit" doesn't auto-size on load so not supported here.
                    writer.Write(" customWidth=\"1\""); // "customWidth" doesn't seem to do anything but write it out to match what Excel does just in case
                var styleId = MapStyle(XlStyle.New(kvp.Value.Style).Inherit(fullSheetStyle));
                if (styleId != 0)
                    writer.Write($" style=\"{styleId}\"");
                writer.WriteLine(" />");
            }
            writer.WriteLine("  </cols>");
        }

        writer.WriteLine("  <sheetData>");
        var sw = new XlSheetWriter(this, writer, s, fullSheetStyle);
        s.WriteSheet(sw);
        sw.Finalise();

        writer.Write(
            """
              </sheetData>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
            </worksheet>
            """);
    }

    private void writeStyles()
    {
        var entry = _zip.CreateEntry($"xl/styles.xml", CompressionLevel.SmallestSize);
        using var writer = new StreamWriter(entry.Open());
        writer.WriteLine(
            """
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
            """);

        if (_sxNumFmts.Count > 0)
        {
            writer.WriteLine($"""  <numFmts count="{_sxNumFmts.Count}">""");
            foreach (var kvp in _sxNumFmts.OrderBy(kvp => kvp.Value))
            {
                writer.Write($"    <numFmt numFmtId=\"{kvp.Value}\" formatCode=\"");
                writer.Write(SecurityElement.Escape(kvp.Key));
                writer.WriteLine("\"/>");
            }
            writer.WriteLine($"""  </numFmts>""");
        }

        writer.WriteLine($"""  <fonts count="{_sxFontsXml.Count}" x14ac:knownFonts="1">""");
        foreach (var kvp in _sxFontsXml.OrderBy(kvp => kvp.Value))
        {
            writer.Write("    ");
            writer.WriteLine(kvp.Key);
        }
        writer.WriteLine($"""  </fonts>""");

        writer.WriteLine($"""  <fills count="{_sxFillsXml.Count}">""");
        foreach (var kvp in _sxFillsXml.OrderBy(kvp => kvp.Value))
        {
            writer.Write("    ");
            writer.WriteLine(kvp.Key);
        }
        writer.WriteLine($"""  </fills>""");

        writer.WriteLine($"""  <borders count="{_sxBordersXml.Count}">""");
        foreach (var kvp in _sxBordersXml.OrderBy(kvp => kvp.Value))
        {
            writer.Write("    ");
            writer.WriteLine(kvp.Key);
        }
        writer.WriteLine($"""  </borders>""");

        writer.WriteLine(
            """
              <cellStyleXfs count="1">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
              </cellStyleXfs>
            """);

        writer.WriteLine($"""  <cellXfs count="{_sxXfsXml.Count}">""");
        foreach (var kvp in _sxXfsXml.OrderBy(kvp => kvp.Value))
        {
            writer.Write("    ");
            writer.WriteLine(kvp.Key);
        }
        writer.WriteLine($"""  </cellXfs>""");

        writer.Write(
            """
              <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0" /></cellStyles>
              <dxfs count="0" />
              <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16" />
              <extLst>
                <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1" /></ext>
                <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1" /></ext>
              </extLst>
            </styleSheet>
            """);
    }

    private void writeMisc()
    {
        writeFile("[Content_Types].xml",
            $$"""
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
              <Default Extension="xml" ContentType="application/xml" />
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
            {{string.Join("\r\n", _wb.Sheets.Select((s, i) => $"""  <Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />"""))}}
              <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
              <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
              <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
              <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
            </Types>
            """);

        writeFile("_rels/.rels",
            """
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
            </Relationships>
            """);

        writeFile("docProps/app.xml",
            $$"""
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
              <Application>Microsoft Excel</Application>
              <DocSecurity>0</DocSecurity>
              <ScaleCrop>false</ScaleCrop>
              <HeadingPairs>
                <vt:vector size="2" baseType="variant">
                  <vt:variant>
                    <vt:lpstr>Worksheets</vt:lpstr>
                  </vt:variant>
                  <vt:variant>
                    <vt:i4>{{_wb.Sheets.Count}}</vt:i4>
                  </vt:variant>
                </vt:vector>
              </HeadingPairs>
              <TitlesOfParts>
                <vt:vector size="{{_wb.Sheets.Count}}" baseType="lpstr">
            {{string.Join("\r\n", _wb.Sheets.Select(s => $"""      <vt:lpstr>{SecurityElement.Escape(s.Name)}</vt:lpstr>"""))}}
                </vt:vector>
              </TitlesOfParts>
              <Company></Company>
              <LinksUpToDate>false</LinksUpToDate>
              <SharedDoc>false</SharedDoc>
              <HyperlinksChanged>false</HyperlinksChanged>
              <AppVersion>16.0300</AppVersion>
            </Properties>
            """);

        writeFile("docProps/core.xml",
            $$"""
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
              <dc:creator>{{SecurityElement.Escape(_wb.Creator)}}</dc:creator>
              <cp:lastModifiedBy>{{SecurityElement.Escape(_wb.LastModifiedBy)}}</cp:lastModifiedBy>
              <dcterms:created xsi:type="dcterms:W3CDTF">{{_wb.CreatedAt.ToUniversalTime():yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'}}</dcterms:created>
              <dcterms:modified xsi:type="dcterms:W3CDTF">{{_wb.ModifiedAt.ToUniversalTime():yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'}}</dcterms:modified>
            </cp:coreProperties>
            """);

        writeFile("xl/_rels/workbook.xml.rels",
            $$"""
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            {{string.Join("\r\n", _wb.Sheets.Select((s, i) => $"""  <Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml" />"""))}}
              <Relationship Id="rId{{_wb.Sheets.Count + 1}}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
              <Relationship Id="rId{{_wb.Sheets.Count + 2}}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
            </Relationships>
            """);

        writeFile("xl/theme/theme1.xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont><a:latin typeface="Aptos Narrow" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{2E142A2C-CD16-42D6-873A-C26D2A0506FA}" vid="{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}"/></a:ext></a:extLst></a:theme>
            """);
    }

    private void writeWorkbook()
    {
        string activeTab = "";
        if (_wb.ActiveSheet != null)
        {
            var index = _wb.Sheets.IndexOf(_wb.ActiveSheet);
            if (index < 0)
                throw new InvalidOperationException("ActiveTab not found in Sheets");
            activeTab = $" activeTab=\"{index}\"";
        }
        writeFile("xl/workbook.xml",
            $$"""
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
              <workbookPr defaultThemeVersion="202300" />
              <bookViews>
                <workbookView xWindow="-110" yWindow="-110" windowWidth="25820" windowHeight="14620"{{activeTab}} xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}" />
              </bookViews>
              <sheets>
            {{string.Join("\r\n", _wb.Sheets.Select((s, i) => $"""    <sheet name="{SecurityElement.Escape(s.Name)}" sheetId="{i + 1}" r:id="rId{i + 1}" />"""))}}
              </sheets>
              <calcPr calcId="191029" />
              <extLst>
                <ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                  <x15:workbookPr chartTrackingRefBase="1" />
                </ext>
                <ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">
                  <xcalcf:calcFeatures>
                    <xcalcf:feature name="microsoft.com:RD" />
                    <xcalcf:feature name="microsoft.com:Single" />
                    <xcalcf:feature name="microsoft.com:FV" />
                    <xcalcf:feature name="microsoft.com:CNMTM" />
                    <xcalcf:feature name="microsoft.com:LET_WF" />
                    <xcalcf:feature name="microsoft.com:LAMBDA_WF" />
                    <xcalcf:feature name="microsoft.com:ARRAYTEXT_WF" />
                  </xcalcf:calcFeatures>
                </ext>
              </extLst>
            </workbook>
            """);
    }

    #region Style mapping

    private static string[] _borderStyleStr = ["", "hair", "thin", "medium", "thick", "dotted", "dashed", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot", "double"];

    private Dictionary<string, int> _sxNumFmts = [];
    private Dictionary<string, int> _sxFontsXml = [];
    private Dictionary<string, int> _sxFillsXml = [];
    private Dictionary<string, int> _sxBordersXml = [];
    private Dictionary<string, int> _sxXfsXml = [];

    private void initialiseStyles()
    {
        var s = _wb.Style;
        void checkNotNull(object? value, string name) { if (value == null) throw new InvalidOperationException($"{name} must not be null in the default workbook style."); }
        checkNotNull(s.Format, nameof(s.Format));
        checkNotNull(s.Font, nameof(s.Font));
        checkNotNull(s.Size, nameof(s.Size));
        checkNotNull(s.Bold, nameof(s.Bold));
        checkNotNull(s.Italic, nameof(s.Italic));
        checkNotNull(s.Color, nameof(s.Color));
        checkNotNull(s.FillColor, nameof(s.FillColor));
        checkNotNull(s.Horz, nameof(s.Horz));
        checkNotNull(s.Vert, nameof(s.Vert));
        checkNotNull(s.Wrap, nameof(s.Wrap));
        checkNotNull(s.BrLeft, nameof(s.BrLeft));
        checkNotNull(s.BrLeftColor, nameof(s.BrLeftColor));
        checkNotNull(s.BrRight, nameof(s.BrRight));
        checkNotNull(s.BrRightColor, nameof(s.BrRightColor));
        checkNotNull(s.BrTop, nameof(s.BrTop));
        checkNotNull(s.BrTopColor, nameof(s.BrTopColor));
        checkNotNull(s.BrBot, nameof(s.BrBot));
        checkNotNull(s.BrBotColor, nameof(s.BrBotColor));
        if (s.FillColor != "") throw new NotSupportedException("A workbook-wide default fill is not supported by Excel.");
        if (s.BrLeft != XlBorder.None || s.BrRight != XlBorder.None || s.BrTop != XlBorder.None || s.BrBot != XlBorder.None) throw new NotSupportedException("A workbook-wide default border is not supported by Excel.");
        if (s.BrLeftColor != "" || s.BrRightColor != "" || s.BrTopColor != "" || s.BrBotColor != "") throw new NotSupportedException("A workbook-wide default border color is not supported by Excel.");

        _sxFontsXml.Add(makeFontXml(s.Font!, s.Size!.Value, s.Bold!.Value, s.Italic!.Value, s.Color!), 0);
        _sxFillsXml.Add(makeFillXml(s.FillColor!), 0);
        _sxFillsXml.Add("""<fill><patternFill patternType="gray125" /></fill>""", 1); // Excel wants this fill to be present
        _sxBordersXml.Add(makeBorderXml(s.BrLeft!.Value, s.BrLeftColor!, s.BrRight!.Value, s.BrRightColor!, s.BrTop!.Value, s.BrTopColor!, s.BrBot!.Value, s.BrBotColor!), 0);
        _sxXfsXml.Add(makeXfXml(0, 0, 0, 0, s.Horz!.Value, s.Vert!.Value, s.Wrap!.Value), 0);
    }

    private static string makeFontXml(string fontName, double fontSize, bool bold, bool italic, string color)
    {
        var sb = new StringBuilder();
        sb.Append("<font>");
        sb.Append($"""<name val="{SecurityElement.Escape(fontName)}"/><sz val="{fontSize:0.#}"/>""");
        if (color == "")
            sb.Append("""<color theme="1"/>""");
        else
            sb.Append($"""<color rgb="{(color.Length == 6 ? "FF" : "")}{color.ToUpper()}"/>""");
        if (bold) sb.Append("<b/>");
        if (italic) sb.Append("<i/>");
        sb.Append("</font>");
        return sb.ToString();
    }

    private static string makeFillXml(string fillColor)
    {
        if (fillColor == "")
            return """<fill><patternFill patternType="none" /></fill>""";
        if (fillColor.Length == 6)
            fillColor = "FF" + fillColor;
        return $"""<fill><patternFill patternType="solid"><fgColor rgb="{fillColor.ToUpper()}" /><bgColor indexed="64" /></patternFill></fill>""";
    }

    private static string makeBorderXml(XlBorder left, string leftCol, XlBorder right, string rightCol, XlBorder top, string topCol, XlBorder bot, string botCol)
    {
        static void addBorder(StringBuilder sb, string name, XlBorder bstyle, string bcolor)
        {
            sb.Append('<');
            sb.Append(name);
            if (bstyle == XlBorder.None)
                sb.Append("/>");
            else
            {
                sb.Append($""" style="{_borderStyleStr[(int)bstyle]}">""");
                if (bcolor == "")
                    sb.Append("""<color auto="1"/>""");
                else
                    sb.Append($"""<color rgb="{(bcolor.Length == 6 ? "FF" : "")}{bcolor.ToUpper()}"/>""");
                sb.Append($"</{name}>");
            }
        }
        var sb = new StringBuilder();
        sb.Append("<border>");
        addBorder(sb, "left", left, leftCol);
        addBorder(sb, "right", right, rightCol);
        addBorder(sb, "top", top, topCol);
        addBorder(sb, "bottom", bot, botCol);
        sb.Append("<diagonal/></border>");
        return sb.ToString();
    }

    private static string makeXfXml(int numFmtId, int fontId, int fillId, int borderId, XlHorz horz, XlVert vert, bool wrap)
    {
        var xf = new XElement("xf",
        new XAttribute("numFmtId", numFmtId),
            new XAttribute("fontId", fontId),
            new XAttribute("fillId", fillId),
            new XAttribute("borderId", borderId),
            new XAttribute("xfId", 0));

        if (numFmtId != 0) xf.Add(new XAttribute("applyNumberFormat", 1));
        if (fontId != 0) xf.Add(new XAttribute("applyFont", 1));
        if (fillId != 0) xf.Add(new XAttribute("applyFill", 1));
        if (borderId != 0) xf.Add(new XAttribute("applyBorder", 1));

        if (horz != XlHorz.Auto || vert != XlVert.Bottom || wrap)
        {
            var alignment = new XElement("alignment");
            if (horz != XlHorz.Auto) alignment.Add(new XAttribute("horizontal", horz.ToString().ToLower()));
            if (vert != XlVert.Bottom) alignment.Add(new XAttribute("vertical", vert.ToString().ToLower()));
            if (wrap) alignment.Add(new XAttribute("wrapText", 1));
            xf.Add(alignment);
            xf.Add(new XAttribute("applyAlignment", 1));
        }

        return xf.ToString(SaveOptions.DisableFormatting);
    }

    private static T nn<T>(T? v) where T : struct => v ?? throw new NullReferenceException();
    private static T nn<T>(T? v) where T : class => v ?? throw new NullReferenceException();

    /// <param name="s">Style with full inheritance applied.</param>
    internal int MapStyle(XlStyle s)
    {
        int numFmtId = XlFmt.StandardNumberFormatId(nn(s.Format));
        if (numFmtId < 0)
            numFmtId = getOrAddXmlId(_sxNumFmts, nn(s.Format), 164);

        var fontXml = makeFontXml(nn(s.Font), nn(s.Size), nn(s.Bold), nn(s.Italic), nn(s.Color));
        int fontId = getOrAddXmlId(_sxFontsXml, fontXml, 0);

        var fillXml = makeFillXml(nn(s.FillColor));
        int fillId = getOrAddXmlId(_sxFillsXml, fillXml, 0);

        var borderXml = makeBorderXml(nn(s.BrLeft), nn(s.BrLeftColor), nn(s.BrRight), nn(s.BrRightColor), nn(s.BrTop), nn(s.BrTopColor), nn(s.BrBot), nn(s.BrBotColor));
        int borderId = getOrAddXmlId(_sxBordersXml, borderXml, 0);

        var xfXml = makeXfXml(numFmtId, fontId, fillId, borderId, nn(s.Horz), nn(s.Vert), nn(s.Wrap));
        return getOrAddXmlId(_sxXfsXml, xfXml, 0);
    }

    private int getOrAddXmlId(Dictionary<string, int> xmls, string xml, int idOfFirst)
    {
        if (xmls.TryGetValue(xml, out var id))
            return id;
        id = idOfFirst + xmls.Count;
        xmls.Add(xml, id);
        return id;
    }

    #endregion
}
