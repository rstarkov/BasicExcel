﻿using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

namespace BasicExcel;

internal class XlWriter : IDisposable
{
    private XlWorkbook _wb;
    private ZipArchive _zip;

    public XlWriter(XlWorkbook workbook, Stream stream)
    {
        _wb = workbook;
        _zip = new ZipArchive(stream, ZipArchiveMode.Create);
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
        initialiseStyles(_wb.Style);

        // check / patchup sheets
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

        // write sheets first, collecting styles
        for (int si = 0; si < _wb.Sheets.Count; si++)
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
            if (_wb.ActiveSheet == _wb.Sheets[si])
                viewXml.Add(new XAttribute("tabSelected", "1"));
            //if (_wb.Sheets[si].Freeze != null)
            //{
            //    var paneXml = new XElement("pane", ?);
            //    viewXml.Add(paneXml);
            //}
            writer.Write("  ");
            writer.Write(new XElement("sheetViews", viewXml).ToString(SaveOptions.DisableFormatting));
            writer.WriteLine(
                """

                  <sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35" />
                  <sheetData>
                """);

            var sw = new XlSheetWriter(this, _wb.Sheets[si], writer);
            _wb.Sheets[si].WriteSheet(sw);
            sw.Finalise();

            writer.Write(
                """
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
                </worksheet>
                """);
        }

        // write styles
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

            writer.WriteLine($"""  <fills count="{_sxFontsXml.Count}">""");
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

        // write all the other files

        writeFile("[Content_Types].xml",
            $$"""
            <?xml version="1.0" encoding="utf-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
              <Default Extension="xml" ContentType="application/xml" />
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
            {{string.Join("", _wb.Sheets.Select((s, i) => $"""  <Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />""" + "\r\n"))}}
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
            {{string.Join("", _wb.Sheets.Select(s => $"""      <vt:lpstr>{SecurityElement.Escape(s.Name)}</vt:lpstr>""" + "\r\n"))}}
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
            {{string.Join("", _wb.Sheets.Select((s, i) => $"""  <Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml" />""" + "\r\n"))}}
              <Relationship Id="rId{{_wb.Sheets.Count + 1}}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
              <Relationship Id="rId{{_wb.Sheets.Count + 2}}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
            </Relationships>
            """);

        writeFile("xl/theme/theme1.xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont><a:latin typeface="Aptos Narrow" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{2E142A2C-CD16-42D6-873A-C26D2A0506FA}" vid="{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}"/></a:ext></a:extLst></a:theme>
            """);

        string activeTab = "";
        if (_wb.ActiveSheet != null)
        {
            var index = _wb.Sheets.IndexOf(_wb.ActiveSheet);
            if (index < 0)
                throw new InvalidOperationException("ActiveTab not found in Sheets");
            activeTab = $" activeTab=\"{index + 1}\"";
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
            {{string.Join("", _wb.Sheets.Select((s, i) => $"""    <sheet name="{SecurityElement.Escape(s.Name)}" sheetId="{i + 1}" r:id="rId{i + 1}" />""" + "\r\n"))}}
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

    private void initialiseStyles(XlStyle wbDefaultStyle)
    {
        var s = wbDefaultStyle;
        void throwErr(string name) => throw new InvalidOperationException($"{name} must not be null in the default workbook style.");
        if (s.Format == null) throwErr(nameof(XlStyle.Format));
        if (s.Font == null) throwErr(nameof(XlStyle.Font));
        if (s.Size == null) throwErr(nameof(XlStyle.Size));
        if (s.Bold == null) throwErr(nameof(XlStyle.Bold));
        if (s.Italic == null) throwErr(nameof(XlStyle.Italic));
        if (s.Color == null) throwErr(nameof(XlStyle.Color));
        if (s.FillColor == null) throwErr(nameof(XlStyle.FillColor));
        if (s.Horz == null) throwErr(nameof(XlStyle.Horz));
        if (s.Vert == null) throwErr(nameof(XlStyle.Vert));
        if (s.Wrap == null) throwErr(nameof(XlStyle.Wrap));
        if (s.BrLeft == null) throwErr(nameof(XlStyle.BrLeft));
        if (s.BrLeftColor == null) throwErr(nameof(XlStyle.BrLeftColor));
        if (s.BrRight == null) throwErr(nameof(XlStyle.BrRight));
        if (s.BrRightColor == null) throwErr(nameof(XlStyle.BrRightColor));
        if (s.BrTop == null) throwErr(nameof(XlStyle.BrTop));
        if (s.BrTopColor == null) throwErr(nameof(XlStyle.BrTopColor));
        if (s.BrBot == null) throwErr(nameof(XlStyle.BrBot));
        if (s.BrBotColor == null) throwErr(nameof(XlStyle.BrBotColor));

        _sxFontsXml.Add(makeFontXml(s.Font!, s.Size!.Value, s.Bold!.Value, s.Italic!.Value, s.Color!), 0);
        _sxFillsXml.Add(makeFillXml(s.FillColor!), 0);
        _sxFillsXml.Add("""<fill><patternFill patternType="gray125" /></fill>""", 1); // Excel wants this fill to be present
        _sxBordersXml.Add(makeBorderXml(s.BrLeft!.Value, s.BrLeftColor!, s.BrRight!.Value, s.BrRightColor!, s.BrTop!.Value, s.BrTopColor!, s.BrBot!.Value, s.BrBotColor!), 0);
        _sxXfsXml.Add(new XElement("xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0), new XAttribute("xfId", 0)).ToString(SaveOptions.DisableFormatting), 0);
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

    internal int MapStyle(XlStyle? style, XlStyle? sheetStyle)
    {
        string Format = style?.Format ?? sheetStyle?.Format ?? _wb.Style.Format!;
        string Font = style?.Font ?? sheetStyle?.Font ?? _wb.Style.Font!;
        double Size = style?.Size ?? sheetStyle?.Size ?? _wb.Style.Size!.Value;
        bool Bold = style?.Bold ?? sheetStyle?.Bold ?? _wb.Style.Bold!.Value;
        bool Italic = style?.Italic ?? sheetStyle?.Italic ?? _wb.Style.Italic!.Value;
        string Color = style?.Color ?? sheetStyle?.Color ?? _wb.Style.Color!;
        string FillColor = style?.FillColor ?? sheetStyle?.FillColor ?? _wb.Style.FillColor!;
        XlHorz Horz = style?.Horz ?? sheetStyle?.Horz ?? _wb.Style.Horz!.Value;
        XlVert Vert = style?.Vert ?? sheetStyle?.Vert ?? _wb.Style.Vert!.Value;
        bool Wrap = style?.Wrap ?? sheetStyle?.Wrap ?? _wb.Style.Wrap!.Value;
        XlBorder BrLeft = style?.BrLeft ?? sheetStyle?.BrLeft ?? _wb.Style.BrLeft!.Value;
        string BrLeftColor = style?.BrLeftColor ?? sheetStyle?.BrLeftColor ?? _wb.Style.BrLeftColor!;
        XlBorder BrRight = style?.BrRight ?? sheetStyle?.BrRight ?? _wb.Style.BrRight!.Value;
        string BrRightColor = style?.BrRightColor ?? sheetStyle?.BrRightColor ?? _wb.Style.BrRightColor!;
        XlBorder BrTop = style?.BrTop ?? sheetStyle?.BrTop ?? _wb.Style.BrTop!.Value;
        string BrTopColor = style?.BrTopColor ?? sheetStyle?.BrTopColor ?? _wb.Style.BrTopColor!;
        XlBorder BrBot = style?.BrBot ?? sheetStyle?.BrBot ?? _wb.Style.BrBot!.Value;
        string BrBotColor = style?.BrBotColor ?? sheetStyle?.BrBotColor ?? _wb.Style.BrBotColor!;

        int numFmtId = XlFmt.StandardNumberFormatId(Format);
        if (numFmtId < 0)
            numFmtId = getOrAddXmlId(_sxNumFmts, Format, 164);

        var fontXml = makeFontXml(Font, Size, Bold, Italic, Color);
        int fontId = getOrAddXmlId(_sxFontsXml, fontXml, 0);

        var fillXml = makeFillXml(FillColor);
        int fillId = getOrAddXmlId(_sxFillsXml, fillXml, 0);

        var borderXml = makeBorderXml(BrLeft, BrLeftColor, BrRight, BrRightColor, BrTop, BrTopColor, BrBot, BrBotColor);
        int borderId = getOrAddXmlId(_sxBordersXml, borderXml, 0);

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

        if (Horz != XlHorz.Left || Vert != XlVert.Bottom || Wrap)
        {
            var alignment = new XElement("alignment");
            if (Horz != XlHorz.Left) alignment.Add(new XAttribute("horizontal", Horz.ToString().ToLower()));
            if (Vert != XlVert.Bottom) alignment.Add(new XAttribute("vertical", Vert.ToString().ToLower()));
            if (Wrap) alignment.Add(new XAttribute("wrapText", 1));
            xf.Add(alignment);
            xf.Add(new XAttribute("applyAlignment", 1));
        }

        var xfXml = xf.ToString(SaveOptions.DisableFormatting);
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
