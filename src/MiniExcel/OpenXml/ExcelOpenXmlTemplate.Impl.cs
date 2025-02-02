using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlTemplate
    {
        public class XRowInfo
        {
            public string FormatText { get; set; }
            public string IEnumerablePropName { get; set; }
            public XmlElement Row { get; set; }
            public Type IEnumerableGenricType { get; set; }
            public IDictionary<string, PropInfo> PropsMap { get; set; }
            public bool IsDictionary { get; set; }
            public bool IsDataTable { get; set; }
            public int CellIEnumerableValuesCount { get; set; }
            public IList<object> CellIlListValues { get; set; }
            public IEnumerable CellIEnumerableValues { get; set; }
            public XMergeCell IEnumerableMercell { get; set; }
            public List<XMergeCell> RowMercells { get; set; }
        }

        public class PropInfo
        {
            public PropertyInfo PropertyInfo { get; set; }
            public Type UnderlyingTypePropType { get; set; }
        }

        public class XMergeCell
        {
            public XMergeCell(XMergeCell mergeCell)
            {
                Width = mergeCell.Width;
                Height = mergeCell.Height;
                X1 = mergeCell.X1;
                Y1 = mergeCell.Y1;
                X2 = mergeCell.X2;
                Y2 = mergeCell.Y2;
                MergeCell = mergeCell.MergeCell;
            }
            public XMergeCell(XmlNode mergeCell)
            {
                var @ref = mergeCell.Attributes["ref"].Value;
                var refs = @ref.Split(':');

                //TODO: width,height
                var xy1 = refs[0];
                X1 = ColumnHelper.GetColumnIndex(StringHelper.GetLetter(refs[0]));
                Y1 = StringHelper.GetNumber(xy1);

                var xy2 = refs[1];
                X2 = ColumnHelper.GetColumnIndex(StringHelper.GetLetter(refs[1]));
                Y2 = StringHelper.GetNumber(xy2);

                Width = Math.Abs(X1 - X2) + 1;
                Height = Math.Abs(Y1 - Y2) + 1;
            }
            public XMergeCell(string x1, int y1, string x2, int y2)
            {
                X1 = ColumnHelper.GetColumnIndex(x1);
                Y1 = y1;

                X2 = ColumnHelper.GetColumnIndex(x2);
                Y2 = y2;

                Width = Math.Abs(X1 - X2) + 1;
                Height = Math.Abs(Y1 - Y2) + 1;
            }

            public string XY1 { get { return $"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}"; } }
            public int X1 { get; set; }
            public int Y1 { get; set; }
            public string XY2 { get { return $"{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}"; } }
            public int X2 { get; set; }
            public int Y2 { get; set; }
            public string Ref { get { return $"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}:{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}"; } }
            public XmlElement MergeCell { get; set; }
            public int Width { get; internal set; }
            public int Height { get; internal set; }

            public string ToXmlString(string prefix)
            {
                return $"<{prefix}mergeCell ref=\"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}:{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}\"/>";
            }
        }

        private List<XRowInfo> XRowInfos { get; set; }

        private Dictionary<string, XMergeCell> XMergeCellInfos { get; set; }
        public List<XMergeCell> NewXMergeCellInfos { get; private set; }

        private void GenerateSheetXmlImpl(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream,
            IReadOnlyDictionary<string, object> inputMaps, IDictionary<int, string> sharedStrings, bool mergeCells = false)
        {
            var doc = new XmlDocument();
            doc.Load(sheetStream);
            sheetStream.Dispose();

            sheetZipEntry.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic

            var worksheet = doc.SelectSingleNode("/x:worksheet", _ns);
            var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", _ns);
            var newSheetData = sheetData!.Clone(); //avoid delete lost data
            var rows = newSheetData.SelectNodes("x:row", _ns);

            ReplaceSharedStringsToStr(sharedStrings, ref rows);
            GetMercells(doc, worksheet);
            UpdateDimensionAndGetRowsInfo(inputMaps, ref doc, ref rows, !mergeCells);
            WriteSheetXml(stream, doc, sheetData, mergeCells);
        }

        private void CopySheetXmlImpl(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream,
            bool mergeCells = false)
        {
            var doc = new XmlDocument();
            doc.Load(sheetStream);
            sheetStream.Dispose();

            sheetZipEntry.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic

            var worksheet = doc.SelectSingleNode("/x:worksheet", _ns);
            var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", _ns);
            var newSheetData = sheetData.Clone(); //avoid delete lost data
            var rows = newSheetData.SelectNodes("x:row", _ns);
            ReplaceSharedStringsToStr(new Dictionary<int, string>(), ref rows);
            GetMercells(doc, worksheet);
            Dictionary<string, object> inputMaps = new();
            UpdateDimensionAndGetRowsInfo(inputMaps, ref doc, ref rows, !mergeCells, ignoreMaps:true);
            WriteSheetXml(stream, doc, sheetData, mergeCells);
        }

        private void GetMercells(XmlDocument doc, XmlNode worksheet)
        {
            var mergeCells = doc.SelectSingleNode("/x:worksheet/x:mergeCells", _ns);

            if (mergeCells == null) return;

            var newMergeCells = mergeCells.Clone();
            worksheet.RemoveChild(mergeCells);

            foreach (var mergerCell in from XmlElement cell in newMergeCells
                     let _ = cell.Attributes["ref"]?.Value select new XMergeCell(cell))
            {
                XMergeCellInfos[mergerCell.XY1] = mergerCell;
            }
        }

        private class MergeCellIndex
        {
            public int RowStart { get; }
            public int RowEnd { get; }

            public MergeCellIndex(int rowStart, int rowEnd)
            {
                RowStart = rowStart;
                RowEnd = rowEnd;
            }
        }
        
        private class XChildNode
        {
            public string InnerText { get; init; }
            public string ColIndex { get; init; }
            public int RowIndex { get; init; }
        }

        private void WriteSheetXml(Stream stream, XmlNode doc, XmlNode sheetData, bool mergeCells = false)
        {
            //Q.Why so complex?
            //A.Because try to use string stream avoid OOM when rendering rows
            sheetData.RemoveAll();
            sheetData.InnerText = "{{{{{{split}}}}}}"; //TODO: bad code smell
            var prefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $"{sheetData.Prefix}:";
            var endPrefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $":{sheetData.Prefix}"; //![image](https://user-images.githubusercontent.com/12729184/115000066-fd02b300-9ed4-11eb-8e65-bf0014015134.png)
            var contents = doc.InnerXml.Split(new[] { $"<{prefix}sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</{prefix}sheetData>" }, StringSplitOptions.None);
            using var writer = new StreamWriter(stream, Encoding.UTF8);
            writer.Write(contents[0]);
            writer.Write($"<{prefix}sheetData>"); // prefix problem

            #region MergeCells
                
            if(mergeCells)
            {
                var columns = XRowInfos.SelectMany(s => s.Row.Cast<XmlElement>())
                    .Where(s => !string.IsNullOrEmpty(s.InnerText)).Select(s =>
                    {
                        var att = s.GetAttribute("r");
                        return new XChildNode
                        {
                            InnerText = s.InnerText,
                            ColIndex = StringHelper.GetLetter(att),
                            RowIndex = StringHelper.GetNumber(att)
                        };
                    }).ToList();
                    
                Dictionary<int, MergeCellIndex> lastMergeCellIndexes = new Dictionary<int, MergeCellIndex>();

                for (int rowNo = 0; rowNo < XRowInfos.Count; rowNo++)
                {
                    var rowInfo = XRowInfos[rowNo];
                    var row = rowInfo.Row;
                    var childNodes = row.ChildNodes.Cast<XmlElement>()
                        .Where(s => !string.IsNullOrEmpty(s.InnerText)).ToList();

                    foreach (var childNode in childNodes)
                    {
                        var childNodeAtt = StringHelper.GetLetter(childNode.GetAttribute("r"));

                        var xmlNodes = columns
                            .Where(j => j.InnerText == childNode.InnerText && j.ColIndex == childNodeAtt)
                            .OrderBy(s => s.RowIndex).ToList();

                        if (xmlNodes.Count > 1)
                        {
                            var firstRow = xmlNodes.FirstOrDefault();
                            var lastRow = xmlNodes.LastOrDefault(s => s.RowIndex <= firstRow?.RowIndex + xmlNodes.Count && s.RowIndex != firstRow?.RowIndex);
                                
                            if (firstRow != null && lastRow != null)
                            {
                                var mergeCell = new XMergeCell(firstRow.ColIndex, firstRow.RowIndex, lastRow.ColIndex, lastRow.RowIndex);
                                    
                                var mergeIndexResult = lastMergeCellIndexes.TryGetValue(mergeCell.X1, out var mergeIndex);

                                if (mergeIndexResult && mergeCell.Y1 >= mergeIndex.RowStart &&
                                    mergeCell.Y2 <= mergeIndex.RowEnd)
                                {
                                    continue;
                                }

                                lastMergeCellIndexes[mergeCell.X1] = new MergeCellIndex(mergeCell.Y1, mergeCell.Y2);

                                if (rowInfo.RowMercells == null)
                                {
                                    rowInfo.RowMercells = new List<XMergeCell>();
                                }
                                    
                                rowInfo.RowMercells.Add(mergeCell);
                            }
                        }
                    }
                }
            }
                
            #endregion

            #region Generate rows and cells

            var rowIndexDiff = 0;
            var rowXml = new StringBuilder();
                
            // for grouped cells
            var groupingStarted = false;
            var hasEverGroupStarted = false;
            var groupStartRowIndex = 0;
            IList<object> cellIEnumerableValues = null;
            var isCellIEnumerableValuesSet = false;
            var cellIEnumerableValuesIndex = 0;
            var groupRowCount = 0;
            var headerDiff = 0;
            var isFirstRound = true;
            var prevHeader = "";

            for (var rowNo = 0; rowNo < XRowInfos.Count; rowNo++)
            {
                var isHeaderRow = false;
                var currentHeader = "";
                    
                var rowInfo = XRowInfos[rowNo];
                var row = rowInfo.Row;

                if (row.InnerText.Contains("@group"))
                {
                    groupingStarted = true;
                    hasEverGroupStarted = true;
                    groupStartRowIndex = rowNo;
                    isFirstRound = true;
                    prevHeader = "";

                    continue;
                }

                if (row.InnerText.Contains("@endgroup"))
                {
                    if(cellIEnumerableValuesIndex >= cellIEnumerableValues.Count - 1)
                    {
                        groupingStarted = false;
                        groupStartRowIndex = 0;
                        cellIEnumerableValues = null;
                        isCellIEnumerableValuesSet = false;
                        headerDiff++;
                        continue;
                    }
                    rowNo = groupStartRowIndex;
                    cellIEnumerableValuesIndex++;
                    isFirstRound = false;
                    continue;
                }
                if (row.InnerText.Contains("@header"))
                {
                    isHeaderRow = true;
                }

                if (groupingStarted && !isCellIEnumerableValuesSet && rowInfo.CellIlListValues != null)
                {
                    cellIEnumerableValues = rowInfo.CellIlListValues;
                    isCellIEnumerableValuesSet = true;
                }

                var groupingRowDiff =
                    (hasEverGroupStarted ? (-1 + cellIEnumerableValuesIndex * groupRowCount - headerDiff) : 0);
                    
                if (groupingStarted)
                {
                    if (isFirstRound)
                    {
                        groupRowCount++;
                    }

                    if(cellIEnumerableValues != null)
                    {
                        rowInfo.CellIEnumerableValuesCount = 1;
                        rowInfo.CellIEnumerableValues =
                            cellIEnumerableValues.Skip(cellIEnumerableValuesIndex).Take(1).ToList();
                    }
                }

                //TODO: some xlsx without r
                var originRowIndex = int.Parse(row.GetAttribute("r"));
                var newRowIndex = originRowIndex + rowIndexDiff + groupingRowDiff;

                var innerXml = row.InnerXml;
                rowXml.Clear()
                    .Append($@"<{row.Name}");
                foreach (var attr in row.Attributes.Cast<XmlAttribute>()
                             .Where(e => e.Name != "r"))
                {
                    rowXml.Append($@" {attr.Name}=""{attr.Value}""");
                }

                var outerXmlOpen = rowXml.ToString();

                if (rowInfo.CellIEnumerableValues != null)
                {
                    var first = true;
                    var iEnumerableIndex = 0;

                    foreach (var item in rowInfo.CellIEnumerableValues)
                    {
                        iEnumerableIndex++;

                        rowXml.Clear()
                            .Append(outerXmlOpen)
                            .Append($@" r=""{newRowIndex}"">")
                            .Append(innerXml)
                            .Replace("{{$rowindex}}", newRowIndex.ToString())
                            .Append($@"</{row.Name}>");
                        if (iEnumerableIndex > 1 && rowInfo.PropsMap.Count > 0)
                        {
                            rowXml = ShiftFormulasBelow(rowXml, iEnumerableIndex - 1);
                        }

                        if (rowInfo.IsDictionary)
                        {
                            var dic = item as IDictionary<string, object>;
                            foreach (var propInfo in rowInfo.PropsMap)
                            {
                                var key = $"{{{{{rowInfo.IEnumerablePropName}.{propInfo.Key}}}}}";
                                if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                {
                                    rowXml.Replace(key, "");
                                    continue;
                                }

                                var cellValue = dic[propInfo.Key];
                                if (cellValue == null)
                                {
                                    rowXml.Replace(key, "");
                                    continue;
                                }


                                var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue?.ToString());
                                var type = propInfo.Value.UnderlyingTypePropType;
                                if (type == typeof(bool))
                                {
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                else if (type == typeof(DateTime))
                                {
                                    cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                                }

                                //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)
                                    
                                rowXml.Replace("@header" + key, cellValueStr);
                                rowXml.Replace(key, cellValueStr);
                                    
                                if(isHeaderRow && row.InnerText.Contains(key))
                                {
                                    currentHeader += cellValueStr;
                                }
                            }
                        }
                        else if (rowInfo.IsDataTable)
                        {
                            var dataRow = item as DataRow;
                            foreach (var propInfo in rowInfo.PropsMap)
                            {
                                var key = $"{{{{{rowInfo.IEnumerablePropName}.{propInfo.Key}}}}}";
                                if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                {
                                    rowXml.Replace(key, "");
                                    continue;
                                }

                                var cellValue = dataRow[propInfo.Key];
                                if (cellValue == null)
                                {
                                    rowXml.Replace(key, "");
                                    continue;
                                }


                                var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue?.ToString());
                                var type = propInfo.Value.UnderlyingTypePropType;
                                if (type == typeof(bool))
                                {
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                else if (type == typeof(DateTime))
                                {
                                    cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                                }

                                //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)
                                    
                                rowXml.Replace("@header" + key, cellValueStr);
                                rowXml.Replace(key, cellValueStr);
                                    
                                if(isHeaderRow && row.InnerText.Contains(key))
                                {
                                    currentHeader += cellValueStr;
                                }
                            }
                        }
                        else
                        {
                            foreach (var propInfo in rowInfo.PropsMap)
                            {
                                var prop = propInfo.Value.PropertyInfo;

                                var key = $"{{{{{rowInfo.IEnumerablePropName}.{prop.Name}}}}}";
                                if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                {
                                    rowXml.Replace(key, "");
                                    continue;
                                }

                                var cellValue = prop.GetValue(item);
                                if (cellValue == null)
                                {
                                    rowXml.Replace(key, "");
                                    continue;
                                }
                                    
                                var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue?.ToString());
                                var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                if (type == typeof(bool))
                                {
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                else if (type == typeof(DateTime))
                                {
                                    cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                                }
                                else if (TypeHelper.IsNumericType(type))
                                {
                                    if (decimal.TryParse(cellValueStr, out var decimalValue))
                                        cellValueStr = decimalValue.ToString(CultureInfo.InvariantCulture);
                                }

                                //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)
                                    
                                rowXml.Replace("@header" + key, cellValueStr);
                                rowXml.Replace(key, cellValueStr);
                                    
                                if(isHeaderRow && row.InnerText.Contains(key))
                                {
                                    currentHeader += cellValueStr;
                                }
                            }
                        }

                        if (isHeaderRow)
                        {
                            if(currentHeader == prevHeader)
                            {
                                headerDiff++;
                                continue;
                            }

                            prevHeader = currentHeader;
                        }

                        // note: only first time need add diff ![image](https://user-images.githubusercontent.com/12729184/114494728-6bceda80-9c4f-11eb-9685-8b5ed054eabe.png)
                        if (!first)
                            rowIndexDiff += rowInfo.IEnumerableMercell?.Height ?? 1; //TODO:base on the merge size
                        first = false;

                        var mergeBaseRowIndex = newRowIndex;
                        newRowIndex += rowInfo.IEnumerableMercell?.Height ?? 1;
                        writer.Write(CleanXml(rowXml.ToString().Replace("xmlns:x=\"urn:schemas-microsoft-com:office:excel\"", ""), endPrefix)); // pass StringBuilder for netcoreapp3.0 or above

                        //mergecells
                        if (rowInfo.RowMercells == null) continue;

                        foreach (var newMergeCell in rowInfo.RowMercells.Select(mergeCell => new XMergeCell(mergeCell)))
                        {
                            newMergeCell.Y1 = newMergeCell.Y1 + rowIndexDiff + groupingRowDiff;
                            newMergeCell.Y2 = newMergeCell.Y2 + rowIndexDiff + groupingRowDiff;
                            NewXMergeCellInfos.Add(newMergeCell);
                        }

                        // Last merge one don't add new row, or it'll get duplicate result like : https://github.com/shps951023/MiniExcel/issues/207#issuecomment-824550950
                        if (iEnumerableIndex == rowInfo.CellIEnumerableValuesCount)
                            continue;

                        if (rowInfo.IEnumerableMercell != null)
                            continue;

                        // https://github.com/shps951023/MiniExcel/issues/207#issuecomment-824518897
                        for (int i = 1; i < rowInfo.IEnumerableMercell.Height; i++)
                        {
                            mergeBaseRowIndex++;
                            var _newRow = row.Clone() as XmlElement;
                            _newRow.SetAttribute("r", mergeBaseRowIndex.ToString());

                            var cs = _newRow.SelectNodes("x:c", _ns);
                            // all v replace by empty
                            // TODO: remove c/v
                            foreach (XmlElement _c in cs)
                            {
                                _c.RemoveAttribute("t");
                                foreach (XmlNode ch in _c.ChildNodes)
                                {
                                    _c.RemoveChild(ch);
                                }
                            }

                            _newRow.InnerXml = new StringBuilder(_newRow.InnerXml).Replace("{{$rowindex}}", mergeBaseRowIndex.ToString()).ToString();
                            writer.Write(CleanXml(_newRow.OuterXml, endPrefix));
                        }
                    }
                }
                else
                {
                    rowXml.Clear()
                        .Append(outerXmlOpen)
                        .Append($@" r=""{newRowIndex}"">")
                        .Append(innerXml)
                        .Replace("{{$rowindex}}", newRowIndex.ToString())
                        .Append($"</{row.Name}>");
                    writer.Write(CleanXml(rowXml, endPrefix)); // pass StringBuilder for netcoreapp3.0 or above

                    //mergecells
                    if (rowInfo.RowMercells != null)
                    {
                        foreach (var mergeCell in rowInfo.RowMercells)
                        {
                            var newMergeCell = new XMergeCell(mergeCell);
                            newMergeCell.Y1 = newMergeCell.Y1 + rowIndexDiff + groupingRowDiff;
                            newMergeCell.Y2 = newMergeCell.Y2 + rowIndexDiff + groupingRowDiff;
                            NewXMergeCellInfos.Add(newMergeCell);
                        }
                    }

                }

                // get the row's all mergecells then update the rowindex
            }
            #endregion

            writer.Write($"</{prefix}sheetData>");

            if (NewXMergeCellInfos.Count != 0)
            {
                writer.Write($"<{prefix}mergeCells count=\"{NewXMergeCellInfos.Count}\">");
                foreach (var cell in NewXMergeCellInfos)
                {
                    writer.Write(cell.ToXmlString(prefix));
                }
                writer.Write($"</{prefix}mergeCells>");
            }

            writer.Write(contents[1]);
        }

        public static Stream ToStream(string str)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(str);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        private static StringBuilder ShiftFormulasBelow(StringBuilder rowXml, int shiftTo)
        {
            var doc = new XmlDocument();
            var settings = new XmlReaderSettings { NameTable = new NameTable() };
            var xmlns = new XmlNamespaceManager(settings.NameTable);
            xmlns.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            xmlns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            xmlns.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            xmlns.AddNamespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            xmlns.AddNamespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            xmlns.AddNamespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            xmlns.AddNamespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            xmlns.AddNamespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            xmlns.AddNamespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            xmlns.AddNamespace("x", "urn:schemas-microsoft-com:office:excel");
            var context = new XmlParserContext(null, xmlns, "", XmlSpace.Default);
            var reader = XmlReader.Create(ToStream(rowXml.ToString()), settings, context);
            doc.Load(reader);

            foreach (var o in doc)
            {
                if (o is not XmlElement { Name: "row" } element) continue;
                foreach (var elementChildNode in element.ChildNodes)
                {
                    if (elementChildNode is not XmlElement { Name: "c" } cell) continue;
                    foreach (var cellChildNode in cell.ChildNodes)
                    {
                        if (cellChildNode is XmlElement { Name: "f" } childElement)
                        {
                            childElement.InnerText = UpdateFormula(childElement.InnerText, shiftTo);
                        }
                    }
                }
            }

            MemoryStream ms = new();
            doc.Save(ms);
            var docStr = ReadToEnd(ms);
            return new StringBuilder(docStr);
        }
        private static string ReadToEnd(Stream str)
        {
            str.Position = 0;
            var r = new StreamReader(str);
            return r.ReadToEnd();
        }

        private const string RxCell = "(?<prefix>\\$|\\!)?\\b(?<cell>(?<let>[A-Z])(?<num>[0-9])+)";

        private static string UpdateFormula(string childElementInnerText, int shiftTo)
        {
            var rx = new Regex(RxCell);
            var newText = rx.Replace(childElementInnerText, m =>
            {
                if (m.Groups.ContainsKey("prefix") && m.Groups["prefix"].Value.Length > 0)
                    return m.Value;
                var num = int.Parse(m.Groups["num"].Value);
                return m.Groups["let"].Value + (num + shiftTo);
            });
            return newText;
        }

        private static string ConvertToDateTimeString(KeyValuePair<string, PropInfo> propInfo, object cellValue)
        {
            //TODO:c.SetAttribute("t", "d"); and custom format
            string format;
            if (propInfo.Value.PropertyInfo == null)
            {
                format = "yyyy-MM-dd HH:mm:ss";
            }
            else
            {
                format = propInfo.Value.PropertyInfo.GetAttributeValue((ExcelFormatAttribute x) => x.Format)
                             ?? propInfo.Value.PropertyInfo.GetAttributeValue((ExcelColumnAttribute x) => x.Format)
                             ?? "yyyy-MM-dd HH:mm:ss";
            }

            return (cellValue as DateTime?)?.ToString(format);
        }

        private static StringBuilder CleanXml(StringBuilder xml, string endPrefix)
        {
            return xml
               .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
               .Replace($"xmlns{endPrefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
        }

        private static string CleanXml(string xml, string endPrefix)
        {
            //TODO: need to optimize
            return CleanXml(new StringBuilder(xml), endPrefix)
                .ToString();
        }

        private static void ReplaceSharedStringsToStr(IDictionary<int, string> sharedStrings, ref XmlNodeList rows)
        {
            foreach (XmlElement row in rows)
            {
                var columns = row.SelectNodes("x:c", _ns);
                if (columns == null)
                    continue;
                foreach (XmlElement c in columns)
                {
                    var t = c.GetAttribute("t");
                    var v = c.SelectSingleNode("x:v", _ns);
                    // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                    if (v?.InnerText == null) //![image](https://user-images.githubusercontent.com/12729184/114363496-075a3f80-9bab-11eb-9883-8e3fec10765c.png)
                        continue;

                    if (t != "s") continue;
                    //need to check sharedstring exist or not
                    if (!sharedStrings.ContainsKey(int.Parse(v.InnerText))) continue;
                    v.InnerText = sharedStrings[int.Parse(v.InnerText)];
                    // change type = str and replace its value
                    c.SetAttribute("t", "str");
                    //TODO: remove sharedstring? 
                }
            }
        }

        private void UpdateDimensionAndGetRowsInfo(
            IReadOnlyDictionary<string, object> inputMaps, ref XmlDocument doc, ref XmlNodeList rows, bool changeRowIndex = true, bool ignoreMaps = false)
        {
            // note : dimension need to put on the top ![image](https://user-images.githubusercontent.com/12729184/114507911-5dd88400-9c66-11eb-94c6-82ed7bdb5aab.png)

            if (doc.SelectSingleNode("/x:worksheet/x:dimension", _ns) is not XmlElement dimension)
            {
                dimension = doc.CreateElement("dimension", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var attr = doc.CreateAttribute("ref");
                attr.Value = "A1:V1000";
                dimension.Attributes.Append(attr);

                if (doc.SelectSingleNode("/x:worksheet/x:autoFilter", _ns)! is XmlElement autoFilter)
                {
                    doc.SelectSingleNode("/x:worksheet", _ns)!.RemoveChild(autoFilter);
                }

                // doc.SelectSingleNode("/x:worksheet", _ns)!.PrependChild(dimension);
            }

            var maxRowIndexDiff = 0;

            foreach (XmlElement row in rows)
            {
                // ==== get ienumerable infomation & maxrowindexdiff ====
                //Type ienumerableGenricType = null;
                //IDictionary<string, PropertyInfo> props = null;
                //IEnumerable ienumerable = null;
                var xRowInfo = new XRowInfo
                {
                    Row = row
                };
                XRowInfos.Add(xRowInfo);
                foreach (XmlElement c in row.SelectNodes("x:c", _ns))
                {
                    var r = c.GetAttribute("r");

                    // ==== mergecells ====
                    if (XMergeCellInfos.TryGetValue(r, out var cellInfo))
                    {
                        xRowInfo.RowMercells ??= new List<XMergeCell>();
                        xRowInfo.RowMercells.Add(cellInfo);
                    }

                    if(changeRowIndex)
                    {
                        c.SetAttribute("r", $"{StringHelper.GetLetter(r)}{{{{$rowindex}}}}"); //TODO:
                    }

                    var v = c.SelectSingleNode("x:v", _ns);
                    var f = c.SelectSingleNode("x:f", _ns);
                    if (v?.InnerText == null || f != null)
                        continue;

                    var matches = IsExpressionRegex.Matches(v.InnerText).GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value).ToArray();
                    var matchCnt = matches.Length;
                    var isMultiMatch = matchCnt > 1 || (matchCnt == 1 && v.InnerText != $"{{{{{matches[0]}}}}}");
                    foreach (var formatText in matches)
                    {
                        xRowInfo.FormatText = formatText;
                        var propNames = formatText.Split('.');
                        if (propNames[0].StartsWith("$")) //e.g:"$rowindex" it doesn't need to check cell value type
                            continue;

                        // TODO: default if not contain property key, clean the template string
                        if (!inputMaps.ContainsKey(propNames[0]))
                        {
                            if (!_configuration.IgnoreTemplateParameterMissing)
                                throw new KeyNotFoundException(
                                    $"Please check {propNames[0]} parameter, it's not exist.");

                            if (!ignoreMaps)
                                v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", "");
                            break;

                        }

                        var cellValue = inputMaps[propNames[0]]; // 1. From left to right, only the first set is used as the basis for the list
                        if (cellValue is (IEnumerable or IList<object>) and not string)
                        {
                            if (XMergeCellInfos.TryGetValue(r, out var info))
                            {
                                xRowInfo.IEnumerableMercell ??= info;
                            }
                            
                            xRowInfo.CellIEnumerableValues = cellValue as IEnumerable;
                            xRowInfo.CellIlListValues = cellValue as IList<object>;

                            // get ienumerable runtime type
                            if (xRowInfo.IEnumerableGenricType == null) //avoid duplicate to add rowindexdiff ![image](https://user-images.githubusercontent.com/12729184/114851348-522ac000-9e14-11eb-8244-4730754d6885.png)
                            {
                                var first = true;
                                //TODO:if CellIEnumerableValues is ICollection or Array then get length or Count

                                foreach (var element in xRowInfo.CellIEnumerableValues) //TODO: optimize performance?
                                {
                                    xRowInfo.CellIEnumerableValuesCount++;

                                    if (xRowInfo.IEnumerableGenricType == null)
                                        if (element != null)
                                        {
                                            xRowInfo.IEnumerablePropName = propNames[0];
                                            xRowInfo.IEnumerableGenricType = element.GetType();
                                            if (element is IDictionary<string, object> dic)
                                            {
                                                xRowInfo.IsDictionary = true;
                                                xRowInfo.PropsMap = dic.Keys.ToDictionary(key => key, key => dic.ContainsKey(key) && dic[key] != null
                                                    ? new PropInfo { UnderlyingTypePropType = Nullable.GetUnderlyingType(dic[key].GetType()) ?? dic[key].GetType() }
                                                    : new PropInfo { UnderlyingTypePropType = typeof(object) });
                                            }
                                            else
                                            {
                                                xRowInfo.PropsMap = xRowInfo.IEnumerableGenricType.GetProperties()
                                                    .ToDictionary(s => s.Name, s => new PropInfo { PropertyInfo = s, UnderlyingTypePropType = Nullable.GetUnderlyingType(s.PropertyType) ?? s.PropertyType });
                                            }
                                        }
                                    // ==== get demension max rowindex ====
                                    if (!first) //avoid duplicate add first one, this row not add status  ![image](https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                        maxRowIndexDiff += xRowInfo.IEnumerableMercell?.Height ?? 1;
                                    first = false;
                                }
                            }

                            //TODO: check if not contain 1 index
                            //only check first one match IEnumerable, so only render one collection at same row

                            // Empty collection parameter will get exception  https://gitee.com/dotnetchina/MiniExcel/issues/I4WM67
                            if (xRowInfo.PropsMap == null)
                            {
                                v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", propNames[1]);
                                break;
                            }
                            // auto check type https://github.com/shps951023/MiniExcel/issues/177
                            if (xRowInfo.PropsMap.ContainsKey(propNames[1]))
                            {
                                var prop = xRowInfo.PropsMap[propNames[1]];
                                var type = prop.UnderlyingTypePropType; //avoid nullable 
                                // 
                                //if (!xRowInfo.PropsMap.ContainsKey(propNames[1]))
                                //throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");

                                if (isMultiMatch)
                                {
                                    c.SetAttribute("t", "str");
                                }
                                else if (TypeHelper.IsNumericType(type))
                                {
                                    c.SetAttribute("t", "n");
                                }
                                else switch (Type.GetTypeCode(type))
                                {
                                    case TypeCode.Boolean:
                                        c.SetAttribute("t", "b");
                                        break;
                                    case TypeCode.DateTime:
                                        c.SetAttribute("t", "str");
                                        break;
                                }
                            }

                            break;
                        }

                        if (cellValue is DataTable dt)
                        {
                            if (xRowInfo.CellIEnumerableValues == null)
                            {
                                xRowInfo.IEnumerablePropName = propNames[0];
                                xRowInfo.IEnumerableGenricType = typeof(DataRow);
                                xRowInfo.IsDataTable = true;
                                xRowInfo.CellIEnumerableValues = dt.Rows.Cast<object>().ToList(); //TODO: need to optimize performance
                                xRowInfo.CellIlListValues = dt.Rows.Cast<object>().ToList();
                                var first = true;
                                foreach (var element in xRowInfo.CellIEnumerableValues)
                                {
                                    // ==== get demension max rowindex ====
                                    if (!first) //avoid duplicate add first one, this row not add status  ![image](https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                        maxRowIndexDiff++;
                                    first = false;
                                }
                                //TODO:need to optimize
                                //maxRowIndexDiff = dt.Rows.Count <= 1 ? 0 : dt.Rows.Count-1;
                                xRowInfo.PropsMap = dt.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col =>
                                    new PropInfo { UnderlyingTypePropType = Nullable.GetUnderlyingType(col.DataType) }
                                );
                            }

                            var column = dt.Columns[propNames[1]];
                            var type = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType; //avoid nullable 
                            if (!xRowInfo.PropsMap.ContainsKey(propNames[1]))
                                throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");

                            if (isMultiMatch)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (TypeHelper.IsNumericType(type))
                            {
                                c.SetAttribute("t", "n");
                            }
                            else switch (Type.GetTypeCode(type))
                            {
                                case TypeCode.Boolean:
                                    c.SetAttribute("t", "b");
                                    break;
                                case TypeCode.DateTime:
                                    c.SetAttribute("t", "str");
                                    break;
                            }
                        }
                        else
                        {
                            var cellValueStr = cellValue?.ToString(); /* value did encodexml, so don't duplicate encode value https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN*/
                            if (isMultiMatch) // if matchs count over 1 need to set type=str ![image](https://user-images.githubusercontent.com/12729184/114530109-39d46d00-9c7d-11eb-8f6b-52ad8600aca3.png)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (decimal.TryParse(cellValueStr, out var outV))
                            {
                                c.SetAttribute("t", "n");
                                cellValueStr = outV.ToString(CultureInfo.InvariantCulture);
                            }
                            else switch (cellValue)
                            {
                                case bool value:
                                    c.SetAttribute("t", "b");
                                    cellValueStr = value ? "1" : "0";
                                    break;
                                case DateTime time:
                                    cellValueStr = time.ToString("yyyy-MM-dd HH:mm:ss");
                                    break;
                            }

                            v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", cellValueStr); //TODO: auto check type and set value
                        }
                    }
                    //if (xRowInfo.CellIEnumerableValues != null) //2. From left to right, only the first set is used as the basis for the list
                    //    break;
                }
            }

            // e.g <dimension ref=\"A1:B6\" /> only need to update B6 to BMaxRowIndex
            var refs = dimension.GetAttribute("ref").Split(':');
            if (refs.Length == 2)
            {
                var letter = new string(refs[1].Where(char.IsLetter).ToArray());
                var digit = int.Parse(new string(refs[1].Where(char.IsDigit).ToArray()));

                dimension.SetAttribute("ref", $"{refs[0]}:{letter}{digit + maxRowIndexDiff}");
            }
            else
            {
                var letter = new string(refs[0].Where(Char.IsLetter).ToArray());
                var digit = int.Parse(new(refs[0].Where(char.IsDigit).ToArray()));

                dimension.SetAttribute("ref", $"A1:{letter}{digit + maxRowIndexDiff}");
            }
        }
    }
}