using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SW.ExportToExcel
{
    public static class IEnumerableExtensions
    {
        private static readonly XNamespace relationshipsNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static readonly XNamespace mainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string styleSheet = @"
            <styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
                <fonts count=""1"">
                    <font>
                        <sz val=""11""/>
                        <color theme=""1""/>
                        <name val=""Calibri""/>
                        <family val=""2""/>
                        <scheme val=""minor""/>
                    </font>
                </fonts>
                <fills count=""2"">
                    <fill><patternFill patternType=""none""/>
                    </fill><fill><patternFill patternType=""gray125""/>
                    </fill></fills>
                <borders count=""1"">
                    <border>
                        <left/><right/><top/>
                        <bottom/><diagonal/>
                    </border>
                </borders>
                <cellStyleXfs count=""1"">
                    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/>
                </cellStyleXfs>
                <cellXfs count=""2"">
                    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>
                    <xf numFmtId=""22"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1""/>
                </cellXfs><cellStyles count=""1"">
                    <cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/>
                </cellStyles>
                <dxfs count=""0""/>
                <tableStyles count=""0"" defaultTableStyle=""TableStyleMedium9"" defaultPivotStyle=""PivotStyleLight16""/>
            </styleSheet>";

        public static async Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data)
        {
            var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
            return await ExportToExcel(data, dictionary);
        }

        public static async Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data,
            IEnumerable<string> columns)
        {
            var dictionary = columns.ToDictionary(k => k, v => v);
            return await ExportToExcel(data, dictionary);
        }

        async public static Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data,
            IDictionary<string, string> columns)
        {
            var tempFile = Path.GetTempFileName();

            try
            {
                await WriteExcel(data, tempFile, columns);

                return File.ReadAllBytes(tempFile);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        public static async Task WriteExcel(this IEnumerable<Dictionary<string, string>> data, Stream stream)
        {
            var headers = data.FirstOrDefault().Select(x => x.Key).ToList();
            var dataRows = data.Select(d => d.Select(p => p.Value).ToList()).ToList();
            var res = GenerateExcel(headers, dataRows);
            await stream.WriteAsync(res, 0, res.Length);
        }


        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath)
        {
            var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
            await WriteExcel(data, filePath, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath,
            IEnumerable<string> columns)
        {
            var dictionary = columns.ToDictionary(k => k, v => v);
            await WriteExcel(data, filePath, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath,
            IDictionary<string, string> columns)
        {
            using (var fileStream = File.Open(filePath, FileMode.Create, FileAccess.ReadWrite))
            {
                await WriteExcel(data, fileStream, columns);
            }
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream)
        {
            var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
            await WriteExcel(data, stream, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream,
            IEnumerable<string> columns)
        {
            var dictionary = columns.ToDictionary(k => k, v => v);
            await WriteExcel(data, stream, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream,
            IDictionary<string, string> columns)
        {
            using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheetPartId = workbookPart.GetIdOfPart(worksheetPart);
                var stylesheetPart = workbookPart.AddNewPart<WorkbookStylesPart>();

                // create the string table
                var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                var xmlStringTable = new XElement(mainNamespace + "sst");
                await WriteXmlToPartAsync(sharedStringTablePart, xmlStringTable);

                var workbookElement = new XElement(mainNamespace + "workbook",
                    new XAttribute("xmlns", mainNamespace.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", relationshipsNamespace.NamespaceName),
                    new XElement("bookViews",
                        new XElement("workbookView")),
                    new XElement("sheets",
                        new XElement("sheet",
                            new XAttribute("name", "Exported"),
                            new XAttribute("sheetId", "1"),
                            new XAttribute(relationshipsNamespace + "id", worksheetPartId))));

                foreach (var element in workbookElement.Descendants())

                    if (element.Name.Namespace == "")
                    {
                        element.Attributes("xmlns").Remove();
                        element.Name = element.Parent.Name.Namespace + element.Name.LocalName;
                    }

                await WriteXmlToPartAsync(workbookPart, workbookElement);

                var propertyInfos = new List<PropertyInfo>();
                foreach (var propertyInfo in typeof(TEntity).GetProperties())

                    switch (true)
                    {
                        case object _ when propertyInfo.PropertyType == typeof(string):
                        case object _ when propertyInfo.PropertyType == typeof(DateTime):
                        case object _ when propertyInfo.PropertyType == typeof(DateTime?):
                        case object _ when propertyInfo.PropertyType == typeof(int):
                        case object _ when propertyInfo.PropertyType == typeof(int?):
                        case object _ when propertyInfo.PropertyType == typeof(byte):
                        case object _ when propertyInfo.PropertyType == typeof(byte?):
                        case object _ when propertyInfo.PropertyType == typeof(short):
                        case object _ when propertyInfo.PropertyType == typeof(short?):
                        case object _ when propertyInfo.PropertyType == typeof(long):
                        case object _ when propertyInfo.PropertyType == typeof(long?):
                        case object _ when propertyInfo.PropertyType == typeof(float):
                        case object _ when propertyInfo.PropertyType == typeof(float?):
                        case object _ when propertyInfo.PropertyType == typeof(double):
                        case object _ when propertyInfo.PropertyType == typeof(double?):
                        case object _ when propertyInfo.PropertyType == typeof(bool):
                        case object _ when propertyInfo.PropertyType == typeof(bool?):
                        case object _ when propertyInfo.PropertyType == typeof(decimal):
                        case object _ when propertyInfo.PropertyType == typeof(decimal?):
                            if (columns.ContainsKey(propertyInfo.Name))
                                propertyInfos.Add(propertyInfo);
                            break;
                    }

                var worksheetElement = new XElement(mainNamespace + "worksheet",
                    new XElement("sheetViews",
                        new XElement("sheetView",
                            new XAttribute("tabSelected", "1"),
                            new XAttribute("workbookViewId", "0")),
                        new XElement("pane",
                            new XAttribute("ySplit", "1"),
                            new XAttribute("topLeftCell", "A2"),
                            new XAttribute("activePane", "bottomLeft"),
                            new XAttribute("state", "frozen")),
                        new XElement("selection",
                            new XAttribute("pane", "bottomLeft"))),
                    new XElement("sheetFormatPr",
                        new XAttribute("defaultRowHeight", "15")),
                    new XElement("cols", GetCols(propertyInfos)),
                    new XElement("sheetData", GetRowValues(data, columns, propertyInfos)));

                foreach (var element in worksheetElement.Descendants())

                    if (element.Name.Namespace == "")
                    {
                        element.Attributes("xmlns").Remove();
                        element.Name = element.Parent.Name.Namespace + element.Name.LocalName;
                    }

                await WriteXmlToPartAsync(worksheetPart, worksheetElement);

                var styleSheetElement = XElement.Parse(styleSheet);

                await WriteXmlToPartAsync(stylesheetPart, styleSheetElement);

                await stream.FlushAsync();
                //stream.Close();  
            }
        }

        private static XElement[] GetRowValues<TEntity>(IEnumerable<TEntity> data,
            IDictionary<string, string> propertyDictionary, IReadOnlyCollection<PropertyInfo> Props)
        {
            var result = new List<XElement>();

            var rowNames = new XElement("row", from p in Props
                select new XElement("c", new XAttribute("t", "inlineStr"),
                    new XElement("is", new XElement("t", propertyDictionary[p.Name]))));
            result.Add(rowNames);

            result.AddRange(data.Select(item =>
                new XElement("row", from _p in Props select BuildCell(_p, _p.GetValue(item, null)))));
            return result.ToArray();
        }

        private static XElement[] GetCols(IEnumerable<PropertyInfo> propertyInfos)
        {
            return propertyInfos.Select(p => new XElement("col",
                new XAttribute("min", "1"),
                new XAttribute("max", "1"),
                new XAttribute("bestFit", "1"),
                new XAttribute("width", "4"))).ToArray();
        }

        private static XElement BuildCell(PropertyInfo propertyInfo, object value)
        {
            XElement element;

            switch (true)
            {
                case object _ when propertyInfo.PropertyType == typeof(string):
                case object _ when propertyInfo.PropertyType == typeof(Guid):
                case object _ when propertyInfo.PropertyType == typeof(Guid?):
                case object _ when propertyInfo.PropertyType == typeof(bool):
                case object _ when propertyInfo.PropertyType == typeof(bool?):

                    element = new XElement("c",
                        new XAttribute("t", "inlineStr"),
                        new XElement("is",
                            new XElement("t", value)));


                    return element;


                case object _ when propertyInfo.PropertyType == typeof(DateTime):
                case object _ when propertyInfo.PropertyType == typeof(DateTime):

                    
                    if (Convert.ToDateTime(value).Year >= 100)
                    {
                        element = new XElement("c",
                            new XAttribute("s", "1"),
                            new XElement("v", Convert.ToDateTime(value).ToOADate()));
                        return element;
                    }
                    else
                    {
                        element = new XElement("c",
                            new XElement("v"));

                        return element;
                    }
                    

                    


                case object _ when propertyInfo.PropertyType == typeof(DateTime?):
                case object _ when propertyInfo.PropertyType == typeof(DateTime?):

                    var dt = (DateTime?) value;

                    if (dt.HasValue && dt.Value.Year >= 100)
                    {
                        element = new XElement("c",
                            new XAttribute("s", "1"),
                            new XElement("v", dt.Value.ToOADate()));

                        return element;
                    }
                    else
                    {
                        element = new XElement("c",
                            new XElement("v"));

                        return element;
                    }

                default:

                    element = new XElement("c",
                        new XElement("v", value));

                    return element;
            }
        }

        private static string ColumnLetter(int intCol)
        {
            var intFirstLetter = intCol / 676 + 64;
            var intSecondLetter = intCol % 676 / 26 + 64;
            var intThirdLetter = intCol % 26 + 65;

            var firstLetter = intFirstLetter > 64
                ? (char) intFirstLetter
                : ' ';
            var secondLetter = intSecondLetter > 64
                ? (char) intSecondLetter
                : ' ';
            var thirdLetter = (char) intThirdLetter;

            return string.Concat(firstLetter, secondLetter,
                thirdLetter).Trim();
        }

        private static Cell CreateTextCell(string header, uint index,
            string text)
        {
            var cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = header + index
            };

            var iString = new InlineString();
            var t = new Text {Text = text};
            iString.AppendChild(t);
            cell.AppendChild(iString);
            return cell;
        }

        private static async Task WriteXmlToPartAsync(OpenXmlPart openXmlPart, XElement element)
        {
            using (var xmlTextWriter = new XmlTextWriter(openXmlPart.GetStream(), Encoding.UTF8))
            {
                xmlTextWriter.Formatting = Formatting.Indented;
                xmlTextWriter.WriteStartDocument();
                element.WriteTo(xmlTextWriter);
                xmlTextWriter.WriteEndDocument();
                xmlTextWriter.Flush();
            }
        }

        public static byte[] GenerateExcel(IEnumerable<string> headers,
            IEnumerable<List<string>> dataRows)
        {
            var stream = new MemoryStream();
            var document = SpreadsheetDocument
                .Create(stream, SpreadsheetDocumentType.Workbook);

            var workbooks = document.AddWorkbookPart();
            workbooks.Workbook = new Workbook();
            var worksheetPart = workbooks.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();

            worksheetPart.Worksheet = new Worksheet(sheetData);

            var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet
            {
                Id = document.WorkbookPart
                    .GetIdOfPart(worksheetPart),
                SheetId = 1, Name = "Sheet 1"
            };
            sheets.AppendChild(sheet);

            uint rowIndex = 0;
            var row = new Row {RowIndex = ++rowIndex};
            sheetData.AppendChild(row);
            var cellIndex = 0;

            foreach (var header in headers)
            {
                row.AppendChild(CreateTextCell(ColumnLetter(cellIndex++),
                    rowIndex, header));
            }


            foreach (var rowData in dataRows)
            {
                cellIndex = 0;
                row = new Row {RowIndex = ++rowIndex};
                sheetData.AppendChild(row);
                foreach (var cell in rowData.Select(callData => CreateTextCell(ColumnLetter(cellIndex++),
                    rowIndex, callData ?? string.Empty)))
                {
                    row.AppendChild(cell);
                }
            }

            workbooks.Workbook.Save();
            document.Close();

            return stream.ToArray();
        }
    }
}