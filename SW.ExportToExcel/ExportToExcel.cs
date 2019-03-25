using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace SW.ExcelTools
{
    public static class ExtensionMethods
    {
        public static byte[] ExportToExcel<TEntity>(this IEnumerable<TEntity> List)
        {
            Dictionary<string, string> _propdic = new Dictionary<string, string>();

            foreach (var _p in typeof(TEntity).GetProperties())
            {
                switch (true)
                {
                    case object _ when _p.PropertyType == typeof(string):
                    case object _ when _p.PropertyType == typeof(DateTime):
                    case object _ when _p.PropertyType == typeof(DateTime?):
                    case object _ when _p.PropertyType == typeof(int):
                    case object _ when _p.PropertyType == typeof(int?):
                    case object _ when _p.PropertyType == typeof(byte):
                    case object _ when _p.PropertyType == typeof(byte?):
                    case object _ when _p.PropertyType == typeof(short):
                    case object _ when _p.PropertyType == typeof(short?):
                    case object _ when _p.PropertyType == typeof(long):
                    case object _ when _p.PropertyType == typeof(long?):
                    case object _ when _p.PropertyType == typeof(float):
                    case object _ when _p.PropertyType == typeof(float?):
                    case object _ when _p.PropertyType == typeof(double):
                    case object _ when _p.PropertyType == typeof(double?):
                    case object _ when _p.PropertyType == typeof(bool):
                    case object _ when _p.PropertyType == typeof(bool?):
                    case object _ when _p.PropertyType == typeof(decimal):
                    case object _ when _p.PropertyType == typeof(decimal?):
                        {
                            _propdic[_p.Name] = _p.Name;
                            break;
                        }
                }
            }

            return ExportToExcel(List, _propdic);
        }

        public static byte[] ExportToExcel<TEntity>(this IEnumerable<TEntity> List, IEnumerable<string> PropertyList)
        {
            Dictionary<string, string> _propdic = new Dictionary<string, string>();

            foreach (var _p in typeof(TEntity).GetProperties())
            {
                switch (true)
                {
                    case object _ when _p.PropertyType == typeof(string):
                    case object _ when _p.PropertyType == typeof(DateTime):
                    case object _ when _p.PropertyType == typeof(DateTime?):
                    case object _ when _p.PropertyType == typeof(int):
                    case object _ when _p.PropertyType == typeof(int?):
                    case object _ when _p.PropertyType == typeof(byte):
                    case object _ when _p.PropertyType == typeof(byte?):
                    case object _ when _p.PropertyType == typeof(short):
                    case object _ when _p.PropertyType == typeof(short?):
                    case object _ when _p.PropertyType == typeof(long):
                    case object _ when _p.PropertyType == typeof(long?):
                    case object _ when _p.PropertyType == typeof(float):
                    case object _ when _p.PropertyType == typeof(float?):
                    case object _ when _p.PropertyType == typeof(double):
                    case object _ when _p.PropertyType == typeof(double?):
                    case object _ when _p.PropertyType == typeof(bool):
                    case object _ when _p.PropertyType == typeof(bool?):
                    case object _ when _p.PropertyType == typeof(decimal):
                    case object _ when _p.PropertyType == typeof(decimal?):
                        {
                            if (PropertyList.Contains(_p.Name))
                                _propdic[_p.Name] = _p.Name;
                            break;
                        }
                }
            }

            return ExportToExcel(List, _propdic);
        }

        public static byte[] ExportToExcel<TEntity>(this IEnumerable<TEntity> List, IDictionary<string, string> PropertyDictionary)
        {
            var _path = System.IO.Path.GetTempFileName();
            try
            {
                using (var doc = SpreadsheetDocument.Create(_path, SpreadsheetDocumentType.Workbook))
                {
                    var workbook = doc.AddWorkbookPart();
                    var stringTable = workbook.AddNewPart<SharedStringTablePart>();
                    var worksheet = workbook.AddNewPart<WorksheetPart>();
                    var stylesheet = workbook.AddNewPart<WorkbookStylesPart>();
                    var sheetId = workbook.GetIdOfPart(worksheet);

                    XNamespace relations = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    XNamespace mainns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                    // create the string table
                    var xmlStringTable = new XElement(mainns + "sst");

                    WriteXmlToPart(stringTable, xmlStringTable);


                    XElement xmlWorkbook = new XElement(mainns + "workbook", new XAttribute("xmlns", mainns.NamespaceName), new XAttribute(XNamespace.Xmlns + "r", relations.NamespaceName), new XElement("bookViews", new XElement("workbookView")), new XElement("sheets", new XElement("sheet", new XAttribute("name", "Exported"), new XAttribute("sheetId", "1"), new XAttribute(relations + "id", sheetId))));


                    foreach (XElement _r in xmlWorkbook.Descendants())
                    {
                        if (_r.Name.Namespace == "")
                        {
                            _r.Attributes("xmlns").Remove();
                            _r.Name = _r.Parent.Name.Namespace + _r.Name.LocalName;
                        }
                    }

                    WriteXmlToPart(workbook, xmlWorkbook);

                    var _pc = new List<PropertyInfo>();
                    foreach (var _p in typeof(TEntity).GetProperties())
                    {
                        switch (true)
                        {
                            case object _ when _p.PropertyType == typeof(string):
                            case object _ when _p.PropertyType == typeof(DateTime):
                            case object _ when _p.PropertyType == typeof(DateTime?):
                            case object _ when _p.PropertyType == typeof(int):
                            case object _ when _p.PropertyType == typeof(int?):
                            case object _ when _p.PropertyType == typeof(byte):
                            case object _ when _p.PropertyType == typeof(byte?):
                            case object _ when _p.PropertyType == typeof(short):
                            case object _ when _p.PropertyType == typeof(short?):
                            case object _ when _p.PropertyType == typeof(long):
                            case object _ when _p.PropertyType == typeof(long?):
                            case object _ when _p.PropertyType == typeof(float):
                            case object _ when _p.PropertyType == typeof(float?):
                            case object _ when _p.PropertyType == typeof(double):
                            case object _ when _p.PropertyType == typeof(double?):
                            case object _ when _p.PropertyType == typeof(bool):
                            case object _ when _p.PropertyType == typeof(bool?):
                            case object _ when _p.PropertyType == typeof(decimal):
                            case object _ when _p.PropertyType == typeof(decimal?):
                                {
                                    if (PropertyDictionary.ContainsKey(_p.Name))
                                        _pc.Add(_p);
                                    break;
                                }
                        }
                    }

                    XElement xmlWorkSheet = new XElement(mainns + "worksheet", new XElement("sheetViews", new XElement("sheetView", new XAttribute("tabSelected", "1"), new XAttribute("workbookViewId", "0")), new XElement("pane", new XAttribute("ySplit", "1"), new XAttribute("topLeftCell", "A2"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen")), new XElement("selection", new XAttribute("pane", "bottomLeft"))), new XElement("sheetFormatPr", new XAttribute("defaultRowHeight", "15")), new XElement("cols", GetCols(_pc)), new XElement("sheetData", GetRowValues<TEntity>(List, PropertyDictionary, _pc)));

                    foreach (XElement _r in xmlWorkSheet.Descendants())
                    {
                        if (_r.Name.Namespace == "")
                        {
                            _r.Attributes("xmlns").Remove();
                            _r.Name = _r.Parent.Name.Namespace + _r.Name.LocalName;
                        }
                    }

                    WriteXmlToPart(worksheet, xmlWorkSheet);

                    var strStyleSheet = @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
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
                    var xmlStyleSheet = XElement.Parse(strStyleSheet);
                    WriteXmlToPart(stylesheet, xmlStyleSheet);
                }

                return System.IO.File.ReadAllBytes(_path);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.IO.File.Delete(_path);
            }
        }

        private static XElement[] GetRowValues<TEntity>(IEnumerable<TEntity> List, IDictionary<string, string> PropertyDictionary, List<PropertyInfo> Props)
        {
            var result = new List<XElement>();

            var _RowNames = new XElement("row", from _p in Props
                                                select new XElement("c", new XAttribute("t", "inlineStr"), new XElement("is", new XElement("t", PropertyDictionary[_p.Name]))));
            result.Add(_RowNames);

            foreach (var _e in List)
            {
                var _row = new XElement("row", from _p in Props
                                               select BuildCell(_p, _p.GetValue(_e, null)));

                result.Add(_row);
            }
            return result.ToArray();
        }

        private static XElement[] GetCols(List<PropertyInfo> Props)
        {
            var result = new List<XElement>();

            foreach (var _r in Props)

                result.Add(new XElement("col", new XAttribute("min", "1"), new XAttribute("max", "1"), new XAttribute("bestFit", "1"), new XAttribute("width", "4")));

            return result.ToArray();
        }
        private static XElement BuildCell(System.Reflection.PropertyInfo p, object value)
        {
            XElement _xele;

            switch (true)
            {
                case object _ when p.PropertyType == typeof(string):
                case object _ when p.PropertyType == typeof(Guid):
                case object _ when p.PropertyType == typeof(Guid?):
                case object _ when p.PropertyType == typeof(bool):
                case object _ when p.PropertyType == typeof(bool?):
                    {
                        _xele = new XElement("c", new XAttribute("t", "inlineStr"), new XElement("is", new XElement("t", value)));

                        // _xele = <c t="inlineStr">
                        // <is>
                        // <t><%= value %></t>
                        // </is>
                        // </c>

                        // _xele.LastAttribute.Remove()

                        return _xele;
                    }

                case object _ when p.PropertyType == typeof(DateTime):
                case object _ when p.PropertyType == typeof(DateTime):
                    {
                        _xele = new XElement("c", new XAttribute("s", "1"), new XElement("v", System.Convert.ToDateTime(value).ToOADate()));
                        return _xele;
                    }

                case object _ when p.PropertyType == typeof(DateTime?):
                case object _ when p.PropertyType == typeof(DateTime?):
                    {
                        DateTime? _dt = (DateTime?)value;

                        if (_dt.HasValue)
                        {
                            _xele = new XElement("c", new XAttribute("s", "1"), new XElement("v", _dt.Value.ToOADate()));
                            return _xele;
                        }
                        else
                        {
                            _xele = new XElement("c", new XElement("v"));
                            return _xele;
                        }

                        break;
                    }

                default:
                    {
                        _xele = new XElement("c", new XElement("v", value));
                        return _xele;
                    }
            }
        }

        private static void WriteXmlToPart(OpenXmlPart part, XElement x)
        {
            // Dim fs As New IO.StreamWriter(part.GetStream, New System.Text.UTF8Encoding)
            using (System.Xml.XmlTextWriter xmlWriter = new System.Xml.XmlTextWriter(part.GetStream(), new UTF8Encoding()))
            {
                xmlWriter.Formatting = System.Xml.Formatting.Indented;
                UTF8Encoding enc = new UTF8Encoding();
                xmlWriter.WriteStartDocument();
                x.WriteTo(xmlWriter);
                xmlWriter.WriteEndDocument();
                xmlWriter.Flush();
                xmlWriter.Close();
            }
        }
    }
}
