using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

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
                var workbook = doc.AddWorkbookPart;
                var stringTable = workbook.AddNewPart<SharedStringTablePart>();
                var worksheet = workbook.AddNewPart<WorksheetPart>();
                var stylesheet = workbook.AddNewPart<WorkbookStylesPart>();
                var sheetId = workbook.GetIdOfPart(worksheet);
                ;/* Cannot convert LocalDeclarationStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<sst></sst>' at character 4144
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertInitializer(VariableDeclaratorSyntax declarator)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator, Boolean preferExplicitType)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.LocalDeclarationStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                'create the string table
                Dim xmlStringTable = <sst></sst>

 */
                WriteXmlToPart(stringTable, xmlStringTable);
                ;/* Cannot convert LocalDeclarationStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<workbook>\r\n           <b...' at character 4292
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertInitializer(VariableDeclaratorSyntax declarator)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator, Boolean preferExplicitType)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.LocalDeclarationStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                'create the workbook
                Dim xmlWorkbook = <workbook>
                                      <bookViews>
                                          <workbookView/>
                                      </bookViews>
                                      <sheets>
                                          <sheet name="Exported" sheetId="1" r:id=<%= sheetId %>></sheet>
                                      </sheets>
                                  </workbook>

 */
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
                };/* Cannot convert LocalDeclarationStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<worksheet>\r\n            ...' at character 6389
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertInitializer(VariableDeclaratorSyntax declarator)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator, Boolean preferExplicitType)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.LocalDeclarationStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                Dim xmlWorkSheet = <worksheet>
                                       <sheetViews>
                                           <sheetView tabSelected="1" workbookViewId="0">
                                               <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
                                               <selection pane="bottomLeft"/>
                                           </sheetView>
                                       </sheetViews>
                                       <sheetFormatPr defaultRowHeight="15"/>
                                       <cols>
                                           <%= From _p In _pc Select _
                                               <col min="1" max="1" bestFit="1" width="4"/> %>
                                       </cols>
                                       <sheetData>
                                           <row>
                                               <%= From _p In _pc Select _
                                                   <c t="inlineStr">
                                                       <is>
                                                           <t><%= PropertyDictionary(_p.Name) %></t>
                                                       </is>
                                                   </c> %>
                                           </row>
                                           <%= From _e In List Select _
                                               <row>
                                                   <%= From _p In _pc Select _
                                                       BuildCell(_p, _p.GetValue(_e, Nothing)) %>
                                               </row> %>
                                       </sheetData>
                                   </worksheet>

 */
                WriteXmlToPart(worksheet, xmlWorkSheet);
                ;/* Cannot convert LocalDeclarationStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<styleSheet>\r\n          <...' at character 8348
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertInitializer(VariableDeclaratorSyntax declarator)
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator, Boolean preferExplicitType)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.LocalDeclarationStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                Dim xmlStyleSheet = <styleSheet>
                                        <fonts count="1">
                                            <font>
                                                <sz val="11"/>
                                                <color theme="1"/>
                                                <name val="Calibri"/>
                                                <family val="2"/>
                                                <scheme val="minor"/>
                                            </font>
                                        </fonts>
                                        <fills count="2">
                                            <fill><patternFill patternType="none"/>
                                            </fill><fill><patternFill patternType="gray125"/>
                                            </fill></fills>
                                        <borders count="1">
                                            <border>
                                                <left/><right/><top/>
                                                <bottom/><diagonal/>
                                            </border>
                                        </borders>
                                        <cellStyleXfs count="1">
                                            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                                        </cellStyleXfs>
                                        <cellXfs count="2">
                                            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                                            <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
                                        </cellXfs><cellStyles count="1">
                                            <cellStyle name="Normal" xfId="0" builtinId="0"/>
                                        </cellStyles>
                                        <dxfs count="0"/>
                                        <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
                                    </styleSheet>

 */
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
                    ;/* Cannot convert AssignmentStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<c t="inlineStr">\r\n      ...' at character 11333
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.AssignmentStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                _xele = <c t="inlineStr">
                            <is>
                                <t><%= value %></t>
                            </is>
                        </c>

 */
                     // _xele.LastAttribute.Remove()

                    return _xele;
                }

            case object _ when p.PropertyType == typeof(DateTime):
            case object _ when p.PropertyType == typeof(DateTime):
                {
                    ;/* Cannot convert ReturnStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<c s="1">\r\n         <v><%...' at character 11699
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitReturnStatement(ReturnStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.ReturnStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                Return <c s="1">
                           <v><%= CType(value, Date).ToOADate %></v>
                       </c>

 */
                    break;
                }

            case object _ when p.PropertyType == typeof(DateTime?):
            case object _ when p.PropertyType == typeof(DateTime?):
                {
                    DateTime? _dt = (DateTime?)value;

                    if (_dt.HasValue)
                        ;/* Cannot convert ReturnStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<c s="1">\r\n          <v><...' at character 12027
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitReturnStatement(ReturnStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.ReturnStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                    Return <c s="1">
                               <v><%= _dt.Value.ToOADate %></v>
                           </c>

 */
                    else
                        ;/* Cannot convert ReturnStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<c>\r\n          <v></v>\r\...' at character 12187
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitReturnStatement(ReturnStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.ReturnStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                    Return <c>
                               <v></v>
                           </c>

 */
                    break;
                }

            default:
                {
                    ;/* Cannot convert ReturnStatementSyntax, CONVERSION ERROR: Conversion for XmlElement not implemented, please report this issue in '<c>\r\n         <v><%= valu...' at character 12341
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitXmlElement(XmlElementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.XmlElementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitReturnStatement(ReturnStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.ReturnStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

                Return <c>
                           <v><%= value %></v>
                       </c>

 */
                    break;
                }
        }
    }

    private static void WriteXmlToPart(OpenXmlPart part, XElement x)
    {
        System.IO.StreamWriter fs = new System.IO.StreamWriter(part.GetStream, new System.Text.UTF8Encoding());
        System.Xml.XmlTextWriter xmlWriter = new System.Xml.XmlTextWriter(part.GetStream, new UTF8Encoding());
        xmlWriter.Formatting = System.Xml.Formatting.Indented;
        UTF8Encoding enc = new UTF8Encoding();
        xmlWriter.WriteStartDocument();
        x.WriteTo(xmlWriter);
        xmlWriter.WriteEndDocument();
        xmlWriter.Flush();
        xmlWriter.Close();
    }
}
