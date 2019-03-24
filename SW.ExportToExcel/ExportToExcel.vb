Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Reflection
Imports DocumentFormat.OpenXml
'Imports <xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
'Imports <xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
Imports DocumentFormat.OpenXml.Packaging

Public Module ExtensionMethods

    <Extension()>
    Public Function ExportToExcel(Of TEntity)(List As IEnumerable(Of TEntity)) As Byte()

        Dim _propdic As New Dictionary(Of String, String)

        For Each _p In GetType(TEntity).GetProperties()
            Select Case True
                Case _p.PropertyType Is GetType(String), _
                _p.PropertyType Is GetType(DateTime), _
                _p.PropertyType Is GetType(DateTime?), _
                _p.PropertyType Is GetType(Integer), _
                _p.PropertyType Is GetType(Integer?), _
                _p.PropertyType Is GetType(Byte), _
                _p.PropertyType Is GetType(Byte?), _
                _p.PropertyType Is GetType(Short), _
                _p.PropertyType Is GetType(Short?), _
                _p.PropertyType Is GetType(Long), _
                _p.PropertyType Is GetType(Long?), _
                _p.PropertyType Is GetType(Single), _
                _p.PropertyType Is GetType(Single?), _
                _p.PropertyType Is GetType(Double), _
                _p.PropertyType Is GetType(Double?), _
                _p.PropertyType Is GetType(Boolean), _
                _p.PropertyType Is GetType(Boolean?), _
                _p.PropertyType Is GetType(Decimal), _
                _p.PropertyType Is GetType(Decimal?)

                    _propdic(_p.Name) = _p.Name
            End Select
        Next

        Return ExportToExcel(List, _propdic)

    End Function

    <Extension()>
    Public Function ExportToExcel(Of TEntity)(List As IEnumerable(Of TEntity), PropertyList As IEnumerable(Of String)) As Byte()
        Dim _propdic As New Dictionary(Of String, String)

        For Each _p In GetType(TEntity).GetProperties()
            Select Case True
                Case _p.PropertyType Is GetType(String), _
                _p.PropertyType Is GetType(DateTime), _
                _p.PropertyType Is GetType(DateTime?), _
                _p.PropertyType Is GetType(Integer), _
                _p.PropertyType Is GetType(Integer?), _
                _p.PropertyType Is GetType(Byte), _
                _p.PropertyType Is GetType(Byte?), _
                _p.PropertyType Is GetType(Short), _
                _p.PropertyType Is GetType(Short?), _
                _p.PropertyType Is GetType(Long), _
                _p.PropertyType Is GetType(Long?), _
                _p.PropertyType Is GetType(Single), _
                _p.PropertyType Is GetType(Single?), _
                _p.PropertyType Is GetType(Double), _
                _p.PropertyType Is GetType(Double?), _
                _p.PropertyType Is GetType(Boolean), _
                _p.PropertyType Is GetType(Boolean?), _
                _p.PropertyType Is GetType(Decimal), _
                _p.PropertyType Is GetType(Decimal?)

                    If PropertyList.Contains(_p.Name) Then _propdic(_p.Name) = _p.Name
            End Select
        Next

        Return ExportToExcel(List, _propdic)
    End Function

    <Extension()>
    Public Function ExportToExcel(Of TEntity)(List As IEnumerable(Of TEntity), PropertyDictionary As IDictionary(Of String, String)) As Byte()

        Dim _path = System.IO.Path.GetTempFileName
        Try
            Using doc = SpreadsheetDocument.Create(_path, SpreadsheetDocumentType.Workbook)
                Dim workbook = doc.AddWorkbookPart
                Dim stringTable = workbook.AddNewPart(Of SharedStringTablePart)()
                Dim worksheet = workbook.AddNewPart(Of WorksheetPart)()
                Dim stylesheet = workbook.AddNewPart(Of WorkbookStylesPart)()
                Dim sheetId = workbook.GetIdOfPart(worksheet)

                'create the string table
                Dim xmlStringTable = <sst></sst>
                WriteXmlToPart(stringTable, xmlStringTable)

                'create the workbook
                Dim xmlWorkbook = <workbook>
                                      <bookViews>
                                          <workbookView/>
                                      </bookViews>
                                      <sheets>
                                          <sheet name="Exported" sheetId="1" r:id=<%= sheetId %>></sheet>
                                      </sheets>
                                  </workbook>
                WriteXmlToPart(workbook, xmlWorkbook)

                Dim _pc = New List(Of PropertyInfo)
                For Each _p In GetType(TEntity).GetProperties()
                    Select Case True
                        Case _p.PropertyType Is GetType(String), _
                        _p.PropertyType Is GetType(DateTime), _
                        _p.PropertyType Is GetType(DateTime?), _
                        _p.PropertyType Is GetType(Integer), _
                        _p.PropertyType Is GetType(Integer?), _
                        _p.PropertyType Is GetType(Byte), _
                        _p.PropertyType Is GetType(Byte?), _
                        _p.PropertyType Is GetType(Short), _
                        _p.PropertyType Is GetType(Short?), _
                        _p.PropertyType Is GetType(Long), _
                        _p.PropertyType Is GetType(Long?), _
                        _p.PropertyType Is GetType(Single), _
                        _p.PropertyType Is GetType(Single?), _
                        _p.PropertyType Is GetType(Double), _
                        _p.PropertyType Is GetType(Double?), _
                        _p.PropertyType Is GetType(Boolean), _
                        _p.PropertyType Is GetType(Boolean?), _
                        _p.PropertyType Is GetType(Decimal), _
                        _p.PropertyType Is GetType(Decimal?)

                            If PropertyDictionary.ContainsKey(_p.Name) Then
                                _pc.Add(_p)
                            End If
                    End Select
                Next

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

                WriteXmlToPart(worksheet, xmlWorkSheet)

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
                WriteXmlToPart(stylesheet, xmlStyleSheet)
            End Using

            Return System.IO.File.ReadAllBytes(_path)

        Catch ex As Exception
            Throw ex
        Finally
            System.IO.File.Delete(_path)
        End Try
    End Function

    Private Function BuildCell(ByVal p As System.Reflection.PropertyInfo, ByVal value As Object) As XElement

        Dim _xele As XElement

        Select Case True
            Case p.PropertyType Is GetType(String), _
                p.PropertyType Is GetType(Guid), _
                p.PropertyType Is GetType(Guid?), _
                p.PropertyType Is GetType(Boolean), _
                p.PropertyType Is GetType(Boolean?)

                _xele = <c t="inlineStr">
                            <is>
                                <t><%= value %></t>
                            </is>
                        </c>

                '_xele.LastAttribute.Remove()

                Return _xele

            Case p.PropertyType Is GetType(DateTime), p.PropertyType Is GetType(Date)

                Return <c s="1">
                           <v><%= CType(value, Date).ToOADate %></v>
                       </c>

            Case p.PropertyType Is GetType(DateTime?), p.PropertyType Is GetType(Date?)

                Dim _dt As Date? = CType(value, Date?)

                If _dt.HasValue Then

                    Return <c s="1">
                               <v><%= _dt.Value.ToOADate %></v>
                           </c>
                Else

                    Return <c>
                               <v></v>
                           </c>

                End If

            Case Else

                Return <c>
                           <v><%= value %></v>
                       </c>

        End Select

    End Function

    Private Sub WriteXmlToPart(ByVal part As OpenXmlPart, ByVal x As XElement)
        Dim fs As New IO.StreamWriter(part.GetStream, New System.Text.UTF8Encoding)
        Dim xmlWriter As New Xml.XmlTextWriter(part.GetStream, New UTF8Encoding)
        xmlWriter.Formatting = Xml.Formatting.Indented
        Dim enc As New UTF8Encoding
        xmlWriter.WriteStartDocument()
        x.WriteTo(xmlWriter)
        xmlWriter.WriteEndDocument()
        xmlWriter.Flush()
        xmlWriter.Close()
    End Sub
End Module