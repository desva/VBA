Attribute VB_Name = "modSerialise"
Option Explicit
' References:
' TypeLib Information
' Microsoft Scripting Runtime

' ToDos:
' Excel8 compatibility check
' Borders still have issues
' Validation
' Conditional formats
' Charts (and non-worksheet sheets generally)
'

' Top level keys
Const WbNameKey As String = "WBNames"
' Sheet level keys
Const PropsKey As String = "Property"
Const FmlaKey As String = "Formula"
Const ConstKey As String = "Constant"
Const CommentKey As String = "Comment"
Const HyperlinkKey As String = "Hyperlink"
Const NameKey As String = "Name"
Const FormatKey As String = "Format"
Const CellSizes As String = "CellSizes"
Const RowColGroups As String = "Outlines"
Const ViewKey As String = "View"

' Formats sub keys
Const FontKey As String = "Fonts"
Const BorderKey As String = "Borders"
Const ColourKey As String = "Colours"
Const NumberFormatKey As String = "NumericFormat"

Public Sub SerializeActiveWorkbook()
  Dim ws As Worksheet
  Dim d As Dictionary
  Dim oPt As New clsPerfTimer
  oPt.Start
  Set d = New Dictionary
  Call d.Add(WbNameKey, WorkbookNames(ActiveWorkbook))
  For Each ws In ActiveWorkbook.Worksheets
    Call d.Add("Worksheet:=" & ws.Name, WorksheetToDictionary(ws))
    'Set d = WorksheetToDictionary(ws)
    'Call SaveJson(d, Environ$("TEMP") & "\" & ws.Name & ".json")
    'Set d = Nothing
  Next
  oPt.Mark
  Debug.Print "Workbook distillation took : " & oPt.AsString
  Call SaveJson(d, Environ$("TEMP") & "\" & ActiveWorkbook.Name & ".json")
  oPt.Mark
  Debug.Print "JSON serialisation took: " & oPt.AsString
  Set oPt = Nothing
End Sub

Public Sub testas()
  Dim ws As Worksheet
  Dim d As Dictionary
  Dim oPt As New clsPerfTimer
  Set ws = ActiveSheet
  Set d = WorksheetToDictionary(ws)
  oPt.Start
  Call SaveJson(d, Environ$("TEMP") & "\" & ActiveSheet.Name & ".json")
  oPt.Mark
  Debug.Print "JSON serialisation took: " & oPt.AsString
  Set d = Nothing
  Set oPt = Nothing
End Sub


'TODO: conditional formats, charts, buttons, validations & pivot tables
Public Function WorksheetToDictionary(ByVal ws As Worksheet) As Dictionary
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim oPt As New clsPerfTimer
  Dim d As Dictionary
  Dim r As Range
  oPt.Start
  ' Worksheet
  Set d = New Dictionary
  Set r = ws.UsedRange
  Set d(PropsKey) = WorkSheetProperties(ws)
  Set d(ViewKey) = WorksheetViews(ws)
  Set d(HyperlinkKey) = WorksheetHyperlinks(ws)
  Set d(NameKey) = WorksheetNames(ws)
  ' Formulas, Constants, Comments etc.
  Set d(FmlaKey) = FormulasToDictionary(r)
  'Set d(ConstKey) = ValuesToDictionary(r)
  'Set d(CommentKey) = CommentsToDictionary(r)
  'Set d(FormatKey) = FormatsToDictionary(r)
  'Set d(CellSizes) = CellSizesToDictionary(r)
  'Set d(RowColGroups) = GroupsToDictionary(r)

  oPt.Mark
  Debug.Print "Total distillation time for " & ws.Name & " : " & oPt.AsString
  Set oPt = Nothing
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Set WorksheetToDictionary = d
  Set d = Nothing
  Set r = Nothing
End Function

Public Sub TestRender()
  JsonToWorksheet "C:\Users\DANIEL~1\AppData\Local\Temp\RayTracer.xls.json"
End Sub

Public Sub JsonToWorksheet(ByVal strFilename As String)
  Dim d1 As Dictionary, d2 As Dictionary, dC As Dictionary
  Dim wb As Workbook, ws As Worksheet, ws0 As Worksheet, wsv As WorksheetView, vKey
  Dim oPt As New clsPerfTimer
  Dim strTempName As String
  Set d1 = LoadJson(strFilename)
  oPt.Start
  ThisWorkbook.UpdateLinks = xlUpdateLinksNever
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
  Set dC = New Dictionary
  Set dC("XlEnableSelection") = XlEnableSelectionConsts()
  Set dC("XlSheetVisibility") = XlSheetVisibilityConsts()
  oPt.Mark
  Set wb = Application.Workbooks.Add(xlWBATWorksheet)
  ' Compatibility? This seems to be the only way to force Excel to create old style workbook...
  strTempName = Environ$("TEMP") & "\" & CreateGuidString & ".xls"
  wb.SaveAs Filename:=strTempName, FileFormat:=xlExcel8
  wb.Close
  Set wb = Workbooks.Open(strTempName)
  Debug.Print "Initialisation took " & oPt.AsString

  ' Phase 1 - set up sheets
  For Each vKey In d1.Keys
    If vKey <> WbNameKey Then
      oPt.Mark
      If ws0 Is Nothing Then
        Set ws = wb.Worksheets.Add
      Else
        Set ws = wb.Worksheets.Add(After:=ws0)
      End If
      Set wsv = GetWorksheetView(ws)
      ws.EnableCalculation = False
      Set d2 = d1(vKey)
      Call LetProperties(wsv, d2(ViewKey), XlWorkSheetViewProperties(), Nothing)
      Call LetProperties(ws, d2(PropsKey), XlWorkSheetProperties(), dC)
      Set ws0 = ws
    End If
  Next
  Call RenderNames(d1(WbNameKey), wb)
  oPt.Mark
  Debug.Print "Sheet set up took " & oPt.AsString
  
  For Each vKey In d1.Keys
    If vKey <> WbNameKey Then
      oPt.Mark
      Set d2 = d1(vKey)
      Set ws = wb.Worksheets(Mid(vKey, InStr(vKey, "=") + 1))
      Call RenderNames(d2(NameKey), ws)
      Call RenderFormulas(d2(FmlaKey), ws)
      ' TODO
      If d2.Exists(ConstKey) Then
        Call RenderValues(d2(ConstKey), ws) ' Need to check literals!
      End If
      Call RenderComments(d2(CommentKey), ws)
      Call RenderFonts(d2(FormatKey)(FontKey), ws)
      Call RenderNumerics(d2(FormatKey)(NumberFormatKey), ws)
      Call RenderColours(d2(FormatKey)(ColourKey), ws)
      Call RenderBorders(d2(FormatKey)(BorderKey), ws)
      Call RenderCellSizes(d2(CellSizes), ws)
      Call RenderOutlines(d2(RowColGroups), ws)
      oPt.Mark
      Debug.Print "Rendering " & vKey & " took " & oPt.AsString
    End If
  Next
  
  Application.ScreenUpdating = True
  oPt.Mark
  Debug.Print "Rendering time: " & oPt.AsString
  Set oPt = Nothing

End Sub

Private Sub SaveJson(ByVal d As Dictionary, fullFilename As String)
  Dim fso As FileSystemObject, ts As TextStream
  Set fso = New FileSystemObject
  Set ts = fso.CreateTextFile(fullFilename, True, True) ' Unicode! Some text is unicode
  Dim s As String
  s = JsonConverter.ConvertToJson(d, 2)
  Call ts.Write(s)
  Call ts.Close
  Set fso = Nothing
End Sub

Private Function LoadJson(ByVal fName As String) As Dictionary
  Dim fso As FileSystemObject, ts As TextStream
  Dim d As Dictionary, sIn As String
  Dim oPt As New clsPerfTimer
  oPt.Start
  Set fso = New FileSystemObject
  Set ts = fso.OpenTextFile(fName, ForReading, False, TristateTrue)
  sIn = ts.ReadAll
  Set d = JsonConverter.ParseJson(sIn)
  Call ts.Close
  Set LoadJson = d
  oPt.Mark
  Debug.Print "Json load took : " & oPt.AsString
  Set fso = Nothing
  Set d = Nothing
End Function

Public Function WorkSheetProperties(ByVal ws As Worksheet) As Dictionary
  Dim dTypes As Dictionary, dConsts As Dictionary, dConsts2 As Dictionary, dProps As Dictionary
  Dim dOut As Dictionary
  Dim vk As Variant, vv As Variant
  Set dOut = New Dictionary
  Set dTypes = xlBaseTypes
  Set dConsts = XlSheetVisibilityConsts(True)
  Set dConsts2 = XlEnableSelectionConsts(True)
  Set dProps = XlWorkSheetProperties
  For Each vk In dProps.Keys
    vv = CallByName(ws, vk, VbGet)
    If Not dTypes.Exists(dProps(vk)) Then
      If dProps(vk) = "XlSheetVisibility" Then
        vv = CStr(dConsts(vv))
      Else
        vv = CStr(dConsts2(vv))
      End If
    End If
    If Not IsEmpty(vv) Then Call dOut.Add(CStr(vk), vv)
  Next
  Set WorkSheetProperties = dOut
  Set dOut = Nothing
  Set dTypes = Nothing
  Set dConsts = Nothing
  Set dProps = Nothing
End Function

' As properties, but with odd navigation to the view
Public Function WorksheetViews(ByVal ws As Worksheet) As Dictionary
  Dim dProps As Dictionary
  Dim vk As Variant, vv As Variant
  Dim view As WorksheetView
  Set WorksheetViews = New Dictionary
  Set view = GetWorksheetView(ws)
  Set dProps = modXlConstants.XlWorkSheetViewProperties() ' All Boolean
  For Each vk In dProps.Keys
    vv = CallByName(view, vk, VbGet)
    If Not IsEmpty(vv) Then Call WorksheetViews.Add(CStr(vk), vv)
  Next
  Set dProps = Nothing
  Set view = Nothing
End Function

' Introduced in Excel 2007 - worksheet view gives view properties of a sheet, but difficult to navigate to directly
Private Function GetWorksheetView(ByVal ws As Worksheet) As WorksheetView
  Dim view As WorksheetView
  For Each view In ws.Parent.Windows(1).SheetViews
      If view.Sheet.Name = ws.Name Then
          Set GetWorksheetView = view
          Exit Function
      End If
  Next
End Function

Public Function WorksheetNames(ByVal ws As Worksheet) As Variant
  Set WorksheetNames = WorksheetOrWorkbookNames(ws)
End Function

Public Function WorkbookNames(ByVal wb As Workbook) As Variant
  Set WorkbookNames = WorksheetOrWorkbookNames(wb)
End Function

Private Function CoalesceRangeDictionary(ByVal d As Dictionary, ByVal vt As VbVarType) As Dictionary
  Dim oSpans As clsSpans, strKey As String, v
  ' Double or string
  Debug.Assert (vt = vbDouble Or vt = vbString)
  Set CoalesceRangeDictionary = New Dictionary
  For Each v In d.Keys
    Set oSpans = d(v)
    Call oSpans.Coalesce
    strKey = CStr(oSpans.GetRange(xlA1)) '  or xlR1C1
    Set oSpans = Nothing
    Set d(v) = Nothing
    If (vt = vbDouble) Then
      Call CoalesceRangeDictionary.Add(strKey, CDbl(v))
    Else
      Call CoalesceRangeDictionary.Add(strKey, CStr(v))
    End If
  Next
End Function

Private Function FormulasToDictionary(ByVal r As Range) As Dictionary
  Dim rngFmla As Range, rngArea As Range, rngCell As Range
  Dim dctThis As Dictionary
  Dim oSpans As clsSpans
  Dim strKey As String, strRange As String
  Dim v As Variant
  Dim i0 As Long, j0 As Long, i As Long, j As Long
  Dim oPt As New clsPerfTimer
  oPt.Start
  Set dctThis = New Dictionary
  On Error Resume Next
  Set rngFmla = r.Cells.SpecialCells(xlCellTypeFormulas)
  If Err.Number = 0 Then
    On Error GoTo 0
    For Each rngArea In rngFmla.Areas
      If rngArea.HasArray = True Or IsNull(rngArea.HasArray) Or rngArea.Cells.Count = 1 Then
        For Each rngCell In rngArea
          strKey = rngCell.FormulaR1C1
          If rngCell.HasArray Then
            strKey = CStr("@" & rngCell.CurrentArray.Address(False, False, xlA1) & strKey) ' avoid merging adjacent but offset identical array formula!
          End If
          If Not dctThis.Exists(strKey) Then
              Set dctThis(strKey) = New clsSpans
          End If
          Set oSpans = dctThis(strKey)
          oSpans.AddCell rngCell.Row, rngCell.Column
          Set oSpans = Nothing
          strKey = vbNullString
        Next
      Else
        ' vectorise
        v = rngArea.FormulaR1C1
        Set rngCell = rngArea.Cells(1)
        i0 = rngCell.Row() - 1
        j0 = rngCell.Column() - 1
        For i = LBound(v, 1) To UBound(v, 1)
          For j = LBound(v, 2) To UBound(v, 2)
              strKey = CStr(v(i, j))
              If Not dctThis.Exists(strKey) Then
                Set dctThis(strKey) = New clsSpans
              End If
              Set oSpans = dctThis(strKey)
              oSpans.AddCell i0 + i, j0 + j
              Set oSpans = Nothing
          Next
        Next
      End If
    Next
    Set dctThis = CoalesceRangeDictionary(dctThis, vbString)
  End If
  Set FormulasToDictionary = dctThis
  Set dctThis = Nothing
  oPt.Mark
  Debug.Print "Formulas took : " & oPt.AsString
  Set oPt = Nothing
End Function

Public Function ValuesToDictionary(ByVal r As Range) As Dictionary
  Dim rngArea As Range
  Dim dctThis As Dictionary
  Dim oSpans As clsSpans
  Dim oPt As New clsPerfTimer
  oPt.Start
  Set dctThis = New Dictionary
  On Error Resume Next
  Set r = r.Cells.SpecialCells(xlCellTypeConstants)
  For Each rngArea In r.Areas
    dctThis(rngArea.Address(False, False, xlA1)) = rngArea.Value2
  Next
  Set ValuesToDictionary = dctThis
  Set dctThis = Nothing
  oPt.Mark
  Debug.Print "Values took : " & oPt.AsString
  Exit Function
End Function

Public Function CommentsToDictionary(ByVal r As Range) As Dictionary
  Dim rngCmnts As Range, rngCell As Range
  Dim dctThis As Dictionary
  Dim oSpans As clsSpans
  Dim strKey As String
  Set dctThis = New Dictionary
  On Error Resume Next
  Set rngCmnts = r.Cells.SpecialCells(xlCellTypeComments)
  If Err.Number = 0 Then
    On Error GoTo 0
    For Each rngCell In rngCmnts
      strKey = rngCell.Comment.Text
      If Not dctThis.Exists(strKey) Then
          Set dctThis(strKey) = New clsSpans
      End If
      Set oSpans = dctThis(strKey)
      Call oSpans.AddCell(rngCell.Row, rngCell.Column)
    Next
    Set dctThis = CoalesceRangeDictionary(dctThis, vbString)
  End If
  Set CommentsToDictionary = dctThis
  Set dctThis = Nothing
End Function

Private Function WorksheetOrWorkbookNames(ByVal o As Object) As Dictionary
  Dim dTypes As Dictionary, dConsts As Dictionary, dProps As Dictionary, d1 As Dictionary, d2 As Dictionary
  Dim vk, vv
  Set dTypes = xlBaseTypes()
  Set dConsts = XlXLMMacroTypeConsts(True)
  Set dProps = XlNameProperties()
  Set d1 = New Dictionary
  If o.Names.Count <> 0 Then
    Dim nm As Name
    On Error Resume Next
    For Each nm In o.Names
      Set d2 = New Dictionary
      For Each vk In dProps.Keys
        vv = vbNullString
        vv = CallByName(nm, vk, VbGet)
        If dProps(vk) = "XlXLMMacroType" Then
            vv = dConsts(vv)
        End If
        d2(vk) = vv
      Next
      Set d1(nm.Name) = d2
      Set d2 = Nothing
    Next
  End If
  Set dTypes = Nothing
  Set dConsts = Nothing
  Set dProps = Nothing
  Set WorksheetOrWorkbookNames = d1
  Set d1 = Nothing
End Function

Public Function WorksheetHyperlinks(ByVal ws As Worksheet) As Dictionary
  Dim d1 As Dictionary, d2 As Dictionary, dProps As Dictionary
  Set d1 = New Dictionary
  Set dProps = XlHyperlinkProperties()
  Dim o As Hyperlink
  Dim vv As Variant, vk As Variant
  Dim strVal As String
  For Each o In ws.Hyperlinks
    Set d2 = New Dictionary
    For Each vk In dProps.Keys
      vv = vbNullString
      vv = CallByName(o, vk, VbGet)
      d2(vk) = vv
    Next
    Set d1(o.Name) = d2
    Set d2 = Nothing
  Next
  Set WorksheetHyperlinks = d1
  Set dProps = Nothing
  Set d1 = Nothing
End Function

' return true if range has contiguous ....
Private Function ContiguousColour(ByVal r As Range) As Boolean
  ContiguousColour = Not IsNull(r.Interior.Pattern) And Not IsNull(r.Interior.PatternColorIndex) And Not IsNull(r.Interior.ColorIndex)
End Function

Private Function ContiguousFont(ByVal r As Range) As Boolean
  ContiguousFont = Not IsNull(r.Font.Name) And Not IsNull(r.Font.Size) And Not IsNull(r.Font.Bold) And Not IsNull(r.Font.Italic) And Not IsNull(r.Font.Underline) And Not IsNull(r.Font.ColorIndex)
End Function

Private Function ContiguousNumericFormat(ByVal r As Range) As Boolean
  ContiguousNumericFormat = Not IsNull(r.NumberFormat) And Not IsNull(r.VerticalAlignment) _
    And Not IsNull(r.HorizontalAlignment) And Not IsNull(r.Orientation) And Not IsNull(r.WrapText) _
    And Not IsNull(r.ShrinkToFit)
End Function

Private Function ContiguousBorderFormat(ByVal r As Range) As Boolean
  ContiguousBorderFormat = Not IsNull(r.Borders.LineStyle) And Not IsNull(r.Borders.Weight) And Not IsNull(r.Borders.ColorIndex)
End Function

Public Function FormatsToDictionary(ByVal r As Range) As Dictionary
  Dim dctFormats As Dictionary
  Dim oPt As New clsPerfTimer
  oPt.Start
  'On Error Resume Next
  'Set rng = r
  'If Err.Number <> 0 Then Exit Function
  Set dctFormats = New Dictionary
  Set dctFormats(FontKey) = FontsToDictionary(r)
  Set dctFormats(NumberFormatKey) = NumericsToDictionary(r)
  Set dctFormats(ColourKey) = ColoursToDictionary(r)
  Set dctFormats(BorderKey) = BordersToDictionary2(r)
  oPt.Mark
  Set FormatsToDictionary = dctFormats
  Set dctFormats = Nothing
  Debug.Print "Formats took : " & oPt.AsString
  Set oPt = Nothing
End Function

Private Function FontsToDictionary(ByVal rng As Range) As Dictionary
  Dim i As Long, j As Long, nRows As Long
  Dim rRow As Range, rCell As Range
  Dim v As Variant
  Dim oSpans As clsSpans
  Dim strKey As String
  nRows = rng.Rows.Count
  Dim d As Dictionary
  Dim dctUl As Dictionary, dctCI As Dictionary
  Set dctUl = xlUnderlineStyleConsts(True)
  Set dctCI = xlColorIndexConsts(True)
  Set d = New Dictionary
  ' Fonts
  If ContiguousFont(rng) Then
    strKey = CStr(rng.Font.Name & "|" & _
      rng.Font.Size & "|" & _
      rng.Font.Bold & "|" & _
      rng.Font.Italic & "|" & _
      dctUl(rng.Font.Underline) & "|")
    If (rng.Font.ColorIndex < 0) Then
      strKey = strKey & CStr(dctCI(rng.Font.ColorIndex))
    Else
      strKey = strKey & CStr(rng.Font.ColorIndex)
    End If
    Call d.Add(rng.Address(False, False, xlA1), strKey)
    strKey = vbNullString
  Else
    For i = 1 To nRows
      Set rRow = rng.Rows(i)
      If ContiguousFont(rRow) Then
        With rRow.Font
          strKey = CStr(.Name & "|" & _
                    .Size & "|" & _
                    .Bold & "|" & _
                    .Italic & "|" & _
                    dctUl(.Underline) & "|")
          If (.ColorIndex < 0) Then
            strKey = strKey & CStr(dctCI(.ColorIndex))
          Else
            strKey = strKey & CStr(.ColorIndex)
          End If
        End With
        If Not d.Exists(strKey) Then
            Set d(strKey) = New clsSpans
        End If
        Set oSpans = d(strKey)
        For j = rRow(1).Column To rRow(1).Column + rRow.Columns.Count
          Call oSpans.AddCell(rRow.Row, j)
        Next
        Set oSpans = Nothing
        strKey = vbNullString
      Else
        ' one at a time
        For Each rCell In rRow.Cells
          With rCell.Font
            strKey = CStr(.Name & "|" & _
                .Size & "|" & _
                .Bold & "|" & _
                .Italic & "|" & _
                dctUl(.Underline) & "|")
            If (.ColorIndex < 0) Then
              strKey = strKey & CStr(dctCI(.ColorIndex))
            Else
              strKey = strKey & CStr(.ColorIndex)
            End If
          End With
          If Not d.Exists(strKey) Then
              Set d(strKey) = New clsSpans
          End If
          Set oSpans = d(strKey)
          Call oSpans.AddCell(rCell.Row, rCell.Column)
          Set oSpans = Nothing
          strKey = vbNullString
        Next
      End If
    Next
    Set d = CoalesceRangeDictionary(d, vbString)
  End If
  Set FontsToDictionary = d
  Set d = Nothing
  Set dctUl = Nothing
  Set dctCI = Nothing
End Function

Private Function NumericsToDictionary(ByVal rng As Range) As Dictionary
  Dim i As Long, j As Long, nRows As Long
  Dim rRow As Range, rCell As Range
  Dim v As Variant
  Dim oSpans As clsSpans
  Dim strKey As String
  Dim d As Dictionary, dctOrient As Dictionary, dctHAlign As Dictionary, dctVAlign As Dictionary
  Set d = New Dictionary
  Set dctOrient = xlOrientationConsts(True)
  Set dctHAlign = xlHAlignConsts(True)
  Set dctVAlign = xlVAlignConsts(True)
  nRows = rng.Rows.Count
  If ContiguousNumericFormat(rng) Then
    strKey = rng.NumberFormat & "|" & _
      dctVAlign(rng.VerticalAlignment) & "|" & _
      dctHAlign(rng.HorizontalAlignment) & "|"
    If (rng.Orientation < -90) Then
      strKey = strKey & dctOrient(rng.Orientation)
    Else
      strKey = strKey & rng.Orientation
    End If
    strKey = strKey & "|" & _
      rng.WrapText & "|" & _
      rng.ShrinkToFit
    Call d.Add(rng.Address(False, False, xlA1), strKey)
  Else
    For i = 1 To nRows
      Set rRow = rng.Rows(i)
      If ContiguousNumericFormat(rRow) Then
        With rRow
          strKey = .NumberFormat & "|" & _
            dctVAlign(.VerticalAlignment) & "|" & _
            dctHAlign(.HorizontalAlignment) & "|"
          If (.Orientation < -90) Then
            strKey = strKey & dctOrient(.Orientation)
          Else
            strKey = strKey & .Orientation
          End If
          strKey = strKey & "|" & _
            .WrapText & "|" & _
            .ShrinkToFit
        End With
        If Not d.Exists(strKey) Then
            Set d(strKey) = New clsSpans
        End If
        Set oSpans = d(strKey)
        For j = rRow(1).Column To rRow(1).Column + rRow.Columns.Count
          Call oSpans.AddCell(rRow.Row, j)
        Next
        Set oSpans = Nothing
      Else
        ' one at a time
        For Each rCell In rRow.Cells
          With rCell
            strKey = .NumberFormat & "|" & _
              dctVAlign(.VerticalAlignment) & "|" & _
              dctHAlign(.HorizontalAlignment) & "|"
            If (.Orientation < -90) Then
              strKey = strKey & dctOrient(.Orientation)
            Else
              strKey = strKey & .Orientation
            End If
            strKey = strKey & "|" & _
              .WrapText & "|" & _
              .ShrinkToFit
          End With
          If Not d.Exists(strKey) Then
              Set d(strKey) = New clsSpans
          End If
          Set oSpans = d(strKey)
          Call oSpans.AddCell(rCell.Row, rCell.Column)
          Set oSpans = Nothing
        Next
      End If
    Next
    Set d = CoalesceRangeDictionary(d, vbString)
  End If
  Set NumericsToDictionary = d
  Set d = Nothing
End Function

Private Function ColoursToDictionary(ByVal rng As Range) As Dictionary
  Dim i As Long, j As Long, nRows As Long
  Dim rRow As Range, rCell As Range
  Dim v As Variant
  Dim oSpans As clsSpans
  Dim strKey As String
  Dim d As Dictionary, dctPat As Dictionary, dctCI As Dictionary
  Set d = New Dictionary
  Set dctPat = xlPatternConsts(True)
  Set dctCI = xlColorIndexConsts(True)
  nRows = rng.Rows.Count
  If ContiguousColour(rng) Then
    strKey = dctPat(rng.Interior.Pattern) & "|"
    If (rng.Interior.PatternColorIndex < 0) Then
      strKey = strKey & dctCI(rng.Interior.PatternColorIndex) & "|"
    Else
      strKey = strKey & rng.Interior.PatternColorIndex & "|"
    End If
    If (rng.Interior.ColorIndex < 0) Then
      strKey = strKey & dctCI(rng.Interior.ColorIndex)
    Else
      strKey = strKey & rng.Interior.ColorIndex
    End If
    Call d.Add(rng.Address(False, False, xlA1), strKey)
  Else
    For i = 1 To nRows
      Set rRow = rng.Rows(i)
      If ContiguousColour(rRow) Then
        With rRow.Interior
          strKey = dctPat(.Pattern) & "|"
          If (.PatternColorIndex < 0) Then
            strKey = strKey & dctCI(.PatternColorIndex) & "|"
          Else
            strKey = strKey & .PatternColorIndex & "|"
          End If
          If (.ColorIndex < 0) Then
            strKey = strKey & dctCI(.ColorIndex)
          Else
            strKey = strKey & .ColorIndex
          End If
        End With
        If Not d.Exists(strKey) Then
            Set d(strKey) = New clsSpans
        End If
        Set oSpans = d(strKey)
        For j = rRow(1).Column To rRow(1).Column + rRow.Columns.Count
          Call oSpans.AddCell(rRow.Row, j)
        Next
        Set oSpans = Nothing
      Else
        ' one at a time
        For Each rCell In rRow.Cells
          With rCell.Interior
            strKey = dctPat(.Pattern) & "|"
            If (.PatternColorIndex < 0) Then
              strKey = strKey & dctCI(.PatternColorIndex) & "|"
            Else
              strKey = strKey & .PatternColorIndex & "|"
            End If
            If (.ColorIndex < 0) Then
              strKey = strKey & dctCI(.ColorIndex)
            Else
              strKey = strKey & .ColorIndex
            End If
            If Not d.Exists(strKey) Then
                Set d(strKey) = New clsSpans
            End If
            Set oSpans = d(strKey)
            Call oSpans.AddCell(rCell.Row, rCell.Column)
            Set oSpans = Nothing
          End With
        Next
      End If
    Next
    Set d = CoalesceRangeDictionary(d, vbString)
  End If
  Set ColoursToDictionary = d
  Set d = Nothing
End Function

Private Function BordersToDictionary(ByVal rng As Range) As Dictionary
  Dim i As Long, j As Long, nRows As Long
  Dim rRow As Range, rCell As Range
  Dim v As Variant
  Dim oSpans As clsSpans
  Dim strKey As String
  Dim b As Border
  Dim d As Dictionary, dctLStyle As Dictionary, dctCI As Dictionary, dctBord As Dictionary
  Set d = New Dictionary
  Set dctLStyle = xlLineStyleConsts(True)
  Set dctCI = xlColorIndexConsts(True)
  Set dctBord = xlBordersIndexConsts(True)
    
  ' Expand range - rather than examine right and bottom
  On Error Resume Next
'  Set rng = rng.Resize(rng.Rows.Count + 1, rng.Columns.Count + 1)
  On Error GoTo 0
  nRows = rng.Rows.Count
  
  If ContiguousBorderFormat(rng) Then
    For Each v In dctBord.Keys
        Set b = rng.Borders(v)
        If b.LineStyle <> xlLineStyleNone Then
          strKey = dctBord(v) & "|" & _
            dctLStyle(b.LineStyle) & "|" & _
            b.Weight & "|"
          If (b.ColorIndex < 0) Then
            strKey = strKey & dctCI(b.ColorIndex)
          Else
            strKey = strKey & b.ColorIndex
          End If
          Call d.Add(rng.Address(False, False, xlA1), strKey)
        End If
    Next
  Else
    For i = 1 To nRows
      Set rRow = rng.Rows(i)
      If False * ContiguousBorderFormat(rRow) Then
        For Each v In dctBord.Keys
          Set b = rRow.Borders(v)
          If b.LineStyle <> xlLineStyleNone Then ' check null?
            strKey = dctBord(v) & "|" & _
            dctLStyle(b.LineStyle) & "|" & _
            b.Weight & "|"
            If (b.ColorIndex < 0) Then
              strKey = strKey & dctCI(b.ColorIndex)
            Else
              strKey = strKey & b.ColorIndex
            End If
            If Not d.Exists(strKey) Then
              Set d(strKey) = New clsSpans
            End If
            Set oSpans = d(strKey)
            For j = rRow(1).Column To rRow(1).Column + rRow.Columns.Count
              Call oSpans.AddCell(rRow.Row, j)
            Next
            Set oSpans = Nothing
          End If
        Next
      Else
        ' one at a time
        For Each rCell In rRow.Cells
          For Each v In dctBord.Keys
            Set b = rCell.Borders(v)
            If b.LineStyle <> xlLineStyleNone Then
              'rCell.Select
              strKey = dctBord(v) & "|" & _
                dctLStyle(b.LineStyle) & "|" & _
                b.Weight & "|"
              If (b.ColorIndex < 0) Then
                strKey = strKey & dctCI(b.ColorIndex)
              Else
                strKey = strKey & b.ColorIndex
              End If
              If Not d.Exists(strKey) Then
                Set d(strKey) = New clsSpans
              End If
              Set oSpans = d(strKey)
              Call oSpans.AddCell(rCell.Row, rCell.Column)
              Set oSpans = Nothing
            End If
          Next
        Next
      End If
    Next
    Set d = CoalesceRangeDictionary(d, vbString)
  End If
  Set rng = Nothing
  Set BordersToDictionary = d
  Set d = Nothing
End Function

Private Function BordersToDictionary2(ByVal r As Range) As Dictionary
  Dim i As Long, j As Long, k As Long
  Dim rRow As Range, rCell As Range
  Dim v As Variant
  Dim oSpans As clsSpans
  Dim strKey As String
  Dim b As Border, bc As Borders
  Dim d As Dictionary, dctLStyle As Dictionary, dctCI As Dictionary, dctBord As Dictionary
  Set d = New Dictionary
  Dim arrBordInd As Variant
  Set dctLStyle = xlLineStyleConsts(True)
  Set dctCI = xlColorIndexConsts(True)
  Set dctBord = xlBordersIndexConsts(True)
  ' Borders collection returned in this order
  arrBordInd = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlDiagonalDown, xlDiagonalUp)
  
  Set bc = r.Borders
  k = LBound(arrBordInd)
  For Each v In bc
    If v.LineStyle = xlLineStyleNone Then
      Call dctBord.Remove(arrBordInd(k))
    End If
    k = k + 1
  Next
  ' Scan Rows - should only do horizontal styles here really, then do columns, and finally do remaining cells, but establishing "remaining" is hard
  For j = 1 To r.Rows.Count
    Set rRow = r.Rows(j)
    Set bc = rRow.Borders
    k = LBound(arrBordInd)
    For Each v In bc
      If dctBord.Exists(arrBordInd(k)) Then
        Set b = v
        If IsNull(b.LineStyle) Or IsNull(b.Weight) Or IsNull(b.ColorIndex) Then
          ' multiple formats, do cell at a time
          For Each rCell In rRow.Cells
            Set b = rCell.Borders(k + 1)
            If b.LineStyle <> xlLineStyleNone Then
              strKey = dctBord(arrBordInd(k)) & "|" & dctLStyle(b.LineStyle) & "|" & b.Weight & "|"
              If (b.ColorIndex < 0) Then
                strKey = strKey & dctCI(b.ColorIndex)
              Else
                strKey = strKey & b.ColorIndex
              End If
              If Not d.Exists(strKey) Then
                Set d(strKey) = New clsSpans
              End If
              Set oSpans = d(strKey)
              Call oSpans.AddCell(rCell.Row, rCell.Column)
              Set oSpans = Nothing
            End If
          Next
        ElseIf b.LineStyle <> xlLineStyleNone Then
          ' Single format for row
          strKey = dctBord(arrBordInd(k)) & "|" & dctLStyle(b.LineStyle) & "|" & b.Weight & "|"
          If (b.ColorIndex < 0) Then
            strKey = strKey & dctCI(b.ColorIndex)
          Else
            strKey = strKey & b.ColorIndex
          End If
          If Not d.Exists(strKey) Then
            Set d(strKey) = New clsSpans
          End If
          Set oSpans = d(strKey)
          For i = rRow(1).Column To rRow(1).Column + rRow.Columns.Count
            Call oSpans.AddCell(rRow.Row, i)
          Next
          Set oSpans = Nothing
        End If
      End If
      k = k + 1
    Next
  Next
  Set BordersToDictionary2 = CoalesceRangeDictionary(d, vbString)
End Function

Public Function CellSizesToDictionary(ByVal r As Range) As Dictionary
  Dim oSpans As clsSpans
  Dim sw As Double, sh As Double
  Dim d As Dictionary, d2 As Dictionary
  Dim rCell As Range
  Dim v As Variant
  Set d = New Dictionary
  Dim strKey As String
  ' Assumes it's a worksheet...
  sw = r.Parent.StandardWidth
  sh = r.Parent.StandardHeight
  ' Columns:
  If Not IsNull(r.ColumnWidth) Then
    If r.ColumnWidth <> sw Then
      d("w") = r.ColumnWidth
    End If ' Else it's the standard width and can ignore
  Else
    Set d2 = New Dictionary
    For Each rCell In r.Columns
      strKey = rCell.ColumnWidth
      If Not d2.Exists(strKey) Then
        Set d2(strKey) = New clsSpans
      End If
      Set oSpans = d2(strKey)
      Call oSpans.AddCell(rCell.Row, rCell.Column)
      Set oSpans = Nothing
    Next
    Set d("w") = CoalesceRangeDictionary(d2, vbDouble)
    Set d2 = Nothing
  End If
  ' Rows:
  If Not IsNull(r.RowHeight) Then
    If r.RowHeight <> sh Then
      d("h") = r.RowHeight
    End If ' Else it's the standard width and can ignore
  Else
    Set d2 = New Dictionary
    For Each rCell In r.Rows
      strKey = rCell.RowHeight
      If Not d2.Exists(strKey) Then
        Set d2(strKey) = New clsSpans
      End If
      Set oSpans = d2(strKey)
      Call oSpans.AddCell(rCell.Row, rCell.Column)
      Set oSpans = Nothing
    Next
    Set d("h") = CoalesceRangeDictionary(d2, vbDouble)
    Set d2 = Nothing
  End If
  Set CellSizesToDictionary = d
  Set d = Nothing
End Function

Public Function GroupsToDictionary(ByVal r As Range) As Dictionary
  Dim oSpans As clsSpans
  Dim d As Dictionary, d2 As Dictionary
  Dim rCell As Range
  Dim v As Variant
  Set d = New Dictionary
  Dim strKey As String
  ' Column grouping
  If Not IsNull(r.Columns.OutlineLevel) Then
    d("c") = r.Columns.OutlineLevel
  Else
    Set d2 = New Dictionary
    For Each rCell In r.Columns
      strKey = rCell.OutlineLevel
      If Not d2.Exists(strKey) Then
        Set d2(strKey) = New clsSpans
      End If
      Set oSpans = d2(strKey)
      oSpans.AddCell rCell.Row, rCell.Column
      Set oSpans = Nothing
    Next
    Set d("c") = CoalesceRangeDictionary(d2, vbDouble)
    Set d2 = Nothing
  End If
  ' Rows:
  If Not IsNull(r.Rows.OutlineLevel) Then
    d("r") = r.Rows.OutlineLevel
  Else
    Set d2 = New Dictionary
    For Each rCell In r.Rows
      strKey = rCell.OutlineLevel
      If Not d2.Exists(strKey) Then
        Set d2(strKey) = New clsSpans
      End If
      Set oSpans = d2(strKey)
      oSpans.AddCell rCell.Row, rCell.Column
      Set oSpans = Nothing
    Next
    Set d("r") = CoalesceRangeDictionary(d2, vbDouble)
    Set d2 = Nothing
  End If
  Set GroupsToDictionary = d
  Set d = Nothing
End Function

''''''''''''''''''''
'' Rendering subs ''
''''''''''''''''''''
Private Sub RenderNames(ByVal d As Dictionary, ByVal o As Object)
  Dim k, dP As Dictionary, dC As Dictionary, dI As Dictionary
  Set dP = XlNameProperties()
  Set dC = New Dictionary
  Set dC = XlXLMMacroTypeConsts
  For Each k In d.Keys
    Set dI = d(k)
    '([Name], [RefersTo], [Visible], [MacroType], [ShortcutKey], [Category], [NameLocal], [RefersToLocal], [CategoryLocal], [RefersToR1C1], [RefersToR1C1Local])
    'On Error Resume Next
    Call o.Names.Add(Name:=dI("Name"), RefersTo:=dI("RefersTo"), Visible:=dI("Visible"), MacroType:=dC(dI("MacroType")), ShortcutKey:=dI("ShortCutKey"))
  Next
End Sub
Private Sub RenderFormulas(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k, v, v2, r As Range
  For Each k In d.Keys
    v = d(k)
    For Each v2 In Split(k, ",")
      Set r = ws.Range(v2)
      If (Left(v, 1) = "@") Then
        r.FormulaArray = Mid(v, InStr(1, v, "="))
      Else
        r.FormulaR1C1 = v
      End If
    Next
  Next
End Sub
Private Sub RenderValues(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k, r As Range, c As Collection, i As Long, j As Long, ki As Long, kj As Long, v, v2, v3
  For Each k In d.Keys
    If VarType(d(k)) = vbObject Then
      Set c = d(k)
      i = c.Count
      If VarType(c(i)) = vbObject Then
        ' two dimensional
        Set v2 = c(i)
        j = v2.Count
        ReDim v(1 To i, 1 To j)
        ki = 1
        For Each v2 In c
          kj = 1
          For Each v3 In v2
            v(ki, kj) = v3
            kj = kj + 1
          Next
          ki = ki + 1
        Next
      Else
        ' one dimensional
        ReDim v(1 To i, 1 To 1)
        ki = 1
        For Each v2 In c
          v(ki, 1) = v2
          k = ki + 1
        Next
      End If
    Else
      v = d(k)
    End If
    Set r = ws.Range(k)
    r.Value2 = v
    If VarType(v) = vbArray Then Erase v
  Next
End Sub
Private Sub RenderComments(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k As Variant, v As Variant, v2 As Variant, r As Range, r2 As Range
  For Each k In d.Keys
    v = d(k)
    For Each v2 In Split(k, ",")
      Set r = ws.Range(v2)
      For Each r2 In r.Cells
        Call r2.AddComment(v)
      Next
    Next
  Next
End Sub
Private Sub RenderFonts(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k, v, v2, r As Range
  Dim dUL As Dictionary, dCI As Dictionary
  Set dUL = xlUnderlineStyleConsts()
  Set dCI = xlColorIndexConsts()
  For Each k In d.Keys
    v = Split(d(k), "|")
    For Each v2 In Split(k, ",")
      Set r = ws.Range(v2)
      With r.Font
        .Name = v(0)
        .Size = CLng(v(1))
        .Bold = CBool(v(2))
        .Italic = CBool(v(3))
          ' Fonts
        .Underline = dUL(v(4))
        If (IsNumeric(v(5))) Then
          .ColorIndex = v(5)
        Else
          .ColorIndex = dCI(v(5))
        End If
      End With
    Next
  Next
End Sub
Private Sub RenderNumerics(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k, v, v2, r As Range
  Dim dV As Dictionary, dH As Dictionary, dOr As Dictionary
  Set dV = xlVAlignConsts()
  Set dH = xlHAlignConsts()
  Set dOr = xlOrientationConsts()
  For Each k In d.Keys
    v = Split(d(k), "|")
    For Each v2 In Split(k, ",")
      Set r = ws.Range(v2)
      r.NumberFormat = v(0)
      r.VerticalAlignment = dV(v(1))
      r.HorizontalAlignment = dH(v(2))
      If (IsNumeric(v(3))) Then
        r.Orientation = v(3)
      Else
        r.Orientation = dOr(v(3))
      End If
      r.WrapText = v(4)
      r.ShrinkToFit = v(5)
    Next
  Next
End Sub
Private Sub RenderColours(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k, v, v2, r As Range
  Dim dCI As Dictionary, dPC As Dictionary
  Set dCI = xlColorIndexConsts()
  Set dPC = xlPatternConsts()
  For Each k In d.Keys
    v = Split(d(k), "|")
    For Each v2 In Split(k, ",")
      Set r = ws.Range(v2)
      With r.Interior
        .Pattern = dPC(v(0))
        If (IsNumeric(v(1))) Then
          .PatternColorIndex = v(1)
        Else
          .PatternColorIndex = dCI(v(1))
        End If
        If (IsNumeric(v(2))) Then
          .ColorIndex = v(2)
        Else
          .ColorIndex = dCI(v(2))
        End If
      End With
    Next
  Next
End Sub
Private Sub RenderBorders(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k, v, v2, r As Range
  Dim dBI As Dictionary
  Dim dLI As Dictionary
  Dim dCI As Dictionary
  Set dBI = xlBordersIndexConsts()
  Set dLI = xlLineStyleConsts()
  Set dCI = xlColorIndexConsts()
  For Each k In d.Keys
    v = Split(d(k), "|")
    For Each v2 In Split(k, ",")
      Set r = ws.Range(v2)
      With r.Borders(dBI(v(0)))
        .LineStyle = dLI(v(1))
        .Weight = v(2)
        If (IsNumeric(v(3))) Then
          .ColorIndex = v(3)
        Else
          .ColorIndex = dCI(v(3))
        End If
      End With
    Next
  Next
End Sub
Private Sub RenderCellSizes(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k As Variant, v As Variant, v2 As Variant, r As Range
  Dim d1 As Dictionary
  ' Cell widths
  If IsObject(d("w")) Then
    Set d1 = d("w")
    For Each k In d1.Keys
      v = d1(k)
      For Each v2 In Split(k, ",")
        ws.Range(v2).ColumnWidth = v
      Next
    Next
  Else
    ws.UsedRange.ColumnWidth = d("w")
  End If
  ' Cell heights
  If IsObject(d("h")) Then
    Set d1 = d("h")
    For Each k In d1.Keys
      v = d1(k)
      For Each v2 In Split(k, ",")
        ws.Range(v2).RowHeight = v
      Next
    Next
  Else
    ws.UsedRange.RowHeight = d("h")
  End If
End Sub
Private Sub RenderOutlines(ByVal d As Dictionary, ByRef ws As Worksheet)
  Dim k As Variant, v As Variant, v2 As Variant, r As Range
  Dim d1 As Dictionary
  If IsObject(d("c")) Then
    Set d1 = d("c")
    For Each k In d1.Keys
      v = d1(k)
      For Each v2 In Split(k, ",")
        ws.Range(v2).Columns.OutlineLevel = v
      Next
    Next
  Else
    ws.UsedRange.Columns.OutlineLevel = d("c")
  End If
  If IsObject(d("r")) Then
    Set d1 = d("r")
    For Each k In d1.Keys
      v = d1(k)
      For Each v2 In Split(k, ",")
        ws.Range(v2).Rows.OutlineLevel = v
      Next
    Next
  Else
    ws.UsedRange.Rows.OutlineLevel = d("r")
  End If
End Sub

' Calls Let on object O, by enumerating keys in d.
' dC should be nested dictionary of Type->ConstName->ConstValue for enumerations
Private Sub LetProperties(ByRef o As Object, ByVal d As Dictionary, ByVal dP As Dictionary, ByVal dC As Dictionary)
  ' Generic Let on an object, based on a KV dictionary
  Dim dT As Dictionary, k, v
  Set dT = modXlConstants.xlBaseTypes()
  For Each k In d.Keys
    If k = "_CodeName" Then
        Debug.Print "Skipping CodeName..."
    Else
    v = d(k)
      If dT.Exists(dP(k)) Then ' built in type
        Call CallByName(o, k, VbLet, v)
      Else
        Call CallByName(o, k, VbLet, (dC(dP(k))(v)))
      End If
    End If
  Next
End Sub
