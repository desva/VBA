Attribute VB_Name = "modFormulaParser"
Option Explicit

' Formula parsing
' Based on Rob van Gelder's parser, which is a port of Eric W. Bachtal's javascript parser
' Updated to do RC type formulas, Excel 2007 expanded sheet size
' Bug fix (to do) - inaccurate reference resolution for


Public Const tkt_Operand = 2 ^ 0
Public Const tkt_OperandUnknown = tkt_Operand Or 2 ^ 1
Public Const tkt_OperandText = tkt_Operand Or 2 ^ 2
Public Const tkt_OperandError = tkt_Operand Or 2 ^ 3
Public Const tkt_OperandNumber = tkt_Operand Or 2 ^ 4
Public Const tkt_OperandBoolean = tkt_Operand Or 2 ^ 5
Public Const tkt_OperandReference = tkt_Operand Or 2 ^ 6
Public Const tkt_OperandReferenceWksQual = tkt_OperandReference Or 2 ^ 7
Public Const tkt_OperandReference3DWksQual = tkt_OperandReferenceWksQual Or 2 ^ 8
Public Const tkt_OperandReferenceRange = tkt_OperandReference Or 2 ^ 9
Public Const tkt_OperandReferenceName = tkt_OperandReference Or 2 ^ 10

Public Const tkt_WhiteSpace = 2 ^ 11

Public Const tkt_OperatorPrefix = 2 ^ 12
Public Const tkt_OperatorInfix = 2 ^ 13
Public Const tkt_OperatorPostfix = 2 ^ 14

Public Const tkt_OperatorArithmetic = 2 ^ 15
Public Const tkt_OperatorComparison = 2 ^ 16
Public Const tkt_OperatorText = 2 ^ 17
Public Const tkt_OperatorReference = 2 ^ 18

Public Const tkt_Expression = 2 ^ 19
Public Const tkt_Function = 2 ^ 20
Public Const tkt_FunctionArgument = 2 ^ 21
Public Const tkt_Array = 2 ^ 22
Public Const tkt_ArrayCol = 2 ^ 23
Public Const tkt_ArrayRow = 2 ^ 24

Public Const tkt_Begin = 2 ^ 25
Public Const tkt_End = 2 ^ 26

'States we can be in as we parse string
Private Const cStateDefault = 2 ^ 0
Private Const cStateArray = 2 ^ 1
Private Const cStateText = 2 ^ 2
Private Const cStateWksQuote = 2 ^ 3
Private Const cStateSqBracket = 2 ^ 4
Private Const cStateError = 2 ^ 5

Const cRefID As String = "Refs@" ' @ is disallowed character for sheet names

Public Type Token
    strValue As String
    lngType As Long
End Type

Public Function ParseFormula(strFormula As String) As Token()
    Dim lngState As Long, str As String, strC As String
    Dim i As Long, j As Long, k As Long
    Dim udtTokens() As Token, udtTokenStack() As Token

    Dim strLeftBrace As String, strRightBrace As String
    Dim strColumnSeparator As String, strRowSeparator As String, strListSeparator As String

    strLeftBrace = Application.International(xlLeftBrace)
    strRightBrace = Application.International(xlRightBrace)
    strColumnSeparator = Application.International(xlColumnSeparator)
    strRowSeparator = Application.International(xlRowSeparator)
    strListSeparator = Application.International(xlListSeparator)

    lngState = cStateDefault
    i = 1

    If Left(strFormula, 1) = "=" Then i = i + 1

    Do Until i > Len(strFormula)
        strC = Mid(strFormula, i, 1)

        If (lngState And cStateText) = cStateText Then
            If strC = """" Then
                If Mid(strFormula, i + 1, 1) = strC Then
                    str = str & strC
                    i = i + 1
                Else
                    TokenPush udtTokens, str, tkt_OperandText
                    lngState = lngState And Not cStateText
                End If
            Else
                str = str & strC
            End If

        ElseIf (lngState And cStateWksQuote) = cStateWksQuote Then
            If strC = "'" Then
                If Mid(strFormula, i + 1, 1) = strC Then
                    str = str & strC
                    i = i + 1
                Else
                    lngState = lngState And Not cStateWksQuote
                End If
            Else
                str = str & strC
            End If

        ElseIf (lngState And cStateSqBracket) = cStateSqBracket Then
            If strC = "[" Then
                j = j + 1
            ElseIf strC = "]" Then
                If j = 0 Then lngState = lngState And Not cStateSqBracket Else j = j - 1
            End If
            str = str & strC

        ElseIf (lngState And cStateError) = cStateError Then
            str = str & strC
            If str = "#NULL!" Or str = "#DIV/0!" Or str = "#VALUE!" Or str = "#REF!" Or _
               str = "#NAME?" Or str = "#NUM!" Or str = "#N/A" Then

                TokenPush udtTokens, str, tkt_OperandError
                lngState = lngState And Not cStateError
            End If

        ElseIf (lngState And cStateDefault) = cStateDefault Then
            If strC = strLeftBrace Then
                lngState = (lngState And Not cStateDefault Or cStateArray)
                j = tkt_Array Or tkt_Begin
                TokenPush udtTokens(), strC, j
                TokenPush udtTokenStack(), strC, j

            ElseIf strC = """" Then
                lngState = lngState Or cStateText

            ElseIf strC = "'" Then
                lngState = lngState Or cStateWksQuote

            ElseIf strC = "[" Then
                j = 0
                str = str & strC
                lngState = lngState Or cStateSqBracket

            ElseIf strC = "#" Then
                str = str & strC
                lngState = lngState Or cStateError

            ElseIf strC = "!" Then
                j = TokenCount(udtTokens())
                If j >= 1 Then
                    If (udtTokens(j).lngType And (tkt_OperatorInfix Or tkt_OperatorReference)) = (tkt_OperatorInfix Or tkt_OperatorReference) And _
                        udtTokens(j).strValue = ":" And _
                       (udtTokens(j - 1).lngType And tkt_Operand) = tkt_Operand Then
                        str = udtTokens(j - 1).strValue & ":" & str
                        TokenPop udtTokens(), True
                        TokenPop udtTokens(), True
                    End If
                End If
                TokenPush udtTokens(), str, IIf(InStr(1, str, ":") = 0, tkt_OperandReferenceWksQual, tkt_OperandReference3DWksQual)
                TokenPush udtTokens(), strC, tkt_OperatorInfix Or tkt_OperatorReference

            ElseIf strC = "+" Or strC = "-" Then
                If str <> "" Then
                    If Right(str, 1) = "E" And IsNumeric(Left(str, Len(str) - 1)) Then
                        str = str & strC
                    Else
                        TokenPush udtTokens(), str, tkt_OperandUnknown
                    End If
                End If
                If str = "" Then
                    j = TokenPop(udtTokens, False).lngType
                    If ((j And (tkt_Array Or tkt_End)) = (tkt_Array Or tkt_End) Or _
                        (j And (tkt_Function Or tkt_End)) = (tkt_Function Or tkt_End) Or _
                        (j And (tkt_Expression Or tkt_End)) = (tkt_Expression Or tkt_End) Or _
                        (j And tkt_Operand) = tkt_Operand Or _
                        (j And tkt_OperatorPostfix) = tkt_OperatorPostfix) Then
                        j = tkt_OperatorInfix Or tkt_OperatorArithmetic
                    Else
                        j = tkt_OperatorPrefix Or tkt_OperatorArithmetic
                    End If
                    TokenPush udtTokens, strC, j
                End If

            ElseIf strC = "*" Or strC = "/" Or strC = "^" Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                TokenPush udtTokens, strC, tkt_OperatorInfix Or tkt_OperatorArithmetic

            ElseIf strC = "%" Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                TokenPush udtTokens, strC, tkt_OperatorPostfix Or tkt_OperatorArithmetic

            ElseIf strC = "=" Or strC = ">" Or strC = "<" Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                Select Case strC & Mid(strFormula, i + 1, 1)
                    Case ">=", "<=", "<>"
                        strC = strC & Mid(strFormula, i + 1, 1)
                        i = i + 1
                End Select
                TokenPush udtTokens(), strC, tkt_OperatorInfix Or tkt_OperatorComparison

            ElseIf strC = "&" Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                TokenPush udtTokens, strC, tkt_OperatorInfix Or tkt_OperatorText

            ElseIf strC = ":" Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                TokenPush udtTokens, strC, tkt_OperatorInfix Or tkt_OperatorReference

            ElseIf strC = " " Or strC = vbLf Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                str = strC
                Do
                    strC = Mid(strFormula, i + 1, 1)
                    If strC = " " Or strC = vbLf Then
                        str = str & strC
                        i = i + 1
                    Else
                        Exit Do
                    End If
                Loop
                TokenPush udtTokens(), str, tkt_WhiteSpace

            ElseIf strC = "(" Then
                j = IIf(str = "", tkt_Expression, tkt_Function) Or tkt_Begin
                str = str & strC
                TokenPush udtTokens(), str, j
                TokenPush udtTokenStack(), str, j

            ElseIf strC = ")" Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                j = TokenPop(udtTokenStack(), True).lngType
                TokenPush udtTokens(), strC, j And Not tkt_Begin Or tkt_End

            ElseIf strC = strListSeparator Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown

                If (TokenPop(udtTokenStack(), False).lngType And tkt_Function) = tkt_Function Then
                    TokenPush udtTokens(), strC, tkt_FunctionArgument
                Else
                    TokenPush udtTokens(), strC, tkt_OperatorInfix Or tkt_OperatorReference
                End If

            Else
                str = str & strC

            End If

        ElseIf (lngState And cStateArray) = cStateArray Then
            If strC = strRightBrace Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                lngState = (lngState And Not cStateArray Or cStateDefault)
                j = TokenPop(udtTokenStack(), True).lngType
                TokenPush udtTokens(), strC, j And Not tkt_Begin Or tkt_End

            ElseIf strC = """" Then
                lngState = lngState Or cStateText

            ElseIf strC = "#" Then
                str = str & strC
                lngState = lngState Or cStateError

            ElseIf strC = strRowSeparator Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                TokenPush udtTokens(), strC, tkt_ArrayRow

            ElseIf strC = strColumnSeparator Then
                If str <> "" Then TokenPush udtTokens(), str, tkt_OperandUnknown
                TokenPush udtTokens(), strC, tkt_ArrayCol

            Else
                str = str & strC

            End If

        End If

        i = i + 1

    Loop

    If str <> "" Then TokenPush udtTokens, str, tkt_OperandUnknown

    j = TokenCount(udtTokens)

    i = 1
    Do Until i > j - 1
        If (udtTokens(i).lngType And tkt_WhiteSpace) = tkt_WhiteSpace Then
            If ((udtTokens(i - 1).lngType And (tkt_Function Or tkt_End)) = (tkt_Function Or tkt_End) Or _
                (udtTokens(i - 1).lngType And (tkt_Expression Or tkt_End)) = (tkt_Expression Or tkt_End) Or _
                (udtTokens(i - 1).lngType And tkt_Operand) = tkt_Operand) And _
               ((udtTokens(i + 1).lngType And (tkt_Function Or tkt_Begin)) = (tkt_Function Or tkt_Begin) Or _
                (udtTokens(i + 1).lngType And (tkt_Expression Or tkt_Begin)) = (tkt_Expression Or tkt_Begin) Or _
                (udtTokens(i + 1).lngType And tkt_Operand) = tkt_Operand) Then
                udtTokens(i).lngType = tkt_OperatorInfix Or tkt_OperatorReference
            End If

        ElseIf (udtTokens(i).lngType And tkt_OperatorReference) = tkt_OperatorReference And udtTokens(i).strValue = ":" And _
               (udtTokens(i - 1).lngType And tkt_Operand) = tkt_Operand And _
               (udtTokens(i + 1).lngType And tkt_Operand) = tkt_Operand Then
            If IsColumn(udtTokens(i - 1).strValue) And IsColumn(udtTokens(i + 1).strValue) Then
                udtTokens(i - 1).strValue = udtTokens(i - 1).strValue & ":" & udtTokens(i + 1).strValue
                udtTokens(i - 1).lngType = tkt_OperandReferenceRange
                TokenPop udtTokens, True, i
                TokenPop udtTokens, True, i
                j = j - 2
                i = i - 1
            ElseIf IsRow(udtTokens(i - 1).strValue) And IsRow(udtTokens(i + 1).strValue) Then
                udtTokens(i - 1).strValue = udtTokens(i - 1).strValue & ":" & udtTokens(i + 1).strValue
                udtTokens(i - 1).lngType = tkt_OperandReferenceRange
                TokenPop udtTokens, True, i
                TokenPop udtTokens, True, i
                j = j - 2
                i = i - 1
            End If
        End If

        i = i + 1
    Loop

    For i = 0 To j
        If (udtTokens(i).lngType And tkt_OperandUnknown) = tkt_OperandUnknown Then
            str = udtTokens(i).strValue

            If IsNumeric(str) Then
                udtTokens(i).lngType = tkt_OperandNumber
            Else
                If UCase(str) = UCase(True) Or UCase(str) = UCase(False) Then
                    udtTokens(i).lngType = tkt_OperandBoolean
                Else
                    If IsReferenceA1orRC(str) Then
                        udtTokens(i).lngType = tkt_OperandReferenceRange
                    Else
                        udtTokens(i).lngType = tkt_OperandReferenceName
                    End If
                End If
            End If
        End If
    Next

    ParseFormula = udtTokens()
End Function

Public Function TokenCount(udtTokens() As Token)
    On Error Resume Next
    TokenCount = -1: TokenCount = UBound(udtTokens)
End Function

Private Function IsColumn(strReference As String) As Boolean
' relative: A:A,C[-1]; Absolute: $A:$A, C[2]
    Dim str As String, i As Long, bln As Boolean
    i = 1
    If Left(strReference, 1) = "$" Then i = i + 1
    str = UCase(Mid(strReference, i))
    If str Like "[A-W][A-Z][A-Z]" Or str Like "X[A-E][A-Z]" Or str Like "XF[A-D]" Then
        bln = True
    ElseIf str Like "[A-Z][A-Z]" Or str Like "I[A-V]" Then
        bln = True
    ElseIf str Like "[A-Z]" Then
        bln = True
    Else
        bln = False
    End If
    IsColumn = bln
End Function

Private Function IsRow(strReference As String) As Boolean
' relative: 2:2,R[-1]; Absolute: $2:$2, R[2]

    Dim str As String, i As Long, lng As Long, bln As Boolean

    bln = True
    i = 1
    If Left(strReference, 1) = "$" Then i = i + 1
    str = Mid(strReference, i)
    If IsNumeric(str) Then
        lng = str
        If lng = str Then
            If Not (lng >= 1 And lng <= 1048576) Then bln = False
        Else
            bln = False
        End If
    Else
        bln = False
    End If
    IsRow = bln
End Function

Private Function IsReferenceA1orRC(strReference As String) As Boolean
    Dim str As String, i As Long, lng As Long, bln As Boolean

    bln = True
    i = 1
    If Left(strReference, 1) = "$" Then i = i + 1
    str = UCase(Mid(strReference, i, 2))
    If str Like "[A-H][A-Z]" Or str Like "I[A-V]" Then
        i = i + 2
    ElseIf str Like "[A-Z]#" Then
        i = i + 1
    ElseIf str Like "[A-Z]$" Then
        i = i + 1
    Else
        bln = False
    End If
    If bln Then
        If Mid(strReference, i, 1) = "$" Then i = i + 1
        str = Mid(strReference, i)

        If IsNumeric(str) Then
            lng = str
            If lng = str Then
                If Not (lng >= 1 And lng <= 1048576) Then bln = False
            Else
                bln = False
            End If
        Else
            bln = False
        End If
    End If
    ' Also deal with RC type references
    IsReferenceA1orRC = bln Or (strReference Like "R*C*") Or (Left(strReference, 2) = "R[") Or (Left(strReference, 2) = "C[")
End Function

Private Sub TokenPush(udtTokens() As Token, strValue As String, lngType As Long)
    Dim i As Long
    i = TokenCount(udtTokens()) + 1
    ReDim Preserve udtTokens(i)
    udtTokens(i).strValue = strValue
    udtTokens(i).lngType = lngType
    strValue = ""
End Sub

Private Function TokenPop(udtTokens() As Token, blnRemove As Boolean, Optional lngOffset As Long = -1) As Token
    Dim i As Long, lngBound As Long
    lngBound = -1: On Error GoTo e: lngBound = UBound(udtTokens): On Error GoTo 0
    If lngOffset <> -1 Then i = lngOffset Else i = lngBound
    TokenPop.strValue = udtTokens(i).strValue
    TokenPop.lngType = udtTokens(i).lngType
    If blnRemove Then
        If lngBound = 0 Then
            Erase udtTokens
        Else
            If lngOffset <> -1 Then
                For i = lngOffset To lngBound - 1
                    udtTokens(i) = udtTokens(i + 1)
                Next
            End If
            ReDim Preserve udtTokens(lngBound - 1)
        End If
    End If
e:
End Function

Public Function TokenTypeDescription(TokenType As Long) As String
    Dim str As String
    Select Case TokenType
        Case tkt_OperandUnknown: str = "Operand Unknown"
        Case tkt_OperandText: str = "Operand Text"
        Case tkt_OperandError: str = "Operand Error"
        Case tkt_OperandNumber: str = "Operand Number"
        Case tkt_OperandBoolean: str = "Operand Boolean"
        Case tkt_OperandReferenceWksQual: str = "Operand Worksheet"
        Case tkt_OperandReference3DWksQual: str = "Operand Worksheet 3D"
        Case tkt_OperandReferenceRange: str = "Operand Reference Range"
        Case tkt_OperandReferenceName: str = "Operand Reference Named Range"
        
        Case tkt_WhiteSpace: str = "White Space"

        Case (tkt_OperatorPrefix Or tkt_OperatorArithmetic): str = "Operator Arithmetic Prefix"
        Case (tkt_OperatorInfix Or tkt_OperatorArithmetic): str = "Operator Arithmetic Infix"
        Case (tkt_OperatorPostfix Or tkt_OperatorArithmetic): str = "Operator Arithmetic Postfix"
        Case (tkt_OperatorInfix Or tkt_OperatorComparison): str = "Operator Comparison Infix"
        Case (tkt_OperatorInfix Or tkt_OperatorText): str = "Operator Text Infix"
        Case (tkt_OperatorInfix Or tkt_OperatorReference): str = "Operator Reference Infix"

        Case (tkt_Begin Or tkt_Expression): str = "Expression Begin"
        Case (tkt_End Or tkt_Expression): str = "Expression End"

        Case (tkt_Begin Or tkt_Function): str = "Function Begin"
        Case (tkt_End Or tkt_Function): str = "Function End"
        Case tkt_FunctionArgument: str = "Function Argument"

        Case (tkt_Begin Or tkt_Array): str = "Array Begin"
        Case (tkt_End Or tkt_Array): str = "Array End"
        Case tkt_ArrayCol: str = "Array Column"
        Case tkt_ArrayRow: str = "Array Row"
    End Select

    TokenTypeDescription = str
End Function

Private Function GatherNames(wb As Workbook) As Object
    Dim d As Object, nm As name, c
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each nm In wb.Names
        If InStr(1, nm.name, "!") > 0 Then
            Call d.Add(nm.name, nm.Parent.name)
        Else
            c = nm.RefersToRange.Parent.name
            If c <> "" Then
                Call d.Add(nm.name, c)
            Else
                Call d.Add(nm.name, "Error: " & nm.name)
            End If
        End If
    Next
    Set GatherNames = d
End Function

Private Function GatherDependencies(wb As Workbook) As Object
    Dim ws As Worksheet
    Dim r As Range, c As Range
    Dim ts() As Token, t As Token, i As Long, j As Long
    Dim d As Object, d2 As Object, d3 As Object
    Dim s As String
    Debug.Print Now
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        Set d2 = CreateObject("Scripting.Dictionary") ' Unique RC formulas
        Set d3 = CreateObject("Scripting.Dictionary") ' Reference -> weight
        Set d(ws.name) = d2
        Set d(cRefID & ws.name) = d3
        i = 0
        Err.Clear
        Set r = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number = 0 Then
            For Each c In r.Cells
                s = c.FormulaR1C1
                If Not d2.Exists(s) Then ' d2: track unique formulas
                    i = i + 1
                    ts = ParseFormula(s)
                    For j = LBound(ts) To UBound(ts)
                        t = ts(j)
                        If t.lngType = tkt_OperandReferenceName Or _
                        t.lngType = tkt_OperandReferenceWksQual Then
                            If Left(t.strValue, 2) = "C[" Or Left(t.strValue, 2) = "R[" Then
                                Debug.Print t.strValue
                                Debug.Assert False
                            End If
                            If Not d3.Exists(t.strValue) Then ' d3: references for this sheet
                                Call d3.Add(t.strValue, 1)
                            Else
                                d3(t.strValue) = d3(t.strValue) + 1
                            End If
                        End If
                    Next
                    Call d2.Add(s, 1)
                    Erase ts
                Else
                    d2(s) = d2(s) + 1
                End If
            Next
        End If
    Debug.Print ws.name & ": " & i & " unique formulas parsed. (" & Now & ")"
    Next
    Set GatherDependencies = d
End Function

Private Function ResolveLinks(ByRef dNames As Object, ByRef dDeps As Object) As Object
    Dim dOut As Object, d As Object, dI As Object
    Dim strKey As String
    Set dOut = CreateObject("Scripting.Dictionary")
    Dim u, v
    For Each u In dDeps.Keys
        If Left(u, Len(cRefID)) = cRefID Then
            Set d = dDeps(u)
            Set dI = CreateObject("Scripting.Dictionary")
            Set dOut(Mid(u, Len(cRefID) + 1)) = dI ' Sheet name -> Object of (precedent sheet name, weight)
            For Each v In d.Keys
                If dNames.Exists(v) Then
                    strKey = dNames(v)
                Else
                    strKey = v ' Direct link
                End If
                If dI.Exists(strKey) Then
                    dI(strKey) = dI(strKey) + 1
                Else
                    dI(strKey) = 1
                End If
            Next
        End If
    Next
    Set ResolveLinks = dOut
End Function

Private Sub WriteOutput(ByRef wb As Workbook, ByRef dOutput As Object)
    Dim dIndex As Object, d As Object
    Dim ws As Worksheet
    Dim t, u, v, i As Long
    Set dIndex = CreateObject("Scripting.Dictionary")
    ReDim u(1 To dOutput.Count + 1, 1 To dOutput.Count + 1)
    i = 1
    ' Phase I - Index all sheets (and prep output array)
    For Each v In dOutput.Keys
        dIndex(v) = i
        u(1, i + 1) = v
        u(i + 1, 1) = v
        i = i + 1
    Next
    ' Phase II - write adjacency matrix, row by row
    i = 1
    For Each v In dOutput.Keys
        Set d = dOutput(v)
        For Each t In d.Keys
            u(i + 1, dIndex(t) + 1) = d(t)
        Next
        i = i + 1
    Next
    Set ws = wb.Worksheets.Add
    ws.Cells(1, 1).Resize(i - 1, i - 1) = u
    ws.Select
    ActiveWindow.Zoom = 75
    ActiveWindow.DisplayGridlines = False
    Erase u
End Sub

Public Sub SheetDependencies()
    Dim wb As Workbook
    Dim dOutput As Object, dNames As Object, dDeps As Object
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    ' Phase 0 - Map names to sheets
    Set dNames = GatherNames(wb)
    ' Phase 1 - Sheets to dependencies
    Set dDeps = GatherDependencies(wb)
    ' Phase 2 - Join to resolve
    Set dOutput = ResolveLinks(dNames, dDeps)
    ' Phase 3 - Output
    Call WriteOutput(wb, dOutput)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Function UnitTest(ByVal strIn As String, ByVal bOutput As Boolean) As String
    Dim a() As Token, t As Token, i As Long, str As String
    a = ParseFormula(strIn)
    If bOutput Then Debug.Print strIn
    str = "["
    For i = LBound(a) To UBound(a)
        t = a(i)
        If bOutput Then Debug.Print t.strValue & vbTab & ": " & TokenTypeDescription(t.lngType)
        str = str & "[" & t.strValue & "," & Hex(t.lngType) & "];"
    Next
    str = Left(str, Len(str) - 1) & "]"
    UnitTest = str
End Function

Public Sub TestMe()
    Dim strTests, v, u, strValidate
    strTests = "SUM(XEB:XEB)@" & _
                "SUM(C[1])@" & _
                "SUM(R2)@" & _
                "SUM($2:$2)@" & _
                "SUM(2:2)@" & _
                "INDEX(Name,1,2)@" & _
                "+(A1=B1)@" & _
                "NOT(RC[1]<>+RC[2])@" & _
                "SUM('Sheet1'!A1:A10)@" & _
                "SUM(Sheet1:Sheet3!A1:C5)@" & _
                "SUM('A1:A3'!A1:C5)@" & _
                "SUM('Sheet1'!A1:'Sheet1'!A10)@"
    strTests = strTests & _
                "IF(TODAY(),Function(Named.Range,""String"",""CCY""))@" & _
                "Function1(Funk(""V"",{""String"",""string""},VolatileFunction()))@" & _
                "CreateThing(NamedRange,""Close"",R[2]C)@" & _
                "ComplexFunc(R4C4,fun1(R[2]C:R[18]C,R[2]C17:R[18]C17),Fun2(R[2]C:R[18]C,R[2]C18:R[18]C18),Fun3(R[2]C:R[18]C,R[2]C24:R[18]C24),Fun4(R[2]C:R[18]C,R[2]C22:R[18]C22),Fun0(R[2]C:R[18]C,R[2]C21:R[18]C21),Fun0(R[2]C:R[18]C,R[2]C20:R[18]C20),Fun0(R[2]C:R[18]C,R[2]C:R[18]C),Fun0(R[2]C:R[18]C,ISNUMBER(R[2]C:R[18]C)),""Argument1"")@" & _
                "RC[-22]&""Y""@" & _
                "IFERROR(Lookup(R30C3,""Ticker"",RC[-22],""Value"")/100,"""")@" & _
                "IF(R[-1]C24="""","""",R[-1]C24&""-"")&RC24@" & _
                "BBGCall(Foo(R[-17]C:R[-1]C[12],,1))@" & _
                "IFERROR(10000*(R[-2]C-R[-3]C),"""")@" & _
                "IF(MOD(dates.today,7)>1.75,IF(MOD(dates.today,7)<2,BumpDate(dates.today,""0d""),dates.today),BumpDate(dates.today,""0m"",,""Preceding""))@" & _
                "IF(AND(MONTH(R[-1]C[1])=7,R[-12]C[3]),""6M"", ""1Y"")" & _
                "-13@" & _
                "-13/2+8*RC[-1]-3/2*RC[-1]^2@" & _
                "IF(bUseName,IF(reports.useReport,R[1]C,FuncCreate(initial.Name,Fun1(FullRefresh!R7C21:R302C21,FullRefresh!R7C11:R302C11),Fun2(FullRefresh!R7C21:R302C21,FullRefresh!R7C12:R302C12),Fun3(FullRefresh!R7C21:R302C21,JacobianRefresh!R7C17:R302C17),Fun6(FullRefresh!R7C21:R302C21,FullRefresh!R7C10:R302C10),Fun0(FullRefresh!R7C21:R302C21,FullRefresh!R7C9:R302C9))))@" & _
                """FilterIt(WakeTime(CreateReport(""""""&RC[-4]&"""""",""""""&RC[-3]&"""""",""""{term}"""",""""""&RC[-2]&""""""),""""""&TEXT(closeDate-200,""dd-mmm-yyyy"")&"""""",""""""&TEXT(closeDate,""dd-mmm-yyyy"")&""""""),""""""&RC[-5]&"""""")""@" & _
                "SUM(R[7])@" & _
                "FetchReport(Output.Filtered,RC[-1])@" & _
                "INDEX(Project1Y!R58,1,MATCH(Project1Y!R[6]C[-1],2Y!R18,0))@" & _
                "VLOOKUP(RC109,C1:C5,MATCH(R2C,R2C109:R2C113,0),0)@" & _
                "#REF!@" & _
                "IFERROR(1*'Data (raw data)'!R[3]C, """")@"
    strTests = strTests & _
                "'Estimate (raw data)'!RC[5]=""Passive""@" & _
                "N(12)@" & _
                "InFill(rates.data, """",TRUE)@" & _
                "SUM(INDIRECT(ADDRESS(ROW(),COLUMN()+1)):INDIRECT(ADDRESS(ROW(),COLUMN()+3*Items.Count+1)))"

    ' problematic:
    strTests = "SUM(Sheet1:Sheet3!A1:C5)@SUM('Sheet1'!A1:'Sheet1'!A10)@"
    strTests = "SUM(Sheet1!A1:Sheet1!A10)"
    v = Split(strTests, "@")
    strValidate = ""
    For Each u In v
        'Debug.Print String$(60, "=")
        strValidate = UnitTest("=" & u, True) & ";"
        Debug.Print strValidate
    Next
End Sub
