Attribute VB_Name = "modFormatAuditor"
Option Explicit

' Format auditor counts number and type of formats used in a given workbook
' Encodes formats to string, and creates map (format encoding) => (range)
' Module (c) 2004 / 2017 Dr. D. Azzopardi

' Requires:
' Microsoft Scripting Runtime (for Dictionary object)

' Encode a unique format as follows:
' <Font name> | <size> | {BI} | <underlinestyle> | halign | valign | Number format | LHTPBOS |  [Font col index] | [interior col index] | <border formats>
' {BI} is bold italic underline
' {LCR} is justification
' {LHTPBOS} => Locked, Formula Hidden, Strikethrough, superscript, subscript, outline, shadow
' <border formats> are as follows:
' for each bd in activeCell.Borders: ? trim(bd.linestyle)&":"&bd.color&":"&bd.weight : next

Const ModuleVersion As String = "1.0.0"
Const ModuleName As String = "modFormatAuditor"
Const cstrCacheSpecifier As String = "~.cache"
Const cstrSeptr As String = ";"

' Assume we are running on something with a performance counter:
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Function GetPerfFrequency() As Currency
    Static nFreq As Currency
    If nFreq = 0 Then
        Call QueryPerformanceFrequency(nFreq)
    End If
    GetPerfFrequency = nFreq
End Function

' Returns elapsed time since last mark, in seconds
Public Function TimingMark() As Double
    Static nLastCount As Currency
    Dim nNext As Currency
    Dim nDiff As Currency
    Call QueryPerformanceCounter(nNext)
    nDiff = nNext - nLastCount
    nLastCount = nNext
    TimingMark = CDbl(nDiff) / CDbl(GetPerfFrequency)
End Function


Public Sub ExtractUniqueFormats(Optional ByVal wb As Workbook = Nothing)
    Dim rngCell As Range
    Dim rngSrc As Range
    Dim strKey As String
    Dim sht As Worksheet
    Dim vnt As Variant
    Dim vntX As Variant
    Dim dctFormats As Dictionary
    Dim dctFormatsThisSheet As Dictionary
    Dim dctSheetRanges As Dictionary
    Dim brd As Border
    
    Dim n As Long
    Dim m As Long
    Dim nTot As Long
    
    Const cmStart As Long = 7
    Const cnStart As Long = 4
    
    
    Set dctFormats = New Dictionary
    Set dctSheetRanges = New Dictionary ' This is a dictionary of dictionaries
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    ' Format -> count
    For Each sht In wb.Sheets
        Set rngSrc = sht.UsedRange
        Set dctFormatsThisSheet = New Dictionary '(TextCompare)
        nTot = rngSrc.Cells.Count
        n = 0
        TimingMark
        
        For Each rngCell In rngSrc
            With rngCell
                ' Font name and size
                strKey = .Font.Name & cstrSeptr & .Font.Size & cstrSeptr
                ' Bold, italic, underline
                If (.Font.Bold) Then
                    strKey = strKey & "B"
                End If
                If (.Font.Italic) Then
                    strKey = strKey & "I"
                End If
                ' Underline
                strKey = strKey & cstrSeptr & .Font.Underline
                ' Justification
                strKey = strKey & cstrSeptr & .HorizontalAlignment
                strKey = strKey & cstrSeptr & .VerticalAlignment
                ' Number format
                strKey = strKey & cstrSeptr & .NumberFormat & cstrSeptr
                '{LHTPBOS} => Locked, Formula Hidden, Strikethrough, superscript, subscript, outline, shadow
                If (.Locked) Then
                    strKey = strKey & "L"
                End If
                If (.FormulaHidden) Then
                    strKey = strKey & "H"
                End If
                With .Font
                    If (.Strikethrough) Then
                        strKey = strKey & "T"
                    End If
                    If (.Superscript) Then
                        strKey = strKey & "P"
                    End If
                    If (.Subscript) Then
                        strKey = strKey & "B"
                    End If
                    If (.OutlineFont) Then
                        strKey = strKey & "O"
                    End If
                    If (.Shadow) Then
                        strKey = strKey & "S"
                    End If
                End With
                strKey = strKey & cstrSeptr & .Style
                ' Font color
                strKey = strKey & cstrSeptr & .Font.ColorIndex
                ' Interior color
                strKey = strKey & cstrSeptr & .Interior.ColorIndex
            End With
            ' Borders
            strKey = strKey & cstrSeptr
            For Each brd In rngCell.Borders
                strKey = strKey & brd.LineStyle & ":" & brd.ColorIndex & ":" & brd.Weight & ":"
            Next brd
            ' This sheet:
            Call RangeUnionizeCache(dctFormatsThisSheet, strKey, rngCell)
            ' Total count across book:
            If dctFormats.Exists(strKey) Then
                dctFormats(strKey) = dctFormats(strKey) + 1
            Else
                dctFormats(strKey) = 1
            End If
            
            n = n + 1
            If (n Mod 1000 = 0) Then
                'Application.ScreenUpdating = True
                Application.StatusBar = "Scanning : " & sht.Name & " (" & sht.Index & "/" & ActiveWorkbook.Sheets.Count & ") : Cells " & "(" & n & "/" & nTot & ")"
                'Application.ScreenUpdating = False
            End If
        Next rngCell
        ' Flush caches:
        Call RangeUnionizeFlush(dctFormatsThisSheet)
        Set dctSheetRanges(sht.Name) = dctFormatsThisSheet
        Debug.Print "Formats in " & sht.Name & ": " & dctFormatsThisSheet.Count & " (Scanning took : " & TimingMark & " seconds)"
    Next sht
    Application.StatusBar = "Output results ..."
    Debug.Print "Unique formats in " & wb.Name & ": " & dctFormats.Count
    ' Now output to new sheet in NEW book (use new book so we can have as many formats as allowable by excel...)
    Application.Workbooks.Add
    Set sht = ActiveWorkbook.Sheets(1)
    sht.Activate
    sht.Cells(1, 1) = "Sheet audit for " & wb.Name
    ' Turn off screen updating to speed this up:
    Application.ScreenUpdating = False
    m = cmStart ' Column number
    ' Sheet names:
    For Each vnt In dctSheetRanges
        strKey = vnt
        sht.Cells(2, m) = strKey
        m = m + 1
    Next vnt
    TimingMark
    n = cnStart ' row number on this sheet
    For Each vnt In dctFormats
        ' Put down formats on this sheet in columns A and C
        Set rngCell = sht.Cells(n, 1)
        strKey = vnt
        rngCell.Value = "Text"
        Call ApplyFormat(rngCell, strKey)
        Set rngCell = sht.Cells(n, 3)
        rngCell.Value = 0
        Call ApplyFormat(rngCell, strKey)
        sht.Cells(n, 5).Value = strKey
        m = cmStart

        For Each vntX In dctSheetRanges
            Set dctFormatsThisSheet = dctSheetRanges(vntX)
            If dctFormatsThisSheet.Exists(strKey) Then
                sht.Cells(n, m) = dctFormatsThisSheet(strKey).Count
                sht.Cells(n + 1, m) = dctFormatsThisSheet(strKey).Address
                If (dctFormatsThisSheet(strKey).Count = 1) Then
                    sht.Hyperlinks.Add Anchor:=sht.Cells(n, m), Address:="", SubAddress:="'[wte40317.xls]01_Status'!$A$1"
                End If
            Else
                sht.Cells(n, m) = 0
            End If
            
            m = m + 1
        Next vntX
        n = n + 2
    Next vnt
    Debug.Print "rendering: " & TimingMark & " sec"
    Application.ScreenUpdating = True
    Application.StatusBar = ""
End Sub

Public Sub ApplyFormat(ByVal rng As Range, ByVal strKey As String)
    ' Applys format to a specific range
    Dim vnt As Variant
    Dim vnt2 As Variant
    Dim n As Long
    Dim brd As Border
    vnt = Split(strKey, cstrSeptr)
    On Error Resume Next
    With rng
        ' Style first!
        .Style = vnt(6)
        .Font.Name = vnt(0)
        .Font.Size = vnt(1)
        If vnt(2) Like "*B*" Then
            .Font.Bold = True
        End If
        If vnt(2) Like "*I*" Then
            .Font.Italic = True
        End If
        .Font.Underline = vnt(3)
        .HorizontalAlignment = vnt(4)
        .VerticalAlignment = vnt(5)
        .NumberFormat = vnt(6)
        If vnt(7) Like "*L*" Then .Locked = True
        If vnt(7) Like "*H*" Then .FormulaHidden = True
        If vnt(7) Like "*T*" Then .Font.Strikethrough = True
        If vnt(7) Like "*P*" Then .Font.Superscript = True
        If vnt(7) Like "*B*" Then .Font.Subscript = True
        If vnt(7) Like "*O*" Then .Font.OutlineFont = True
        If vnt(7) Like "*S*" Then .Font.Shadow = True

        .Font.ColorIndex = vnt(8)
        .Interior.ColorIndex = vnt(9)
        vnt2 = Split(vnt(10), ":")
        n = 0
        For Each brd In rng.Borders
            If (vnt2(n) <> xlLineStyleNone) Then
                brd.LineStyle = vnt2(n)
                brd.ColorIndex = vnt2(n + 1)
                brd.Weight = vnt2(n + 2)
            End If
            n = n + 3
        Next brd
    End With
    
End Sub


Private Sub RangeUnionizeCache(ByVal dct As Dictionary, ByVal strKey As String, ByVal rng As Range, Optional ByVal bCoalesce As Boolean = True)
    ' Typically, RangeUnionize is called one cell at a time
    ' Can be significantly sped up if the root range is kept small
    ' Thus, we implement a two tier strategy as follows:
    ' 1) If it doesn't already exist, create a new key, of the form strKey & "~.cache"
    ' 2) Unionize against this range until the number of cells in this range is greater than some threshold
    ' 2a) Once this threshold is breached, append this range to the original root, and
    ' 2b) Reset cache range to nothing
    ' 3) Allow client to force flushing of cache by providing an additional sub that
    '    iterates through a dictionary flushing each cache and removing cache items.
    '    Called at each proto-distiller's termination
    
    Const cNMaxCacheSize As Long = 2000
    Dim strCacheKey As String
    
    strCacheKey = strKey & cstrCacheSpecifier
    If dct.Exists(strCacheKey) Then
        Set rng = Union(dct(strCacheKey), rng)
        If (rng.Cells.Count > cNMaxCacheSize) Then
            If dct.Exists(strKey) Then
                Set rng = Union(dct(strKey), dct(strCacheKey))
                If bCoalesce Then Set rng = Union(rng, rng)
                Set dct(strKey) = rng
            Else
                Set dct(strKey) = rng
            End If
            dct.Remove (strCacheKey)
        Else
            Set dct(strCacheKey) = rng
        End If
    Else
        Set dct(strCacheKey) = rng
    End If
End Sub

Private Sub RangeUnionizeFlush(ByVal dct As Dictionary, Optional ByVal strKey As String = "")
    ' If an individual key is given, flush that. Otherwise, iterate over directory and flush as required
    Dim rng As Range
    Dim rngTgt As Range
    Dim strCacheKey As String
    Dim vKey As Variant
    If strKey <> "" Then
        If (Right(strKey, Len(cstrCacheSpecifier)) = cstrCacheSpecifier) Then
            strCacheKey = strKey
            strKey = Left(strKey, Len(strKey) - Len(cstrCacheSpecifier))
        Else
            strCacheKey = strKey & cstrCacheSpecifier
        End If
        If (dct.Exists(strCacheKey)) Then
            Set rng = dct(strCacheKey)
            If (dct.Exists(strKey)) Then
                Set dct(strKey) = Union(dct(strKey), rng)
            Else
                Set dct(strKey) = dct(strCacheKey)
            End If
            Set dct(strKey) = Union(dct(strKey), dct(strKey))
            dct.Remove (strCacheKey)
        End If
    Else
        For Each vKey In dct
            strKey = vKey
            Call RangeUnionizeFlush(dct, strKey)
        Next vKey
    End If
End Sub


