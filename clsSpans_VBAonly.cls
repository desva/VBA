VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Uses only native types
' Use as a cell-by-cell replacement for Application.Union

Const MODULE_NAME As String = "clsSpans"
Const MODULE_VERSION As String = "1.0.3"
Const cColToken As Long = -1 ' Must be non-valid column/row number

Private m_dctSpans As Dictionary ' maps column number to array list, until we coalesce - then maps to dictionaries
Private m_dctCols As Dictionary ' Hit map of columns
Private m_colOutput As Collection ' once we have coalesced this is populated, and no longer accept new cells
Private m_nCells As Long ' Count of cells accoumulated

Property Get CellCount() As Long
    CellCount = m_nCells
End Property

' Or provide as c,r but remember to transpose when you coalesce
Public Sub AddCell(ByVal r As Long, ByVal c As Long)
    If m_colOutput Is Nothing Then
        If (r >= 0 And c >= 0) Then
            AddCell_Internal r, c
            m_nCells = m_nCells + 1
        End If
    End If
End Sub

Public Function GetRange(ByVal style As XlReferenceStyle) As String
    Dim v As Variant
    If m_colOutput Is Nothing Then Exit Function
    For Each v In m_colOutput
        If style = xlA1 Then
            GetRange = GetRange & Replace(Application.ConvertFormula(v, xlR1C1, xlA1), "$", vbNullString) & ","
        Else
            GetRange = GetRange & v & ","
        End If
    Next
    GetRange = Left$(GetRange, Len(GetRange) - 1)
End Function

Private Sub AddCell_Internal(ByVal r As Long, ByVal c As Long)
    Dim oC As Collection
    If m_dctSpans.Exists(c) Then
        Set oC = m_dctSpans(c)
    Else
        Set oC = New Collection
        Set m_dctSpans(c) = oC
    End If
    oC.Add r
    If (c = cColToken) Then Exit Sub
    ' Column cColToken is fake - each r actually indicates which spans of columns occupied
    If Not m_dctCols.Exists(c) Then
        m_dctCols(c) = 0
        Call AddCell_Internal(c, cColToken)
    End If
End Sub

Public Sub Coalesce(Optional ByVal bTranspose As Boolean = False)
    Dim v As Variant, v1 As Variant
    Dim i As Long, j As Long, jj As Long, rs As Long, re As Long
    Dim dctIn As Dictionary, d0 As Dictionary

    Set m_colOutput = New Collection
    ' Check if we only have one cell:
    If m_nCells = 1 Then
      i = m_dctSpans.Keys(0)
      j = m_dctSpans(i).Item(1)
      If bTranspose Then
        m_colOutput.Add CStr("R" & i & "C" & j)
      Else
        m_colOutput.Add CStr("R" & j & "C" & i)
        Exit Sub
        End If
    End If
    ' First attempt to coalesce each column
    For Each v In m_dctSpans.Keys
        CoalesceCol v
    Next
    
    ' This should also have coalesced column cColToken - we process columns according to this dictionary in order to try to create
    ' rectangular areas
    Set dctIn = m_dctSpans(cColToken)

    For Each v In dctIn.Keys
        ' Start and end column for this span
        i = CLng(v)
        j = dctIn(i)
        Do
            Set d0 = m_dctSpans(i)
            For Each v1 In d0.Keys
                rs = v1
                re = d0(rs)
                jj = i
                ' See how many similar columns there are - remove associated values from relevant dictionaries!
                Do While (jj < j)
                    If m_dctSpans(jj + 1).Exists(rs) Then
                        If m_dctSpans(jj + 1)(rs) = re Then
                            ' remove this span from consideration
                            m_dctSpans(jj + 1).Remove rs
                            ' increase width
                            jj = jj + 1
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
                ' Have an area from (rs,i) to (re,jj), add to output
                If bTranspose Then
                    If (i = jj And rs = re) Then ' single cell
                        m_colOutput.Add CStr("R" & i & "C" & rs)
                    Else
                        m_colOutput.Add CStr("R" & i & "C" & rs & ":" & "R" & jj & "C" & re)
                    End If
                Else
                    If (i = jj And rs = re) Then ' single cell
                        m_colOutput.Add "R" & rs & "C" & i
                    Else
                        m_colOutput.Add "R" & rs & "C" & i & ":" & "R" & re & "C" & jj
                    End If
                End If
            Next
            Set d0 = Nothing
            i = i + 1
        Loop While (i <= j)
    Next
    Set dctIn = Nothing
End Sub

Private Sub CoalesceCol(ByVal c As Long)
    Dim oC As Collection
    Dim i As Long, l As Long, r As Long, m As Long, nArr() As Long, v
    Dim dctSpans As Dictionary
    If Not m_dctSpans.Exists(c) Then Exit Sub
    Set oC = m_dctSpans(c)
    Set dctSpans = New Dictionary
    m = oC.Count - 1
    ReDim nArr(0 To m)
    i = 0
    For Each v In oC
      nArr(i) = v
      i = i + 1
    Next
    Call Quicksort(nArr, 0, m)
    For i = 0 To m
        l = nArr(i)
        r = l
        Do While (i < m)
            If (nArr(i + 1) = r + 1) Then
                r = r + 1
            Else
                Exit Do
            End If
            i = i + 1
        Loop
        dctSpans.Add l, r
    Next
    Set m_dctSpans(c) = dctSpans
End Sub

Private Sub Quicksort(vArray As Variant, arrLbound As Long, arrUbound As Long)
  Dim pivotVal As Variant
  Dim vSwap    As Variant
  Dim tmpLow   As Long
  Dim tmpHi    As Long
   
  tmpLow = arrLbound
  tmpHi = arrUbound
  pivotVal = vArray((arrLbound + arrUbound) \ 2)
   
  While (tmpLow <= tmpHi) 'divide
     While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
        tmpLow = tmpLow + 1
     Wend
    
     While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
        tmpHi = tmpHi - 1
     Wend
   
     If (tmpLow <= tmpHi) Then
        vSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = vSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend
 
  If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi 'conquer
  If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound 'conquer
End Sub

Private Sub Class_Initialize()
    Set m_dctSpans = New Dictionary
    Set m_dctCols = New Dictionary
    m_nCells = 0
End Sub

Private Sub Class_Terminate()
    Set m_dctSpans = Nothing
    Set m_dctCols = Nothing
    Set m_colOutput = Nothing
End Sub

