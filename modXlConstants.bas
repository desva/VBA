Attribute VB_Name = "modXlConstants"
Option Explicit
' Originally used dynamic introspection for this - made code a bit difficult to reason about without running it

Private Function CreateKVDictionary(ByVal Keys As String, ByVal Vals As Variant, ByVal bInvert As Boolean) As Dictionary
  Dim v, k, i As Long
  If Not IsArray(Keys) Then k = Split(Keys, "|") Else k = Keys
  If Not IsArray(Vals) Then v = Split(Vals, "|") Else v = Vals
  Set CreateKVDictionary = New Dictionary
  For i = LBound(k, 1) To UBound(k, 1)
    If bInvert Then
      Call CreateKVDictionary.Add(v(i), k(i))
    Else
      Call CreateKVDictionary.Add(k(i), v(i))
    End If
  Next
End Function
'
' Properties
'
Public Function XlWorkSheetProperties(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "_CodeName|Name|OnDoubleClick|OnSheetActivate|OnSheetDeactivate|Visible|TransitionExpEval|" & _
                        "AutoFilterMode|EnableCalculation|DisplayAutomaticPageBreaks|EnableAutoFilter|EnableSelection|" & _
                        "EnableOutlining|EnablePivotTable|OnCalculate|OnData|OnEntry|ScrollArea|StandardWidth|TransitionFormEntry|" & _
                        "DisplayPageBreaks|DisplayRightToLeft|EnableFormatConditionsCalculation"
  Const cV As String = "vbString|vbString|vbString|vbString|vbString|XlSheetVisibility|vbBoolean|vbBoolean|vbBoolean|vbBoolean|vbBoolean|" & _
                        "XlEnableSelection|vbBoolean|vbBoolean|vbString|vbString|vbString|vbString|vbDouble|vbBoolean|vbBoolean|vbBoolean|vbBoolean"
  Set XlWorkSheetProperties = CreateKVDictionary(ck, cV, Invert)
End Function
Public Function XlWorkSheetViewProperties(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "DisplayGridlines|DisplayFormulas|DisplayHeadings|DisplayOutline|DisplayZeros"
  Const cV As String = "vbBoolean|vbBoolean|vbBoolean|vbBoolean|vbBoolean"
  Set XlWorkSheetViewProperties = CreateKVDictionary(ck, cV, Invert)
End Function
Public Function XlNameProperties(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "Category|MacroType|Name|RefersTo|ShortcutKey|Value|Visible|RefersToR1C1|Comment|WorkbookParameter"
  Const cV As String = "vbString|XlXLMMacroType|vbString|vbVariant|vbString|vbString|vbBoolean|vbVariant|vbString|vbBoolean"
  Set XlNameProperties = CreateKVDictionary(ck, cV, Invert)
End Function
Public Function XlHyperlinkProperties(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "SubAddress|Address|EmailSubject|ScreenTip|TextToDisplay"
  Const cV As String = "vbString|vbString|vbString|vbString|vbString"
  Set XlHyperlinkProperties = CreateKVDictionary(ck, cV, Invert)
End Function
'
' Constants
'
Public Function XlSheetVisibilityConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlSheetVisible|xlSheetHidden|xlSheetVeryHidden"
  Dim v
  v = Array(xlSheetVisible, xlSheetHidden, xlSheetVeryHidden)
  Set XlSheetVisibilityConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function XlEnableSelectionConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlNoRestrictions|xlUnlockedCells|xlNoSelection"
  Dim v
  v = Array(xlNoRestrictions, xlUnlockedCells, xlNoSelection)
  Set XlEnableSelectionConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function XlXLMMacroTypeConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlCommand|xlFunction|xlNotXLM|xlNone" ' Can also be xlNone!
  Dim v
  v = Array(xlCommand, xlFunction, xlNotXLM, xlNone)
  Set XlXLMMacroTypeConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlUnderlineStyleConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlUnderlineStyleDouble|xlUnderlineStyleDoubleAccounting|xlUnderlineStyleNone|xlUnderlineStyleSingle|xlUnderlineStyleSingleAccounting"
  Dim v
  v = Array(xlUnderlineStyleDouble, xlUnderlineStyleDoubleAccounting, xlUnderlineStyleNone, xlUnderlineStyleSingle, xlUnderlineStyleSingleAccounting)
  Set xlUnderlineStyleConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlColorIndexConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlColorIndexAutomatic|xlColorIndexNone"
  Dim v
  v = Array(xlColorIndexAutomatic, xlColorIndexNone)
  Set xlColorIndexConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlOrientationConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlDownward|xlHorizontal|xlUpward|xlVertical"
  Dim v
  v = Array(xlDownward, xlHorizontal, xlUpward, xlVertical)
  Set xlOrientationConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlVAlignConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlVAlignBottom|xlVAlignCenter|xlVAlignDistributed|xlVAlignJustify|xlVAlignTop"
  Dim v
  v = Array(xlVAlignBottom, xlVAlignCenter, xlVAlignDistributed, xlVAlignJustify, xlVAlignTop)
  Set xlVAlignConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlHAlignConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlHAlignCenter|xlHAlignCenterAcrossSelection|xlHAlignDistributed|xlHAlignFill|xlHAlignGeneral|xlHAlignJustify|xlHAlignLeft|xlHAlignRight"
  Dim v
  v = Array(xlHAlignCenter, xlHAlignCenterAcrossSelection, xlHAlignDistributed, xlHAlignFill, xlHAlignGeneral, xlHAlignJustify, xlHAlignLeft, xlHAlignRight)
  Set xlHAlignConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlPatternConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlPatternAutomatic|xlPatternChecker|xlPatternCrissCross|xlPatternDown|xlPatternGray16|xlPatternGray25|xlPatternGray50|" & _
                        "xlPatternGray75|xlPatternGray8|xlPatternGrid|xlPatternHorizontal|xlPatternLightDown|xlPatternLightHorizontal|xlPatternLightUp|" & _
                        "xlPatternLightVertical|xlPatternNone|xlPatternSemiGray75|xlPatternSolid|xlPatternUp|xlPatternVertical|xlPatternLinearGradient|" & _
                        "xlPatternRectangularGradient"
  Dim v
  v = Array(xlPatternAutomatic, xlPatternChecker, xlPatternCrissCross, xlPatternDown, xlPatternGray16, xlPatternGray25, _
    xlPatternGray50, xlPatternGray75, xlPatternGray8, xlPatternGrid, xlPatternHorizontal, xlPatternLightDown, xlPatternLightHorizontal, _
    xlPatternLightUp, xlPatternLightVertical, xlPatternNone, xlPatternSemiGray75, xlPatternSolid, xlPatternUp, xlPatternVertical, _
    xlPatternLinearGradient, xlPatternRectangularGradient)
  Set xlPatternConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlLineStyleConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlContinuous|xlDash|xlDashDot|xlDashDotDot|xlDot|xlDouble|xlSlantDashDot|xlLineStyleNone"
  Dim v
  v = Array(xlContinuous, xlDash, xlDashDot, xlDashDotDot, xlDot, xlDouble, xlSlantDashDot, xlLineStyleNone)
  Set xlLineStyleConsts = CreateKVDictionary(ck, v, Invert)
End Function
Public Function xlBordersIndexConsts(Optional ByVal Invert As Boolean = False) As Dictionary
  Const ck As String = "xlEdgeLeft|xlEdgeRight|xlEdgeTop|xlEdgeBottom|xlDiagonalDown|xlDiagonalUp"
  Dim v
  v = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlDiagonalDown, xlDiagonalUp)
  Set xlBordersIndexConsts = CreateKVDictionary(ck, v, Invert)
End Function
' The serializable base types. Not really sure about variant!
Public Function xlBaseTypes() As Dictionary
  Const ck As String = "vbBoolean|vbString|vbDouble|vbVariant"
  Set xlBaseTypes = CreateKVDictionary(ck, ck, True)
End Function



