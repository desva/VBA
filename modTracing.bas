Attribute VB_Name = "modTracing"
Option Explicit
 
' Old VBA profiler - review 2018 for 64-bit Excel / VBA
' (c) 2010-2018 Dr. D. E. Azzopardi
' Requires reference to Microsoft Visual Basic for Applications extensibility 5.3
 
Const cstrThisModuleName As String = "modTracing"    ' Important this matches to avoid instrumenting this module
Const cstrInstObjName As String = "clsInstrumentation" ' Same comment, but especially important for object as it is used at runtime
Const cstrVersion As String = "1.0.2"

Const cstrInstrumentationMarker As String = "InstrumentationAlreadyCreated" ' marker on a workbook that indicates we've already instrumented - don't re-instrument!
Const cstrTraceMarker As String = "' Trace :"

' %1 Object name, %2: Method marker
Const cstrInstrumentationCode As String = "TraceLog ""%2"" ' %1 "
    ' "Dim %1 As clsInstrumentation: Set %1 = New clsInstrumentation: Call %1.Initialize(""%2"")"

' No VT_GUID available so must declare type GUID
' CreateObject("Scriptlet.TypeLib") No longer considered secure, so disabled on modern OS
'
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (Guid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (Guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr
Private Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (ccy As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "Kernel32" (ccy As Currency) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32.dll" (Guid As GUID_TYPE) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (Guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "Kernel32" (ccy As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "Kernel32" (ccy As Currency) As Long
#End If

Public g_stackDepth As Long

Public Sub TraceLog(ByRef strMessage As String)
    Dim ccyStart As Currency
    QueryPerformanceCounter ccyStart
    strMessage = ccyStart & "|" & strMessage
    'Debug.Print strMessage
    Dim iFileNum As Integer
    iFileNum = FreeFile()
    Open FileName For Append Access Write As iFileNum
    Print #iFileNum, strMessage
    Close #iFileNum
End Sub

Private Property Get FileName()
    FileName = ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".VBA.Trace"
End Property

Function CreateGuidString() As String
    Dim Guid As GUID_TYPE
    Dim strGUID As String
    Dim retValue As LongPtr
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    retValue = CoCreateGuid(Guid)
    If retValue = 0 Then
        strGUID = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(Guid, StrPtr(strGUID), guidLength)
        If retValue = guidLength Then
        ' valid GUID as a string
            CreateGuidString = strGUID
        End If
    End If
End Function

Public Sub ResetStack()
    g_stackDepth = 1
End Sub

Public Sub BootstrapInstrumentation()
    Dim val As Variant
    Dim colOut As Collection
    On Error Resume Next
    val = ThisWorkbook.Names(cstrInstrumentationMarker)
    On Error GoTo 0
    If val <> Empty Then
        ' Name is already defined - attempt to remove Instrumentation
        If MsgBox("Instrumentation already created for this workbook; attempt to remove?", vbYesNo, "Already Instrumented?") = vbYes Then
            Set colOut = InstrumentProject(False)
            ThisWorkbook.Names(cstrInstrumentationMarker).Delete
            Call MsgBox("Processed " & colOut.count & " items", vbOKOnly, "Code modifcation completed")
        End If
        Exit Sub
    End If
    ' Get list of functions
    Set colOut = InstrumentProject(True)
    ' Tag work book project as instrumented
    Call ThisWorkbook.Names.Add(Name:=cstrInstrumentationMarker, RefersToR1C1:="=TODAY()")
    Call MsgBox("Processed " & colOut.count & " items", vbOKOnly, "Code modifcation completed")
End Sub

Private Function InstrumentProject(ByVal bAddInstrumentation As Boolean) As Collection
    Dim oXLApp As Excel.Application
    Dim colFunctions As Collection
    Dim VBAEditor As VBIDE.VBE
    Dim oProject As VBIDE.VBProject
    Dim oComponent As VBIDE.VBComponent
    Dim oCode As VBIDE.CodeModule
    Dim iLine As Long
    Dim strProcName As String, strDeclaration As String, strLineToInsert As String
    Dim strObjName As String, strGUID As String
    Dim pk As vbext_ProcKind

    Set VBAEditor = Application.VBE
    Set colFunctions = New Collection
        
    'For Each oProject In VBAEditor.VBProjects
    Set oProject = VBAEditor.ActiveVBProject
    
        For Each oComponent In oProject.VBComponents
            Set oCode = oComponent.CodeModule
            ' Protect this module, and class
            If oComponent.Name <> cstrThisModuleName And oComponent.Name <> cstrInstObjName Then ' Skip module
                iLine = 1
                Do While iLine < oCode.CountOfLines
                    strProcName = oCode.ProcOfLine(iLine, pk)
                    If strProcName <> "" Then
                        ' Found a procedure
                        iLine = oCode.ProcBodyLine(strProcName, pk) ' Whitespace before proc counts as part of proc so move to first line of body (declaration)
                        strDeclaration = oCode.Lines(iLine, 1)
                        ' Deal with continuation characters
                        While (Right(strDeclaration, 2) = " _")
                            iLine = iLine + 1
                            strDeclaration = Left(strDeclaration, Len(strDeclaration) - 2)
                            strDeclaration = strDeclaration & LTrim(oCode.Lines(iLine, 1))
                        Wend
                        Call colFunctions.Add(oComponent.Name & ": " & strDeclaration)
                        If bAddInstrumentation Then
                            ' Re-use strDeclaration
                            strGUID = CreateGuidString()
                            strObjName = "obj" & Mid(Replace(strGUID, "-", ""), 2, 32)
                            strDeclaration = "[" & ModuleType(oComponent.Type) & "]" & oComponent.Name & "::[" & MethodType(pk) & "]" & strProcName
                            strLineToInsert = Replace(Replace(cstrInstrumentationCode, "%1", strObjName), "%2", strDeclaration)
                            Call oCode.InsertLines(iLine + 1, cstrTraceMarker & strGUID)
                            Call oCode.InsertLines(iLine + 2, strLineToInsert)

                        Else
                            ' Try to remove
                            strDeclaration = oCode.Lines(iLine + 1, 1)
                            If oCode.CountOfLines > iLine + 2 Then
                                If Left(strDeclaration, Len(cstrTraceMarker)) = cstrTraceMarker Then
                                    strObjName = Mid(Replace(Mid(strDeclaration, Len(cstrTraceMarker) + 1), "-", ""), 2, 32)
                                    ' Confirm this is found on next line
                                    If InStr(1, oCode.Lines(iLine + 2, 1), strObjName) <> 0 Then
                                        Call oCode.DeleteLines(iLine + 1, 2)
                                    End If
                                End If
                            End If
                        End If
                        iLine = iLine + oCode.ProcCountLines(strProcName, pk)
                    Else
                        iLine = iLine + 1
                    End If
                Loop
                Set oCode = Nothing
                Set oComponent = Nothing
            End If
        Next
    Set InstrumentProject = colFunctions
End Function

Private Function ModuleType(ByVal t As vbext_ComponentType) As String
    If t = vbext_ct_ActiveXDesigner Then
        ModuleType = "ActiveX"
    ElseIf t = vbext_ct_ClassModule Then
        ModuleType = "Class"
    ElseIf t = vbext_ct_Document Then
        ModuleType = "Doc"
    ElseIf t = vbext_ct_MSForm Then
        ModuleType = "Form"
    ElseIf t = vbext_ct_StdModule Then
        ModuleType = "Module"
    Else
    End If
End Function

Private Function MethodType(ByVal t As vbext_ProcKind) As String
    If t = vbext_pk_Get Then
        MethodType = "Get"
    ElseIf t = vbext_pk_Let Then
        MethodType = "Let"
    ElseIf t = vbext_pk_Proc Then
        MethodType = "Proc"
    ElseIf t = vbext_pk_Set Then
        MethodType = "Set"
    Else
        MethodType = "Unknown"
    End If
End Function

