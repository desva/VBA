Attribute VB_Name = "modMemoryUtil"
Option Explicit

' Diagnostics for Excel processes
' Modules loaded (and their type)
' Working set usage
' Process memory
' Fragmentation
' Handles

' (c) 2017 Dr. D. Azzopardi

' Requires reference to
' Microsoft scripting runtime (for dictionary)
' Renders results to a sheet called Report (cstrReportSheet)

' 32-bit only for now

Private Const ModuleVersion As String = "1.0.2"

Private Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
   dwTotalPhys As Long
   dwAvailPhys As Long
   dwTotalPageFile As Long
   dwAvailPageFile As Long
   dwTotalVirtual As Long
   dwAvailVirtual As Long
End Type

Private Type SYSTEM_INFO
   dwOemID As Long
   dwPageSize As Long
   lpMinimumApplicationAddress As Long
   lpMaximumApplicationAddress As Long
   dwActiveProcessorMask As Long
   dwNumberOrfProcessors As Long
   dwProcessorType As Long
   dwAllocationGranularity As Long
   dwReserved As Long
End Type

Private Type MODULEINFO
   lpBaseOfDll                   As Long
   SizeOfImage                   As Long
   EntryPoint                    As Long
End Type

Private Type PROCESS_MEMORY_COUNTERS
  cb As Long
  PageFaultCount As Long
  PeakWorkingSetSize As Long
  WorkingSetSize As Long
  QuotaPeakPagedPoolUsage As Long
  QuotaPagedPoolUsage As Long
  QuotaPeakNonPagedPoolUsage As Long
  QuotaNonPagedPoolUsage As Long
  PagefileUsage As Long
  PeakPagefileUsage As Long
End Type
 
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    Type As Long
End Type
 
' psapi functions:
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function GetModuleInformation Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, LPMODULEINFO As MODULEINFO, cb As Long) As Boolean

' psapi - enumerate working set pages
Private Declare Function QueryWorkingSet Lib "psapi.dll" (ByVal hProcess As Long, ByRef out As Any, ByVal cb As Long) As Boolean
Private Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Boolean

' Global (machine level) information
Private Declare Function GetSystemInfo Lib "kernel32" (ByRef SI As SYSTEM_INFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

' Process level information
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal strLibrary As String, ByVal hFile As Long, ByVal flags As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hLib As Long, ByVal strName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CloseLibrary Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function VirtualQuery Lib "kernel32" (ByVal p As Long, i As MEMORY_BASIC_INFORMATION, ByVal s As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hLib As Long, ByVal lpPath As String, ByVal cb As Long) As Long
Private Declare Function GetProcessHandleCount Lib "kernel32" (ByVal hProcess As Long, ByRef pdwHandleCount As Long) As Long

' Handles (user32 resource)
Private Declare Function GetGuiResources Lib "user32" (ByVal hProcess As Long, ByVal nType As Long) As Long
 
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const DONT_RESOLVE_DLL_REFERENCES = &H1
' Change this as required
Private Const cstrReportSheet As String = "Report"
 
Private m_arrModules() As Variant   ' Module information
Private m_dctModSizes As Dictionary ' Module sizes
Private m_dctOutput As Dictionary   ' Output
Private m_dctMemMap As Dictionary   ' Process space memory map
 
Public Sub GetProcessInfo()
    Dim sht As Worksheet
    ' Make sure report sheet exists
    If Not Application.Evaluate("IsRef('" & cstrReportSheet & "'!A1)") Then
        Set sht = ThisWorkbook.Sheets.Add
        sht.Name = cstrReportSheet
    End If
    
    Application.ScreenUpdating = False
    Call ResetReport
    Set m_dctOutput = New Dictionary
    Call EnumerateModules
    Call GetProcInfoEx
    Call RenderMemoryMap
    Call RenderModules
    Call RenderProcessOutput
errHandler:
    Application.ScreenUpdating = True
    Set m_dctMemMap = Nothing
    Set m_dctModSizes = Nothing
    Set m_dctOutput = Nothing
    Erase m_arrModules
End Sub
 
Public Sub ResetReport()
    Dim tgt As Range
    Set tgt = ThisWorkbook.Sheets(cstrReportSheet).Range("B2")
    Set tgt = Range(tgt, tgt.SpecialCells(xlLastCell))
    Call tgt.EntireRow.Delete(Shift:=xlUp)
End Sub
 
Private Sub EnumerateModules()
    Const MAX_PATH As Long = 260
    Dim lModules() As Long, lRet As Long, hProcess As Long, i As Long, hLib As Long
    Dim sName As String, sPath As String
    Dim mi As MODULEINFO
    hProcess = GetCurrentProcess ' = -1
    Dim k As Currency
    Dim bXll As Boolean, bCom As Boolean
    Dim tgt As Range
    Dim vOut As Variant
    Set m_dctModSizes = New Dictionary
    If hProcess Then
        ReDim lModules(1023)
        If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
            'Image base : Image Size|Name|isXLL|isCOM|Entry point
            ReDim m_arrModules(0 To lRet \ 4, 0 To 1)
            For i = 0 To lRet \ 4
                sName = String$(MAX_PATH, vbNullChar)
                Call GetModuleBaseName(hProcess, lModules(i), sName, MAX_PATH)
                sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                hLib = LoadLibraryEx(sName, 0, DONT_RESOLVE_DLL_REFERENCES)
                sPath = String$(MAX_PATH, vbNullChar)
                Call GetModuleFileName(hLib, sPath, MAX_PATH)
                bXll = (0 <> GetProcAddress(hLib, "xlAutoOpen"))
                bCom = (0 <> GetProcAddress(hLib, "DllGetClassObject"))
                If GetModuleInformation(hProcess, hLib, mi, Len(mi)) Then
                    k = ConvLong(mi.lpBaseOfDll)
                    If Not m_dctModSizes.Exists(k) Then
                        Call m_dctModSizes.Add(k, mi.SizeOfImage)
                        m_arrModules(i, 0) = k
                        m_arrModules(i, 1) = mi.SizeOfImage & "|" & _
                                             sPath & "|" & _
                                             CStr(bXll) & "|" & _
                                             CStr(bCom) & "|" & _
                                             mi.EntryPoint & "|" & _
                                             sName
                    Else
                        m_arrModules(i, 0) = k
                        m_arrModules(i, 1) = "0|" & _
                                             sPath & "|" & _
                                             CStr(bXll) & "|" & _
                                             CStr(bCom) & "|" & _
                                             mi.EntryPoint & "|" & _
                                             sName
                    End If
                End If
'                Call CloseHandle(hLib)
            Next
        End If
    End If
    Call CloseHandle(hProcess)
    Call HeapSort(m_arrModules)
End Sub

' searches array for module that is intersected by addr
Private Function FindModule(ByVal addr As Currency) As String
    FindModule = ""
    Dim l As Long, r As Long, s As Long, m As Long
    Dim v As Currency
    l = LBound(m_arrModules, 1)
    r = UBound(m_arrModules, 1)
    ' check within entire range
    If m_arrModules(l, 0) > addr Or m_arrModules(r, 0) < addr Then Exit Function
    Dim i As Long
    i = 0
    Do
        m = (l + r) \ 2
        v = m_arrModules(m, 0)
        If addr < v Then
            r = m - 1
        Else
            l = m + 1
            s = m_dctModSizes(v)
            If v + s > addr Then
                'addr is within interval
                FindModule = Split(m_arrModules(m, 1), "|")(5)
                Exit Function
            End If
            If l <> r And m_arrModules(m + 1, 0) > addr Then
                ' next Module doesn't contain addr, so quit
                Exit Function
            End If
        End If
    Loop While (l <= r)
End Function


Private Sub GetProcInfoEx()
    Dim MS As MEMORYSTATUS
    Dim SI As SYSTEM_INFO
    Dim PMC As PROCESS_MEMORY_COUNTERS
    Dim nProcHdl As Long
    
    MS.dwLength = Len(MS)
    ' Minimize WS size to attempt repeatability
    Call EmptyWorkingSet(GetCurrentProcess)
    Call GlobalMemoryStatus(MS)
    Call GetSystemInfo(SI)

    Call GetProcessMemoryInfo(GetCurrentProcess, PMC, Len(PMC))
    Call GetProcessHandleCount(GetCurrentProcess, nProcHdl)
    
    With m_dctOutput
        .Add "Memory Load", Format(MS.dwMemoryLoad / 100, "##.##%")
        .Add "Total Physical/Kb", Format(Fix(ConvLong(MS.dwTotalPhys) / 1024), "###,###")
        .Add "Available Physical/Kb", Format(Fix(ConvLong(MS.dwAvailPhys) / 1024), "###,###")
        .Add "Total Virtual/Kb", Format(Fix(ConvLong(MS.dwTotalVirtual) / 1024), "###,###")
        .Add "Available Virtual/Kb", Format(Fix(ConvLong(MS.dwAvailVirtual) / 1024), "###,###")
        .Add "Process VM Space/Kb", Format(Fix((ConvLong(SI.lpMaximumApplicationAddress) - ConvLong(SI.lpMinimumApplicationAddress) + 1) / 1024), "###,###")
        .Add "Page file usage/Kb", Format(Fix(ConvLong(PMC.PagefileUsage) / 1024), "###,###")
        .Add "Working set size/Kb", Format(Fix(ConvLong(PMC.WorkingSetSize) / 1024), "###,###")
        .Add "Page fault count", Format(Fix(ConvLong(PMC.PageFaultCount)), "###,###")
        .Add "GDI Objects", Format(Fix(GetGuiResources(GetCurrentProcess(), 0)), "###,###")
        .Add "USER Objects", Format(Fix(GetGuiResources(GetCurrentProcess(), 1)), "###,###")
        .Add "Handles", Format(nProcHdl, "###,###")
        .Add "Page Size/Bytes", SI.dwPageSize
    End With
    
    ' Enumerate Working set pages
    Dim nReq As Long, i As Long
    
    nReq = 4 * (1 + (PMC.WorkingSetSize / SI.dwPageSize))
    Dim pWsInfo() As Long
    ReDim pWsInfo(nReq)
    i = QueryWorkingSet(GetCurrentProcess, pWsInfo(0), 4 * nReq)
    If (i <> 0) Then
        Dim pgInfo As Long
        Dim nShared As Long, nCoW As Long, nExec As Long, nRW As Long, nRO As Long
        For i = 0 To pWsInfo(0)
            If (pWsInfo(i) < &H70000000) Then
                pgInfo = pWsInfo(i) And &HFFF
                If (pgInfo And &H100) Then nShared = nShared + 1
                If (pgInfo And &H5) Then nCoW = nCoW + 1
                If (pgInfo And &H2) Then nExec = nExec + 1
                If ((pgInfo And &H4) And (0 = (pgInfo And &H1))) Then nRW = nRW + 1
                If ((pgInfo And &H1) And (0 = (pgInfo And &H4))) Then nRO = nRO + 1
            End If
        Next
        Call m_dctOutput.Add("Shared Pages", Format(Fix(nShared), "###,###"))
        Call m_dctOutput.Add("Copy-on-write Pages", Format(Fix(nCoW), "###,###"))
        Call m_dctOutput.Add("Executable Pages", Format(Fix(nExec), "###,###"))
        Call m_dctOutput.Add("Read-only Pages", Format(Fix(nRO), "###,###"))
        Call m_dctOutput.Add("Read-write pages", Format(Fix(nRW), "###,###"))
    End If
    
    ' Create a map of the entire process space
    Dim ii As Currency
    ii = ConvLong(SI.lpMinimumApplicationAddress)
    Set m_dctMemMap = New Dictionary
    Dim MBI As MEMORY_BASIC_INFORMATION
    Do
        Call VirtualQuery(ii, MBI, Len(MBI))
        Call m_dctMemMap.Add(ii, MBI.AllocationProtect + MBI.State + MBI.Protect + MBI.Type)
        ii = ii + MBI.RegionSize
    Loop While (ii <= ConvLong(SI.lpMaximumApplicationAddress))
    If Not m_dctMemMap.Exists(ii) Then
        Call m_dctMemMap.Add(ii, 0)
    End If
End Sub

Private Sub RenderProcessOutput()
    Dim v As Variant, vOut As Variant
    Dim i As Long
    i = 1
    Dim tgt As Range
    Set tgt = ThisWorkbook.Sheets(cstrReportSheet).Range("B2").Resize(1 + m_dctOutput.Count, 2)
    ReDim vOut(0 To m_dctOutput.Count, 0 To 1)
    vOut(0, 0) = "Metric"
    vOut(0, 1) = "Value"
    For Each v In m_dctOutput
        vOut(i, 0) = v
        vOut(i, 1) = m_dctOutput(v)
        i = i + 1
    Next
    tgt.Value = vOut
    Erase vOut
End Sub

Private Sub RenderMemoryMap()
    Dim k As Variant, k1 As Variant
    Dim nSizeKB As Long
    Dim dFreeTot As Double, dFreeMax As Double
    Dim dCommitTot As Double, dReserveTot As Double
    Dim tgt As Range
    Dim vOut As Variant
    Set tgt = ThisWorkbook.Sheets(cstrReportSheet).Range("L2").Resize(m_dctMemMap.Count + 1, 5)
    Dim i As Long
    k1 = m_dctMemMap.Keys
    Dim v As Long
    Dim strImage As String
    Dim strState As String
    Dim strType As String
    Dim strAccess As String
    i = 0
    dFreeTot = 0
    dFreeMax = 0
    dCommitTot = 0
    dReserveTot = 0
    ReDim vOut(0 To m_dctMemMap.Count, 0 To 4)
    vOut(0, 0) = "Address"
    vOut(0, 1) = "Size"
    vOut(0, 2) = "State"
    vOut(0, 3) = "Type"
    vOut(0, 4) = "Access"
    Dim strPrevImg As String
    For Each k In m_dctMemMap
        v = m_dctMemMap(k)
        nSizeKB = (k1(i + 1) - k) / 1024
        If (v And &H10000) Then
            strState = "FREE"
            dFreeTot = dFreeTot + nSizeKB
            If dFreeMax < nSizeKB Then dFreeMax = nSizeKB
        ElseIf (v And &H1000) Then
            strState = "COMMITTED"
            dCommitTot = dCommitTot + nSizeKB
        ElseIf (v And &H2000) Then
            strState = "RESERVED"
            dReserveTot = dReserveTot + nSizeKB
        Else
            strState = ""
        End If
        If (v And &H1000000) Then
            strType = "IMAGE"
        ElseIf (v And &H40000) Then
            strType = "MAPPED"
        ElseIf (v And &H20000) Then
            strType = "PRIVATE"
        Else
            strType = ""
        End If
        strAccess = "----"
        If (v And &H10) Then
            strAccess = "X---"
        ElseIf (v And &H20) Then strAccess = "XR--"
        ElseIf (v And &H40) Then strAccess = "XRW-"
        ElseIf (v And &H80) Then strAccess = "XRWC"
        ElseIf (v And &H1) Then strAccess = "-R--"
        ElseIf (v And &H2) Then strAccess = "-RW-"
        ElseIf (v And &H4) Then strAccess = "-R--"
        ElseIf (v And &H8) Then strAccess = "-RW-"
        End If
        If (v And &H100) Then
            strAccess = strAccess & "(GD)"
        ElseIf (v And &H200) Then
            strAccess = strAccess & "(NC)"
        End If
        strImage = FindModule(k)
        If (strImage <> "") Then
            If (strImage = strPrevImg) Then
                strAccess = strAccess & " (" & strImage & " cont.)"
            Else
                strAccess = strAccess & " (" & strImage & ")"
            End If
            strPrevImg = strImage
        Else
            strPrevImg = ""
        End If
        vOut(i + 1, 0) = "0x" & Hex(k)
        vOut(i + 1, 1) = Format(nSizeKB, "###,###")
        vOut(i + 1, 2) = strState
        vOut(i + 1, 3) = strType
        vOut(i + 1, 4) = strAccess
        i = i + 1
        'If (i > 1700) Then Debug.Assert 0
        If (i = UBound(k1)) Then Exit For
    Next
    tgt.Value = vOut
    Erase vOut
    m_dctOutput("Memory fragmentation") = Format(1# - dFreeMax / dFreeTot, "##.##%")
    m_dctOutput("Largest free block/K") = Format(dFreeMax, "###,###")
    m_dctOutput("Total free/K") = Format(dFreeTot, "###,###")
    m_dctOutput("Total commited/K") = Format(dCommitTot, "###,###")
    m_dctOutput("Total reserved/K") = Format(dReserveTot, "###,###")
End Sub

Private Sub RenderModules()
    Dim tgt As Range
    Dim n As Long, i As Long, j As Long
    n = UBound(m_arrModules, 1) - LBound(m_arrModules, 1) + 1
    Set tgt = ThisWorkbook.Sheets(cstrReportSheet).Range("F2").Resize(n + 1, 4)
    Dim vOut As Variant
    Dim vLine As Variant
    ReDim vOut(0 To n, 0 To 3)
    vOut(0, 0) = "Name"
    vOut(0, 1) = "Loaded at"
    vOut(0, 2) = "Size"
    vOut(0, 3) = "Type"
    Dim strType As String
    j = 1
    For i = 0 To n - 1
        vLine = Split(m_arrModules(i, 1), "|")
        If (CLng(vLine(0)) > 0) Then
            vOut(j, 0) = vLine(1)
            vOut(j, 1) = "0x" & Hex(m_arrModules(i, 0))
            vOut(j, 2) = "0x" & Hex(vLine(0))
            strType = ""
            If (vLine(2)) Then strType = "Xll "
            If (vLine(3)) Then strType = strType & " COM"
            vOut(j, 3) = strType
            j = j + 1
        End If
    Next
    tgt.Value = vOut
    Erase vOut
End Sub

Private Function ConvLong(ByVal lng As Long) As Currency
    If (lng >= 0) Then
        ConvLong = CCur(lng)
    Else
        ConvLong = 2 ^ 31 + lng + 1
    End If
End Function

' Sort an Nx2 array based on first column
Public Sub HeapSort(ByRef arr() As Variant)
    Dim b As Long, n As Long, bb As Long
    b = LBound(arr, 1) ' count from 0 or 1?
    bb = LBound(arr, 2)
    n = UBound(arr, 1) - b + 1 ' Number of items
    If n = 0 Then Exit Sub ' Nothing to do
    Dim k As Variant, v As Variant
    Dim p As Long, i As Long, c As Long
    p = n / 2
    Do While (True)
        If (p > b) Then
            ' first stage - Sorting the heap
            p = p - 1
            k = arr(p, bb)
            v = arr(p, bb + 1)
        Else
            ' second stage - Extracting elements in-place
            n = n - 1
            If n = 0 Then Exit Do
            k = arr(n, bb)
            v = arr(n, bb + 1)
            arr(n, bb) = arr(b, bb)
            arr(n, bb + 1) = arr(b, bb + 1)
        End If
        ' insert operation - pushing t down the heap to replace the parent
        i = p ' start at the parent index
        c = i * 2 + 1 ' Get its left child index
        Do While (c < n)
            ' choose the largest child
            If (c + 1 < n) Then
                If arr(c + 1, bb) > arr(c, bb) Then c = c + 1 ' right child exists and is bigger
            End If
            ' is the largest child larger than the entry?
            If (arr(c, bb) > k) Then
                arr(i, bb) = arr(c, bb) ' overwrite entry with child
                arr(i, bb + 1) = arr(c, bb + 1)
                i = c ' move index to the child
                c = i * 2 + 1 ' get the left child and go around again
            Else
                Exit Do ' t's place is found
            End If
        Loop
        ' store the temporary value at its new location
        arr(i, bb) = k
        arr(i, bb + 1) = v
    Loop
End Sub
