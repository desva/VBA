Attribute VB_Name = "modMemoryUtil"
Option Explicit

' Diagnostics for Excel processes
' Modules loaded (and their type)
' Working set usage

' Requires reference to
' Microsoft scripting runtime (for dictionary)

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
 
' Handles (user32 resource)
Private Declare Function GetGuiResources Lib "user32" (ByVal hProcess As Long, ByVal nType As Long) As Long
 
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const DONT_RESOLVE_DLL_REFERENCES = &H1
 
Private Sub EnumerateModules(ByVal tgt As Range)
    Const MAX_PATH As Long = 260
    Dim lModules() As Long, N As Long, lRet As Long, hProcess As Long, i As Long, hLib As Long
    Dim sName As String
    Dim mi As MODULEINFO
    hProcess = GetCurrentProcess ' = -1
    Dim strItem As String
    If hProcess Then
        ReDim lModules(1023)
        If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
            tgt.Resize(1, 6).Value = Array("Name", "XLL", "COM", "Entry point", "Image Base", "Image Size")
            For i = 0 To lRet \ 4
                sName = String$(MAX_PATH, vbNullChar)
                GetModuleBaseName hProcess, lModules(i), sName, MAX_PATH
                sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                hLib = LoadLibraryEx(sName, 0, DONT_RESOLVE_DLL_REFERENCES)
                strItem = sName & ";"
                If (0 <> GetProcAddress(hLib, "xlAutoOpen")) Then
                    strItem = strItem & "[Xll];"
                Else
                    strItem = strItem & ";"
                End If
                If (0 <> GetProcAddress(hLib, "DllGetClassObject")) Then
                    strItem = strItem & "[COM];"
                Else
                    strItem = strItem & ";"
                End If
                If GetModuleInformation(hProcess, hLib, mi, Len(mi)) Then
                    strItem = strItem & "0x" & Hex(mi.EntryPoint) & ";0x" & Hex(mi.lpBaseOfDll) & ";0x" & Hex(mi.SizeOfImage) & ";"
                Else
                    strItem = strItem & ";;;"
                End If
                tgt.Offset(i + 1, 0) = strItem
            Next
            Call tgt.Offset(1).Resize(i, 1).TextToColumns(Semicolon:=True)
            Call tgt.Resize(1, 6).EntireColumn.AutoFit
        End If
    End If
    CloseHandle hProcess
    
End Sub

Public Sub GetProcessInfo()
    Dim dct As Dictionary
    Set dct = New Dictionary
    Call GetProcInfoEx(dct)
    Call RenderOutput(dct)
End Sub

Private Sub GetProcInfoEx(ByRef dctIn As Dictionary)
    Dim MS As MEMORYSTATUS
    Dim SI As SYSTEM_INFO
    Dim PMC As PROCESS_MEMORY_COUNTERS

    MS.dwLength = Len(MS)
    Call EmptyWorkingSet(GetCurrentProcess)
    Call GlobalMemoryStatus(MS)
    Call GetSystemInfo(SI)

    Call GetProcessMemoryInfo(GetCurrentProcess, PMC, Len(PMC))
    
    Call dctIn.Add("Memory Load/%", MS.dwMemoryLoad)
    Call dctIn.Add("Total Physical/K", Format(Fix(MS.dwTotalPhys / 1024), "###,###"))
    Call dctIn.Add("Available Physical/K", Format(Fix(MS.dwAvailPhys / 1024), "###,###"))
    Call dctIn.Add("Total Virtual/K", Format(Fix(MS.dwTotalVirtual / 1024), "###,###"))
    Call dctIn.Add("Available Virtual/K", Format(Fix(MS.dwAvailVirtual / 1024), "###,###"))
    Call dctIn.Add("Process VM Space/K", Format(Fix((SI.lpMaximumApplicationAddress - SI.lpMinimumApplicationAddress + 1) / 1024), "###,###"))
    Call dctIn.Add("Page file usage/K", Format(Fix(PMC.PagefileUsage / 1024), "###,###"))
    Call dctIn.Add("Working set size/K", Format(Fix(PMC.WorkingSetSize / 1024), "###,###"))
    Call dctIn.Add("Page fault count", Format(Fix(PMC.PageFaultCount), "###,###"))
    Call dctIn.Add("GDI Objects", Format(Fix(GetGuiResources(GetCurrentProcess(), 0)), "###,###"))
    Call dctIn.Add("USER Objects", Format(Fix(GetGuiResources(GetCurrentProcess(), 1)), "###,###"))
    Call dctIn.Add("Page Size", SI.dwPageSize)

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
        Call dctIn.Add("Shared Pages", Format(Fix(nShared), "###,###"))
        Call dctIn.Add("Copy-on-write Pages", Format(Fix(nCoW), "###,###"))
        Call dctIn.Add("Executable Pages", Format(Fix(nExec), "###,###"))
        Call dctIn.Add("Read-only Pages", Format(Fix(nRO), "###,###"))
        Call dctIn.Add("Read-write pages", Format(Fix(nRW), "###,###"))
    End If
    ' Attempt to create a map of the entire process space
    i = SI.lpMinimumApplicationAddress
    Dim dctMemMap As Dictionary
    Set dctMemMap = New Dictionary
    Dim MBI As MEMORY_BASIC_INFORMATION
    Do
        Call VirtualQuery(i, MBI, Len(MBI))
        Call dctMemMap.Add(i, MBI.AllocationProtect + MBI.State + MBI.Protect + MBI.Type)
        i = i + MBI.RegionSize
    Loop While (i <= SI.lpMaximumApplicationAddress)
    Call dctMemMap.Add(SI.lpMaximumApplicationAddress, 0)
    Set dctIn("Memory map") = dctMemMap
End Sub

Private Sub RenderOutput(ByRef dct As Dictionary)
    Call EnumerateModules(ThisWorkbook.Names("Modules").RefersToRange)
    Dim v As Variant
    Dim i As Long
    i = 1
    Dim tgt As Range
    Set tgt = ThisWorkbook.Names("ProcInfo").RefersToRange
    For Each v In dct
        If v = "Memory map" Then
            Call RenderMemoryMap(dct(v))
        Else
            tgt.Offset(i, 0) = v
            tgt.Offset(i, 1) = dct(v)
            i = i + 1
        End If
    Next
End Sub

Private Sub RenderMemoryMap(ByRef d As Dictionary)
    Dim k As Variant, k1 As Variant
    Dim tgt As Range
    Set tgt = ThisWorkbook.Names("VMMAP").RefersToRange
    Dim i As Long
    k1 = d.Keys
    For Each k In d
        tgt.Offset(i).Resize(1, 3) = Array("0x" & Hex(k), (k1(i + 1) - k) / 1024, "0x" & Hex(d(k)))
        i = i + 1
    Next
End Sub
