Attribute VB_Name = "mdlMemory"
'Credits goes to Bagus Judistira

Option Explicit

Dim pdhStatus As PDH_STATUS
Dim Counters(0 To 99) As CounterInfo
Dim hQuery As Long
Private QueryObject As Object

Public Function GetMemory(ProcessID As Long) As String
    
    On Error Resume Next
    Dim byteSize As Double, hProcess As Long, ProcMem As PROCESS_MEMORY_COUNTERS
    
    ProcMem.cb = LenB(ProcMem)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
    
    If hProcess <= 0 Then GetMemory = "N/A": Exit Function
    
    GetProcessMemoryInfo hProcess, ProcMem, ProcMem.cb
    
    byteSize = ProcMem.WorkingSetSize
    GetMemory = byteSize
    
    Call CloseHandle(hProcess)
    
End Function

Sub MemoryInfo(lbPhysMem As Label, lbAvaiPhyMem As Label, lbUsedPhyMem As Label, _
                                lbMemLoad As Frame, lbPagFile As Label, lbAvaiPagFile As Label, _
                                lbPagFileUsg As Label, lbAvaiVirMem As Label, lbUsedVirMem As Label, _
                                lbVirMem As Label, Prog As ProgressBar, MemUs As StatusBar)

    Dim mem As MEMORYSTATUS
    
    mem.dwLength = Len(mem)
    
    GlobalMemoryStatus mem
    
    lbPhysMem.Caption = "Total Physical Memory : " & Format(mem.dwTotalPhys \ 1024, "") & " KB"
    lbAvaiPhyMem.Caption = "Available Physical Memory : " & Format(mem.dwAvailPhys \ 1024, "") & " KB"
    lbUsedPhyMem.Caption = "Used Physical Memory : " & Format(mem.dwTotalPhys \ 1024 - mem.dwAvailPhys \ 1024, "") & " KB"
    lbMemLoad.Caption = "Memory Used : " & mem.dwMemoryLoad & " %"
    lbPagFile.Caption = "Paging File : " & Format(mem.dwTotalPageFile \ 1024, "") & " KB"
    lbAvaiPagFile.Caption = "Available Paging File : " & Format(mem.dwAvailPageFile \ 1024, "") & " KB"
    lbPagFileUsg.Caption = "Paging File Usage : " & Format(mem.dwTotalPageFile \ 1024 - mem.dwAvailPageFile \ 1024, "") & " KB"
    lbAvaiVirMem.Caption = "Available Virtual Memory : " & Format(mem.dwAvailVirtual \ 1024, "") & " KB"
    lbUsedVirMem.Caption = "Used Virtual Memory : " & Format(mem.dwTotalVirtual \ 1024 - mem.dwAvailVirtual \ 1024, "") & " KB"
    lbVirMem.Caption = "Virtual Memory : " & Format(mem.dwTotalVirtual \ 1024, "") & " KB"
    frmScanVirus.ProgMemUsed.value = mem.dwMemoryLoad
    'MemUs.Panels(3).Text = "Memory Usage : " & mem.dwMemoryLoad & " %"
End Sub

Sub GetCPUInfo(lCPU As Label, lCPUStat As ProgressBar, CPUusg As StatusBar)
    
    pdhStatus = PdhOpenQuery(0, 1, hQuery)
    
    AddCounter "\Processor(0)\% Processor Time", hQuery
    UpdateValues lCPU, lCPUStat, CPUusg
    
End Sub

Sub UpdateValues(lCPU As Label, lCPUStat As ProgressBar, CPUusg As StatusBar)
    
    Dim dblCounterValue As Double
    Dim pdhStatus As Long
    Dim strInfo As String
    Dim I As Long
        
    PdhCollectQueryData (hQuery)
    
    I = 0
    dblCounterValue = PdhVbGetDoubleCounterValue( _
        Counters(I).hCounter, pdhStatus)
    
    If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus _
        = PDH_CSTATUS_NEW_DATA) Then
        lCPU.Caption = Format(dblCounterValue, "0") & " %"
        lCPUStat.value = dblCounterValue
        CPUusg.Panels(3).Text = "CPU Usage : " & lCPU.Caption
    End If
        
End Sub

Sub AddCounter(strCounterName As String, _
    hQuery As Long)
    
    Dim pdhStatus As PDH_STATUS
    Dim hCounter As Long, currentCounterIdx As Long
    
    pdhStatus = PdhVbAddCounter(hQuery, strCounterName, _
        hCounter)
    Counters(currentCounterIdx).hCounter = hCounter
    Counters(currentCounterIdx).strName = strCounterName
    currentCounterIdx = currentCounterIdx + 1
    
End Sub



