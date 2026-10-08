Attribute VB_Name = "modResources"

Public Enum ResourceID
    resNone
    resBinoculars
    resDelete
    resDownload
    resError
    resFolder
    resInfo
    resLogHistory
    resLogHistoryPanelIcon
    resMessageBox
    resQuestion
    resRevert
    resRevertPanelIcon
    resSuccess
    resWarning
    resFolderClosed
    resFolderOpen
    resBlackFolder
End Enum

Private mImages As New Dictionary(Of ResourceID, LongPtr)
Private mStdPics As New Dictionary(Of ResourceID, StdPicture)

Public Sub LoadImages()
    On Error Resume Next
    StartupTrace "LoadImages enter"
    If Not GDIPlus_IsStarted() Then GDIPlus_EnsureStarted
    If Err.Number <> 0 Then StartupTraceError "LoadImages GDIPlus_EnsureStarted"
    Set mImages = New Dictionary(Of ResourceID, LongPtr)
    If Err.Number <> 0 Then StartupTraceError "LoadImages new image dictionary"
    Set mStdPics = New Dictionary(Of ResourceID, StdPicture)
    If Err.Number <> 0 Then StartupTraceError "LoadImages new picture dictionary"
    LoadImagesFromResources
    If Err.Number <> 0 Then StartupTraceError "LoadImages LoadImagesFromResources"
    StartupTrace "LoadImages leave"
    Err.Clear
End Sub

Private Sub AddImage(id As ResourceID, pic As LongPtr)
    mImages.Add id, pic
End Sub

Private Function ResourceFileName(id As ResourceID) As String
    Select Case id
        Case resBinoculars: ResourceFileName = "binoculars.png"
        Case resDelete: ResourceFileName = "delete.png"
        Case resDownload: ResourceFileName = "download.png"
        Case resError: ResourceFileName = "error.png"
        Case resFolder: ResourceFileName = "folder.png"
        Case resInfo: ResourceFileName = "info.png"
        Case resLogHistory: ResourceFileName = "logHistory.png"
        Case resLogHistoryPanelIcon: ResourceFileName = "logHistorypanel icon.png"
        Case resMessageBox: ResourceFileName = "messagebox.png"
        Case resQuestion: ResourceFileName = "question.png"
        Case resRevert: ResourceFileName = "revert.png"
        Case resRevertPanelIcon: ResourceFileName = "revert panel icon.png"
        Case resSuccess: ResourceFileName = "success.png"
        Case resWarning: ResourceFileName = "warning.png"
        Case resFolderClosed: ResourceFileName = "folder_closed.ico"
        Case resFolderOpen: ResourceFileName = "folder_open.ico"
        Case resBlackFolder: ResourceFileName = "black_folder_open.ico"
    End Select
End Function

Private Function ResourceFolderFromFileName(ByVal resName As String) As String
    Dim ext As String
    If InStrRev(resName, ".") = 0 Then Exit Function
    ext = LCase$(Mid$(resName, InStrRev(resName, ".") + 1))
    Select Case ext
        Case "png": ResourceFolderFromFileName = "PNG"
        Case "ico": ResourceFolderFromFileName = "ICON"
        Case "cur": ResourceFolderFromFileName = "CUR"
    End Select
End Function

Private Function ResBytesAreValid(data() As Byte) As Boolean
    On Error Resume Next
    ResBytesAreValid = False
    ResBytesAreValid = (UBound(data) >= LBound(data))
    Err.Clear
End Function

Private Function TryLoadResData(ByVal resName As String, ByVal folder As String) As Byte()
    Dim data() As Byte
    Dim baseName As String
    
    On Error Resume Next
    If Len(resName) = 0 Or Len(folder) = 0 Then
        Err.Clear
        Exit Function
    End If
    
    Err.Clear
    data = LoadResData(resName, folder)
    If Not ResBytesAreValid(data) Then
        Err.Clear
        data = LoadResData(UCase$(resName), folder)
    End If
    If Not ResBytesAreValid(data) Then
        baseName = resName
        If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
        Err.Clear
        data = LoadResData(baseName, folder)
    End If
    
    If ResBytesAreValid(data) Then TryLoadResData = data
    Err.Clear
End Function

Private Function PictureFromResBytes(data() As Byte) As StdPicture
    Dim pic As StdPicture
    
    On Error Resume Next
    If Not ResBytesAreValid(data) Then
        Err.Clear
        Exit Function
    End If
    Set pic = LoadPicture(data)
    If Not pic Is Nothing Then Set PictureFromResBytes = pic
    Err.Clear
End Function

' Stacked Deck Trainer / Pitch And Time Shifter pattern: LoadResData + LoadPicture.
' The runtime reads the project Resources tree in the IDE and the EXE resource section when compiled.
' Pending errors are cleared before exit so a missing resource cannot raise error 91 in the caller.
Public Function GetResPicture(ByVal resName As String, ByVal resFolder As String) As StdPicture
    Dim data() As Byte
    Dim pic As StdPicture
    Dim baseName As String
    Dim ext As String
    
    On Error Resume Next
    If Len(resFolder) = 0 Then resFolder = ResourceFolderFromFileName(resName)
    If Len(resName) = 0 Or Len(resFolder) = 0 Then
        Err.Clear
        Exit Function
    End If
    
    data = TryLoadResData(resName, resFolder)
    Set pic = PictureFromResBytes(data)
    
    If pic Is Nothing Then
        ext = LCase$(Mid$(resName, InStrRev(resName, ".") + 1))
        baseName = resName
        If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
        Err.Clear
        If ext = "ico" Then
            Set pic = LoadResPicture(resName, vbResIcon)
            If pic Is Nothing Then Set pic = LoadResPicture(baseName, vbResIcon)
        ElseIf ext = "cur" Then
            Set pic = LoadResPicture(resName, vbResCursor)
            If pic Is Nothing Then Set pic = LoadResPicture(baseName, vbResCursor)
        End If
    End If
    
    If Not pic Is Nothing Then
        Set GetResPicture = pic
        StartupTrace "GetResPicture " & resName & " ok"
    Else
        StartupTrace "GetResPicture " & resName & " picture is Nothing"
    End If
    If Err.Number <> 0 Then StartupTraceError "GetResPicture " & resName
    Err.Clear
End Function

Private Function LoadResBytes(ByVal resName As String) As Byte()
    Dim folder As String
    Dim data() As Byte
    
    On Error Resume Next
    folder = ResourceFolderFromFileName(resName)
    data = TryLoadResData(resName, folder)
    If ResBytesAreValid(data) Then LoadResBytes = data
    Err.Clear
End Function

Private Sub LoadOneImage(id As ResourceID, ByVal resName As String)
    Dim p As LongPtr
    
    On Error Resume Next
    StartupTrace "LoadOneImage " & resName
    p = LoadGDIPlusBitmapFromResource(resName)
    If Err.Number <> 0 Then
        StartupTraceError "LoadOneImage " & resName
    ElseIf p = 0 Then
        StartupTrace "LoadOneImage " & resName & " no bitmap"
    End If
    If p <> 0 Then
        If Not mImages.Exists(id) Then mImages.Add id, p
    End If
    Err.Clear
End Sub

Private Sub LoadImagesFromResources()
    On Error Resume Next

    LoadOneImage resBinoculars, "binoculars.png"
    LoadOneImage resDelete, "delete.png"
    LoadOneImage resDownload, "download.png"
    LoadOneImage resError, "error.png"
    LoadOneImage resFolder, "folder.png"
    LoadOneImage resInfo, "info.png"
    LoadOneImage resLogHistory, "logHistory.png"
    LoadOneImage resLogHistoryPanelIcon, "logHistorypanel icon.png"
    LoadOneImage resMessageBox, "messagebox.png"
    LoadOneImage resQuestion, "question.png"
    LoadOneImage resRevert, "revert.png"
    LoadOneImage resRevertPanelIcon, "revert panel icon.png"
    LoadOneImage resSuccess, "success.png"
    LoadOneImage resWarning, "warning.png"
    LoadOneImage resFolderClosed, "folder_closed.ico"
    LoadOneImage resFolderOpen, "folder_open.ico"
    LoadOneImage resBlackFolder, "black_folder_open.ico"
    Err.Clear
End Sub

Public Function GetImagePtr(id As ResourceID) As LongPtr
    On Error Resume Next
    If id = resNone Then
        Err.Clear
        Exit Function
    End If
    If mImages Is Nothing Then
        Err.Clear
        Exit Function
    End If
    If mImages.Exists(id) Then GetImagePtr = mImages(id)
    Err.Clear
End Function

Public Function GetImageStd(id As ResourceID) As StdPicture
    Dim pic As StdPicture
    Dim resName As String
    
    On Error Resume Next
    If id = resNone Then
        Err.Clear
        Exit Function
    End If
    If mStdPics Is Nothing Then Set mStdPics = New Dictionary(Of ResourceID, StdPicture)
    
    If mStdPics.Exists(id) Then
        Set pic = mStdPics(id)
        If Not pic Is Nothing Then Set GetImageStd = pic
        Err.Clear
        Exit Function
    End If
    
    resName = ResourceFileName(id)
    If Len(resName) > 0 Then
        Set pic = GetResPicture(resName, ResourceFolderFromFileName(resName))
    End If
    
    If pic Is Nothing Then
        If Not mImages Is Nothing Then
            If mImages.Exists(id) Then Set pic = GpBitmapToStdPicture(mImages(id))
        End If
    End If
    
    If Not pic Is Nothing Then
        mStdPics.Add id, pic
        Set GetImageStd = pic
    End If
    Err.Clear
End Function

Public Function LoadGDIPlusBitmapFromResource(resName As String) As LongPtr
    Dim data() As Byte
    Dim size As Long
    Dim hGlobal As LongPtr
    Dim pGlobal As LongPtr
    Dim pStream As IUnknown
    Dim gpBmp As LongPtr
    Dim status As Long

    ' Same bytes LoadResData returns in the IDE and in the EXE.
    ' Direct GDI+ calls, not WinDevLib.wdAPI / wdGDIP: those objects are what raised error 91 in the EXE.
    On Error Resume Next
    If Not GDIPlus_IsStarted() Then GDIPlus_EnsureStarted
    data = LoadResBytes(resName)
    If Not ResBytesAreValid(data) Then
        Err.Clear
        Exit Function
    End If

    size = UBound(data) - LBound(data) + 1
    If size <= 0 Then
        Err.Clear
        Exit Function
    End If

    hGlobal = GlobalAlloc(&H2, size)
    If hGlobal = 0 Then
        Err.Clear
        Exit Function
    End If

    pGlobal = GlobalLock(hGlobal)
    If pGlobal = 0 Then
        GlobalFree hGlobal
        Err.Clear
        Exit Function
    End If

    CopyMemory pGlobal, VarPtr(data(LBound(data))), size
    GlobalUnlock hGlobal

    ' fDeleteOnRelease = 1: the stream frees hGlobal
    status = CreateStreamOnHGlobal(hGlobal, 1, pStream)
    If Err.Number <> 0 Or status <> 0 Or pStream Is Nothing Then
        GlobalFree hGlobal
        Err.Clear
        Exit Function
    End If

    status = GdipCreateBitmapFromStream(pStream, gpBmp)
    Set pStream = Nothing
    If Err.Number = 0 And status = 0 And gpBmp <> 0 Then
        LoadGDIPlusBitmapFromResource = gpBmp
    End If
    Err.Clear
End Function

Private Function GpBitmapToStdPicture(gpBmp As LongPtr) As StdPicture
    Dim hBmp As LongPtr
    Dim status As Long
    Dim pic As StdPicture
    Dim pd As PICTDESC

    On Error Resume Next
    If gpBmp = 0 Then
        Err.Clear
        Exit Function
    End If

    status = GdipCreateHBITMAPFromBitmap(gpBmp, hBmp, &HFFFFFFFF)
    If status <> 0 Or hBmp = 0 Then
        Err.Clear
        Exit Function
    End If
    
    'WriteToDebugLogFile "  GpBitmapToStdPicture: passed  status <> 0 Or hBmp = 0"
    With pd
        .cbSizeofstruct = Len(pd)
        .picType = vbPicTypeBitmap
        .hImage = hBmp
        .hPalette = 0
    End With

    'WriteToDebugLogFile "  GpBitmapToStdPicture: calling OleCreatePictureIndirect "
    OleCreatePictureIndirect pd, IID_IPicture, True, pic
    If Not pic Is Nothing Then Set GpBitmapToStdPicture = pic
    Err.Clear
End Function

' checking to make sure the resources are embedded in the EXE
Private Declare Function EnumResourceTypesW Lib "kernel32" ( _
    ByVal hModule As LongPtr, _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr) As Long

Private Declare Function EnumResourceNamesW Lib "kernel32" ( _
    ByVal hModule As LongPtr, _
    ByVal lpType As LongPtr, _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" ( _
    ByVal lpModuleName As LongPtr) As LongPtr

Public Sub EnumerateAllResources()
    On Error Resume Next
    Dim hMod As LongPtr
    hMod = GetModuleHandleW(0)

    Debug.Print "Enumerating resources..."
    EnumResourceTypesW hMod, AddressOf EnumTypesCallback, 0
End Sub

Private Function EnumTypesCallback( _
    ByVal hModule As LongPtr, _
    ByVal lpType As LongPtr, _
    ByVal lParam As LongPtr) As Long

    Dim typeName As String

    If lpType < &H10000 Then
        typeName = "#" & CStr(lpType)
    Else
        typeName = StrFromPtrW(lpType)
    End If

    LogToFile "Resource Type: " & typeName
    
    EnumResourceNamesW hModule, lpType, AddressOf EnumNamesCallback, 0

    EnumTypesCallback = 1 ' continue enumeration
End Function

Private Function EnumNamesCallback( _
    ByVal hModule As LongPtr, _
    ByVal lpType As LongPtr, _
    ByVal lpName As LongPtr, _
    ByVal lParam As LongPtr) As Long

    Dim name As String

    If lpName < &H10000 Then
        name = "#" & CStr(lpName)
    Else
        name = StrFromPtrW(lpName)
    End If

    LogToFile "     Name: " & name

    EnumNamesCallback = 1 ' continue enumeration
End Function

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

Public Function StringFromPtrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function

    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function

    Dim s As String
    s = String$(cch, vbNullChar)

    CopyMemory StrPtr(s), p, cch * 2

    StringFromPtrW = s
End Function

Private Function StrFromPtrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    StrFromPtrW = StringFromPtrW(p)
End Function

Private Sub LogToFile(ByVal text As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\resource_dump.txt" For Append As #f
    If Err.Number <> 0 Then Exit Sub
    Print #f, text
    Close #f
End Sub