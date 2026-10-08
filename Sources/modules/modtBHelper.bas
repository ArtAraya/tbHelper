Attribute VB_Name = "modtBHelper"
Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
Public FO_DELETE As Long = &H3
Public FOF_ALLOWUNDO As Long = &H40

Public Const FO_COPY = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10

Public Declare Function SHFO_UnZip Lib "shell32" Alias "SHFileOperationW" (ByVal lpFileOp As Long) As Long

Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public BINDF_GETNEWESTVERSION As Long = &H10
    
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public SW_HIDE As Integer = 0
Public Const SW_SHOWNORMAL As Long = 1

' Set when install/revert is aborted because elevation was declined, cancelled,
' or a relaunch was started. Callers should skip the generic extract-failed UI.
Public g_ElevationAbortQuiet As Boolean
Public g_ElevationRelaunchStarted As Boolean

' Form1 is PredeclaredId. Unload Me from a nested click handler, then touching
' Me on the way back out, reloads a hidden second instance. End after Unload
' so the unelevated process actually exits.
Public Sub ExitAfterElevatedRelaunch()
    On Error Resume Next
    Unload Form1
    End
End Sub

Declare Function IsUserAnAdmin Lib "shell32.dll" () As Long

Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
    ByVal lpFileName As LongPtr, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_DELETE As Long = &H10000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const FILE_SHARE_DELETE As Long = 4
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As LongPtr = -1

Public Const WRITEPROBE_OK As Long = 0
Public Const WRITEPROBE_ACCESS_DENIED As Long = 5
Public Const WRITEPROBE_SHARING As Long = 32
Public Const WRITEPROBE_FAILED As Long = -1

Public Function ImageFilePath(ByVal fileName As String) As String
    On Error Resume Next
    ImageFilePath = App.Path & "\Resources\Images\" & fileName
End Function

Public Sub ConfigureCustomButton(theButton As ucCustomButton, buttonCaption As String, bkColor As OLE_COLOR, frColor As OLE_COLOR, _
    resID As ResourceID, iconSize As Integer, startEnabled As Boolean, boldFont As Boolean, _
    Optional borderColor As OLE_COLOR = 0, Optional borderWidth As Integer = 0)
    
    Dim dpiScale As Double = GetDPIScale()
    
    ' new button configuration code from AARays on VBForums
    With theButton
        .Caption = buttonCaption
        .BackColor = bkColor
        .ForeColor = frColor
        If borderWidth > 0 Then
            .BorderColor = borderColor
            .BorderWidth = borderWidth
        End If
        .FontSize = 11
        .BorderRadius = 3 * dpiScale
        .FontBold = boldFont
        .ButtonImagePtr = GetImagePtr(resID)
        .IconSize = iconSize * dpiScale
        .IconSpacing = 8 * dpiScale
        .Enabled = startEnabled
    End With
    
End Sub

Public Sub WriteToDebugLogFile(logFileLine As String)
    'Dim debugLogFile As TextStream
    
    ' its possible that a user control can call this before the form load event triggers
    If fso Is Nothing Then
        'Set fso = New FileSystemObject
        Debug.Print "fso is nothing: message is " & logFileLine
        Exit Sub
    End If

    Dim logFileName As String = App.Path & "\debug_log.txt"
    Static debugLogFile As TextStream
    
    ' if this was not used then no reason to close it
    If logFileLine = "CLOSE" And debugLogFile Is Nothing Then Exit Sub

    ' open this once during app run
    On Error GoTo errorHandler
    If debugLogFile Is Nothing Then
        
        Set debugLogFile = fso.OpenTextFile(logFileName, ForAppending, True)
    End If
    
    'Set debugLogFile = fso.OpenTextFile(logFileName, ForAppending, True)
    
    debugLogFile.WriteLine(Format(Now, "mm/dd/yy hh:MM:ss") & ": " & logFileLine)
    
    ' only close it if this is received - which should only be during form1 unload
    If logFileLine = "CLOSE" Then debugLogFile.Close()
    'debugLogFile.Close()
    
errorHandler:
    ' just skip if there is an error 
    If Err.Number <> 0 Then
        Debug.Print "WriteToDebugLogFile error " & Err.Description
    End If
End Sub

' Startup runs before Form_Load creates fso, so this log uses Open/Print and never raises.
Public Sub StartupTrace(ByVal stepName As String)
    Dim fn As Integer
    Dim logPath As String
    Static sessionStarted As Boolean
    
    On Error Resume Next
    logPath = App.Path & "\startup_log.txt"
    
    If Not sessionStarted Then
        fn = FreeFile
        Open logPath For Output As #fn
        If Err.Number = 0 Then
            Print #fn, Format(Now, "yyyy-mm-dd hh:nn:ss"); " startup path="; App.Path
            Close #fn
            sessionStarted = True
        Else
            Close #fn
            Err.Clear
            Exit Sub
        End If
    End If
    
    fn = FreeFile
    Open logPath For Append As #fn
    If Err.Number = 0 Then
        Print #fn, Format(Now, "hh:nn:ss"); " "; stepName
        Close #fn
    Else
        Close #fn
    End If
    Err.Clear
End Sub

Public Sub StartupTraceError(ByVal stepName As String)
    Dim n As Long
    Dim d As String
    Dim msg As String
    Static announced As Boolean
    
    n = Err.Number
    d = Err.Description
    Err.Clear
    
    If n = 0 Then
        StartupTrace stepName & " (no error number)"
        Exit Sub
    End If
    
    StartupTrace stepName & " ERROR " & CStr(n) & ": " & d
    
    If announced Then Exit Sub
    If IsCodeRunningInTheIDE() Then Exit Sub
    announced = True
    msg = "Startup stopped at: " & stepName & vbCrLf
    msg = msg & "Error " & CStr(n) & ": " & d & vbCrLf & vbCrLf
    msg = msg & "Details were written to:" & vbCrLf & App.Path & "\startup_log.txt"
    MsgBox msg, vbExclamation, "tBHelper"
End Sub

Public Function PixelsToTwips(pixels As Long) As Long
    PixelsToTwips = CLng(pixels * Screen.TwipsPerPixelY)
End Function

Public Function PixelsToTwipsX(pixels As Long) As Long
    PixelsToTwipsX = CLng(pixels * Screen.TwipsPerPixelX)
End Function

' Physical client size. twinBASIC is DPI-aware, so GetClientRect is the HWND
' size in device pixels. Width \ TwipsPerPixelX is not reliable above 100% DPI.
Public Sub GetClientSizePx(ByVal hWnd As LongPtr, ByRef pxWidth As Long, ByRef pxHeight As Long)
    Dim rc As RECT
    GetClientRect hWnd, rc
    pxWidth = rc.Right
    pxHeight = rc.Bottom
End Sub

Public Function ClientWidthPx(ByVal hWnd As LongPtr) As Long
    Dim w As Long, h As Long
    GetClientSizePx hWnd, w, h
    ClientWidthPx = w
End Function

Public Function ClientHeightPx(ByVal hWnd As LongPtr) As Long
    Dim w As Long, h As Long
    GetClientSizePx hWnd, w, h
    ClientHeightPx = h
End Function

Public Function ScaleDesignPx(ByVal designPixels As Long) As Long
    Dim n As Long
    n = CLng(designPixels * GetDPIScale())
    If n < 1 And designPixels > 0 Then n = 1
    ScaleDesignPx = n
End Function

' add your procedures here
Public tbHelperSettings As clsSettings
Public fso As FileSystemObject
Public chgLogs As New colChangeLogItems
Public githubReleasesURL As String = "https://github.com/twinbasic/twinbasic/releases"
Public activityLog As ucActivityLog

Public Function GetTBVersionInFolder(tBFolder As String) As String

    'WriteToDebugLogFile("GetCurrentTBVersion " & tBFolder)
    ' attempt to find the version number of twinBasic in use
    Dim fileWithVersionInfo As String
    Dim versionIndicator As String = "BETA"
    Dim fileContents As String
    Dim tempString As String
    
    fileWithVersionInfo = NormalizeFolderPath(tBFolder) & "ide\build.js"
    
    If Not fso.FileExists(fileWithVersionInfo) Then
        GetTBVersionInFolder = 0
        Exit Function
    End If
        
    ' open the file designated as the one with the version number
    fileContents = fsoFileRead(fileWithVersionInfo)
    
    ' parse the text for the version number (build.js: "twinBASIC IDE BETA 1003")
    tempString = Mid(fileContents, InStr(fileContents, versionIndicator) + Len(versionIndicator))
    tempString = Trim$(tempString)
    GetTBVersionInFolder = Val(Left$(tempString, 4))
        
    'WriteToDebugLogFile("Exit GetCurrentTBVersion")
    
End Function

Public Function fsoFileRead(filePath As String) As String
    
    If Not fso.FileExists(filePath) Then Return "Failed fsoFileRead"
    
    On Error GoTo readError
    
    Dim fso As New Scripting.FileSystemObject
        Dim fileToRead As TextStream
        
        Set fileToRead = fso.OpenTextFile(filePath, ForReading)
            fsoFileRead = fileToRead.ReadAll()
readError:
        If fsoFileRead = vbNullString Then
            MsgBox("Unable to read " & filePath, "error", "FileRead")
        End If
        fileToRead.Close()
    Set fso = Nothing
    
End Function

Public Function NormalizeFolderPath(folderPath As String) As String
    NormalizeFolderPath = folderPath
    If Len(NormalizeFolderPath) = 0 Then Exit Function
    If Right$(NormalizeFolderPath, 1) <> "\" Then NormalizeFolderPath = NormalizeFolderPath & "\"
End Function

' Can this process replace files in folderPath?
' If twinBASIC.exe exists, open it for write/delete (not virtualized).
' If the folder is empty, fall back to creating a temp file.
Public Function ProbeDestWriteAccess(folderPath As String) As Long
    Dim exePath As String
    Dim hFile As LongPtr
    Dim lastErr As Long
    
    On Error GoTo probeFailed
    ProbeDestWriteAccess = WRITEPROBE_FAILED
    If fso Is Nothing Then Exit Function
    If Not fso.FolderExists(folderPath) Then Exit Function
    
    exePath = NormalizeFolderPath(folderPath) & "twinBASIC.exe"
    If fso.FileExists(exePath) Then
        hFile = CreateFileW(StrPtr(exePath), GENERIC_WRITE Or FILE_DELETE, _
            FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
            0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        If hFile = INVALID_HANDLE_VALUE Then
            lastErr = Err.LastDllError
            If lastErr = WRITEPROBE_ACCESS_DENIED Or lastErr = WRITEPROBE_SHARING Then
                ProbeDestWriteAccess = lastErr
            Else
                ProbeDestWriteAccess = WRITEPROBE_FAILED
            End If
            Exit Function
        End If
        CloseHandle hFile
        ProbeDestWriteAccess = WRITEPROBE_OK
        Exit Function
    End If
    
    If FolderAllowsFileCreate(folderPath) Then
        ProbeDestWriteAccess = WRITEPROBE_OK
    Else
        ProbeDestWriteAccess = WRITEPROBE_ACCESS_DENIED
    End If
    Exit Function
probeFailed:
    ProbeDestWriteAccess = WRITEPROBE_ACCESS_DENIED
End Function

Public Function EnsureFolderExists(folderPath As String) As Boolean
    On Error GoTo failed
    If fso.FolderExists(folderPath) Then
        EnsureFolderExists = True
        Exit Function
    End If
    fso.CreateFolder folderPath
    EnsureFolderExists = fso.FolderExists(folderPath)
    Exit Function
failed:
    EnsureFolderExists = False
End Function

' True if a new file can be created (and deleted) inside an existing folder.
Public Function FolderAllowsFileCreate(folderPath As String) As Boolean
    On Error GoTo writeFailed
    Dim probe As String
    If Not fso.FolderExists(folderPath) Then Exit Function
    probe = NormalizeFolderPath(folderPath) & "tbHelper_write_probe.tmp"
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(probe, True)
    ts.Close
    fso.DeleteFile probe, True
    FolderAllowsFileCreate = True
    Exit Function
writeFailed:
    On Error Resume Next
    If fso.FileExists(probe) Then fso.DeleteFile probe, True
    FolderAllowsFileCreate = False
End Function

' True if a new subfolder can be created (and deleted) in parentPath.
Public Function FolderAllowsCreateSubfolder(parentPath As String) As Boolean
    On Error GoTo writeFailed
    Dim probe As String
    If Not fso.FolderExists(parentPath) Then Exit Function
    probe = NormalizeFolderPath(parentPath) & "tbHelper_write_probe_dir"
    If fso.FolderExists(probe) Then fso.DeleteFolder probe, True
    fso.CreateFolder probe
    fso.DeleteFolder probe, True
    FolderAllowsCreateSubfolder = True
    Exit Function
writeFailed:
    On Error Resume Next
    If fso.FolderExists(probe) Then fso.DeleteFolder probe, True
    FolderAllowsCreateSubfolder = False
End Function

Public Function ClearFolderContents(folderPath As String) As Boolean
    On Error GoTo clearFailed
    Dim fld As Folder
    Dim f As File
    Dim sf As Folder
    If Not fso.FolderExists(folderPath) Then
        ClearFolderContents = True
        Exit Function
    End If
    Set fld = fso.GetFolder(folderPath)
    For Each f In fld.Files
        f.Delete True
    Next
    For Each sf In fld.SubFolders
        fso.DeleteFolder sf.Path, True
    Next
    ClearFolderContents = True
    Exit Function
clearFailed:
    ClearFolderContents = False
End Function

Public Function CopyFolderContents(sourcePath As String, destPath As String) As Boolean
    On Error GoTo copyFailed
    Dim src As Folder
    Dim f As File
    Dim sf As Folder
    Dim dest As String
    dest = NormalizeFolderPath(destPath)
    If Not fso.FolderExists(dest) Then fso.CreateFolder dest
    Set src = fso.GetFolder(sourcePath)
    For Each f In src.Files
        fso.CopyFile f.Path, dest & f.Name, True
    Next
    For Each sf In src.SubFolders
        fso.CopyFolder sf.Path, dest & sf.Name, True
    Next
    CopyFolderContents = True
    Exit Function
copyFailed:
    CopyFolderContents = False
End Function

Public Function GettBParentFolder() As String
        
    Dim idx As Integer
    Dim slashCount As Integer
        
    ' loop backwards until the second \ is found - which will indicate where
    ' the parent folder for twinBASIC is
    For idx = Len(tbHelperSettings.twinBASICFolder.Path) To 1 Step -1
        If Mid(tbHelperSettings.twinBASICFolder.Path, idx, 1) = "\" Then slashCount += 1
        If slashCount = 2 Then Exit For
    Next
        
    ' truncate the value in the textbox holding the install folder, to get the parent folder
    GettBParentFolder = Left(tbHelperSettings.twinBASICFolder.Path, idx)
        
End Function

Public Function IsCodeRunningInTheIDE() As Boolean
    
    Dim strFileName As String
    Dim lngCount As Long

    On Error Resume Next
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)
    
    IsCodeRunningInTheIDE = Not InStr(UCase(strFileName), "TWINBASIC_WIN32") = 0
    If Err.Number <> 0 Then IsCodeRunningInTheIDE = False
     
End Function

Public Function IsElevated() As Boolean
    On Error Resume Next
    IsElevated = (IsUserAnAdmin() <> 0)
    If Err.Number <> 0 Then IsElevated = False
End Function

Public Function ParseResumeCommand(ByRef resumeSwitch As String, ByRef zipPath As String) As Boolean
    Dim cmd As String
    Dim q As Long

    On Error GoTo parseFailed
    resumeSwitch = vbNullString
    zipPath = vbNullString
    cmd = Trim$(Command$)
    If Len(cmd) = 0 Then Exit Function

    If LCase$(Left$(cmd, 14)) = "/resumeinstall" Then
        resumeSwitch = "/resumeInstall"
        zipPath = Trim$(Mid$(cmd, 15))
    ElseIf LCase$(Left$(cmd, 13)) = "/resumerevert" Then
        resumeSwitch = "/resumeRevert"
        zipPath = Trim$(Mid$(cmd, 14))
    Else
        Exit Function
    End If

    If Left$(zipPath, 1) = """" Then
        zipPath = Mid$(zipPath, 2)
        q = InStr(zipPath, """")
        If q > 0 Then zipPath = Left$(zipPath, q - 1)
    End If

    ParseResumeCommand = (Len(zipPath) > 0)
    Exit Function
parseFailed:
    resumeSwitch = vbNullString
    zipPath = vbNullString
    ParseResumeCommand = False
End Function

Public Function RelaunchElevatedAndResume(ByVal ownerHwnd As Long, ByVal resumeSwitch As String, ByVal zipPath As String) As Boolean
    Dim exePath As String
    Dim args As String
    Dim rc As Long
    Dim exeName As String

    On Error GoTo relaunchFailed
    If IsCodeRunningInTheIDE Then
        RelaunchElevatedAndResume = False
        Exit Function
    End If

    exePath = App.Path
    If Right$(exePath, 1) <> "\" Then exePath = exePath & "\"
    exeName = App.EXEName
    If LCase$(Right$(exeName, 4)) <> ".exe" Then exeName = exeName & ".exe"
    exePath = exePath & exeName

    args = resumeSwitch & " """ & zipPath & """"
    rc = ShellExecute(ownerHwnd, "runas", exePath, args, App.Path, SW_SHOWNORMAL)
    RelaunchElevatedAndResume = (rc > 32)
    Exit Function
relaunchFailed:
    RelaunchElevatedAndResume = False
End Function

Public Function IsProcessRunning(ByVal ProcessName As String) As Boolean
    
    ' is twinBASIC running? 
    Dim objWMI As Object, colProcesses As Variant, objProcess As Variant

    On Error GoTo wmiFailed
    ' Get the WMI service object
    Set objWMI = GetObject("winmgmts:\\")

    ' Query for processes
    Set colProcesses = objWMI.ExecQuery("Select * From Win32_Process Where Name='" & ProcessName & "'")

    ' Check if any processes matching the name were found
    If colProcesses.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If

    ' Clean up objects
    Set objProcess = Nothing
    Set colProcesses = Nothing
    Set objWMI = Nothing
    Exit Function
wmiFailed:
    IsProcessRunning = False
End Function

Public Sub UpdateActivityLog(statMessage As String, Optional updatePreviousStatus As Boolean = False)
    
    'WriteToDebugLogFile("In ShowStatusMessage " & statMessage)
    ' write the message to the activity log
    If updatePreviousStatus Then
        activityLog.AddEntry "", statMessage, True
    Else
        activityLog.AddEntry Format(Now, "MM/dd/yy hh:mm:ss AM/PM: "), statMessage
    End If

   ' WriteToDebugLogFile("Out ShowStatusMessage ")
End Sub

Private Sub CenterPanel(pnlToCenter As Frame, Optional inObject As Object)
    
    ' default the object to center the panel in to Form1 if
    ' none is supplied
    If inObject Is Nothing Then
        Set inObject = Form1
    End If
        
    Dim x As Long, y As Long
    x = (inObject.Width - pnlToCenter.Width) \ 2
    y = (inObject.Height - pnlToCenter.Height) \ 2
    pnlToCenter.Left = x
    pnlToCenter.Top = y
        
End Sub

Public currentPanelTop As Long
Public currentPanelLeft As Long
Private picIcon As PictureBox

Public Sub ShowPanelView(innerPanel As Frame, Optional radius As Long = 10)
    
    ' this will be called to display a hidden panel / frames 
    ' use the parent "frame panel" as the drop shadow for the "frame form"
    Dim parentPanel As Frame
    If TypeOf innerPanel Is Frame Then
        Set parentPanel = innerPanel.Container
    End If
    
    currentPanelTop = parentPanel.Top
    currentPanelLeft = parentPanel.Left
    
    Dim dpi As Double = GetDPIScale()
    
    parentPanel.BackColor = RGB(180, 180, 180)  ' light gray shadow
    
    innerPanel.BackColor = RGB(240, 240, 240)   ' lighter background
    
    ' 90/100 are in the same scale as Frame.Width/Height (twips)
    parentPanel.Width = innerPanel.Width + 90
    parentPanel.Height = innerPanel.Height + 100
    
    CenterPanel parentPanel                ' center the parent of the inner panel in the mail form
    CenterPanel innerPanel, parentPanel    ' center the inner panel in the parent
    
    ' add icon to the panel
    Set picIcon = Form1.picPanelIcon
    If InStr(innerPanel.Name, "Revert") > 0 Then
        DisplayPanelIcon resRevertPanelIcon, innerPanel
    
    ElseIf InStr(innerPanel.Name, "ViewLog") > 0 Then
        DisplayPanelIcon resLogHistoryPanelIcon, innerPanel
        
    ElseIf InStr(innerPanel.Name, "Folder") > 0 Then
        DisplayPanelIcon resBlackFolder, innerPanel

    Else
        DisplayPanelIcon resMessageBox, innerPanel
            
    End If
    
    parentPanel.Visible = True
    parentPanel.ZOrder 0
    
    ' Region must match the final HWND size (GetClientRect is valid after Visible)
    ApplyRoundedRegion parentPanel, CLng(radius * dpi)
    ApplyRoundedRegion innerPanel, CLng(12 * dpi)

    Form1.isAPanelDisplayed = True
End Sub

Private Sub DisplayPanelIcon(iconResourceID As ResourceID, parentContainer As Frame)
    Dim pic As StdPicture
    
    On Error Resume Next
    Set pic = GetImageStd(iconResourceID)
    If picIcon Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    If Not pic Is Nothing Then picIcon.Picture = pic
    picIcon.AutoSize = True
    picIcon.Top = parentContainer.Top + 35
    picIcon.Left = parentContainer.Left + 60
    picIcon.PictureDpiScaling = True
    picIcon.Visible = True
    
    SetParent picIcon.hWnd, parentContainer.hWnd
    Err.Clear
End Sub


Public Sub HidePanelView(parentPanel As Frame)
    
    ' hide the panel and put it back where it was
    Form1.isAPanelDisplayed = False
    parentPanel.Visible = False
        
    ' put the frame back off the screen
    parentPanel.Top = currentPanelTop
    parentPanel.Left = currentPanelLeft
    
    FlushRedraws()
    
End Sub

Private Declare PtrSafe Function SHGetKnownFolderPath Lib "shell32" ( _
    ByRef rfid As GUID, _
    ByVal dwFlags As Long, _
    ByVal hToken As LongPtr, _
    ByRef pszPath As LongPtr _
) As Long

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' needed to declare this here, even though it is in WinDevLib because the call to it
' was failing to get the folder I was asking for
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As LongPtr, _
    ByRef pclsid As GUID _
) As Long

Sub GetLocalFolders()
    Dim hr As Long
    Dim newFolder As clsSettingsFolder
    
    On Error GoTo LocalFoldersFailed
    StartupTrace "GetLocalFolders enter"
    
    ' this retrieves the default download folder for the user
    Dim FOLDERID_Downloads As String = "{374DE290-123F-4565-9164-39C4925E467B}"
    
    Dim folderGUID As GUID
    CLSIDFromString StrPtr(FOLDERID_Downloads), folderGUID

    Dim pszPath As LongPtr
    hr = SHGetKnownFolderPath(folderGUID, 0, 0, pszPath)
    StartupTrace "GetLocalFolders SHGetKnownFolderPath " & CStr(hr)
    If hr = 0 Then
        Set newFolder = New clsSettingsFolder
        newFolder.IstwinBASIC = False
        newFolder.Path = SysAllocString(pszPath) & "\"
        StartupTrace "GetLocalFolders download path set"
        tbHelperSettings.DownloadFolder = newFolder
        CoTaskMemFree pszPath
    End If
    
    StartupTrace "GetLocalFolders twinBASIC folder"
    tbHelperSettings.twinBASICFolder.Path = GettwinBASICInstallPath()
    StartupTrace "GetLocalFolders leave"
    Exit Sub
LocalFoldersFailed:
    StartupTraceError "GetLocalFolders"
End Sub

Function GettwinBASICInstallPath() As String

    Dim hKey As LongPtr
    Dim result As Long
    Dim dataSize As Long
    Dim dataBuffer() As Byte
    Dim lpType As Long
    Dim regValue As String = ""

    ' Open the registry key
    result = RegOpenKeyEx(HKEY_CLASSES_ROOT, "Applications\twinBASIC.exe\shell\open\command", 0, KEY_READ, hKey)
    
    If result = 0 Then
        ' returned no error, get the value
        dataSize = 1024
        ReDim dataBuffer(dataSize - 1)
        
        result = RegQueryValueEx(hKey, "", 0, lpType, dataBuffer(0), dataSize)
        If result = 0 Then
            
            regValue = BytesToUnicodeString(dataBuffer) ' get the complete registry value
            regValue = Left(regValue, InStr(UCase(regValue), "TWINBASIC.EXE") - 1) ' to get the path, read up to the exe name
            regValue = Replace(regValue, Chr(34), "") ' remove any extra double quotes in the string
            
        End If
                       
        RegCloseKey hKey
    End If
        
    GettwinBASICInstallPath = regValue
    
End Function

Function BytesToUnicodeString(dataBuffer() As Byte) As String
    
    ' loop the databuffer to rebuild the string
    Dim i As Long
    Dim result As String

    ' every even byte is a 0, skip them to build the string, 
    ' when a 0 is encountered where a valid character should be, exit the for
    ' as everything required has been added to the string
    For i = 0 To UBound(dataBuffer) Step 2
        If dataBuffer(i) = 0 Then Exit For
        result = result & ChrW(dataBuffer(i))
    Next i

    BytesToUnicodeString = result
End Function

Private Const BASE_DPI As Long = 96

Public Function GetDPIScale() As Double
    ' code given by AARays on VBForums 
    ' I'm using the declarations in WinDevLib for the API calls referenced here (at first)
    ' Returns the system DPI scaling factor (e.g., 1.5 for 150% scaling)
    
    #If VBA7 Then
        Dim hDC As LongPtr
    #Else
        Dim hDC As Long
    #End If
    Dim CurrentDPI As Long
    Dim ScaleFactor As Double

    ' 1. Get the Device Context (DC) for the desktop window (hWnd=0)
    hDC = GetDC(0)

    If hDC <> 0 Then
        ' 2. Retrieve the current DPI value (e.g., 144 DPI for 150% scaling)
        CurrentDPI = GetDeviceCaps(hDC, LOGPIXELSX)

        ' 3. Release the Device Context (essential cleanup)
        Call ReleaseDC(0, hDC)

        ' 4. Calculate the fractional scale factor using floating-point division
        If CurrentDPI > 0 Then
            ScaleFactor = CDbl(CurrentDPI) / BASE_DPI
            GetDPIScale = ScaleFactor
        Else
            GetDPIScale = 1.0 ' Default to 1.0 (100%) if DPI failed to retrieve
        End If
    Else
        ' Failed to get DC
        GetDPIScale = 1.0
    End If
    
End Function

