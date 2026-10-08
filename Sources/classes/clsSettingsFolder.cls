VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "clsSettingsFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' add your member fields and procedures here
' clsSettingsFolder.cls
Option Explicit

Private _path As String
Private _isValid As Boolean
Private _invalidReason As String
Private _istwinBASIC As Boolean
Private _tBVersion As Integer

Public Property Let IstwinBASIC(ByVal value As Boolean)
    _istwinBASIC = value
End Property

Public Property Get IstwinBASIC() As Boolean
    IstwinBASIC = _istwinBASIC
End Property

Public Property Get tBVersion() As Integer
    tBVersion = _tBVersion
End Property

Public Property Get Path() As String
    Path = _path
End Property

Public Property Let Path(ByVal value As String)
    _path = value
    ValidatePath
End Property

Public Property Get IsValid() As Boolean
    IsValid = _isValid
End Property

Public Property Get InvalidReason() As String
    InvalidReason = _invalidReason
End Property

Public Sub ForceValid()
    ' used when a new folder is created and hasn't been extracted to yet
    _isValid = True
End Sub

Private Sub ValidatePath()

    If _path = "" Then
        _isValid = False
        _invalidReason = "Path is empty."
        
    ElseIf Not fso.FolderExists(_path) Then
        _isValid = False
        _invalidReason = "Folder does not exist."
        
    ElseIf _istwinBASIC And (Not fso.FileExists(_path & "\twinBASIC.exe")) Then
        _isValid = False
        _invalidReason = "Folder does not contain twinBASIC."
        
    ElseIf _istwinBASIC And GettwinBASICVersion() = 0 Then
        _isValid = False
        _invalidReason = "twinBASIC folder contains unkown version."
    
    Else
        _isValid = True
        _invalidReason = ""
        
    End If
End Sub

Private Function GettwinBASICVersion() As Integer

    'WriteToDebugLogFile("GetCurrentTBVersion " & _path)
    ' attempt to find the version number of twinBasic in use
    Dim fileWithVersionInfo As String = _path & "\ide\build.js"
    Dim versionIndicator As String = "BETA"
    Dim fileContents As String
    Dim tempString As String
    
    fileWithVersionInfo = Replace(fileWithVersionInfo, "\\", "\") ' make sure there are just single forward slashes in the path
    
    If Not fso.FileExists(fileWithVersionInfo) Then
        _tBVersion = 0
        GettwinBASICVersion = 0
        _isValid = False
        'WriteToDebugLogFile("Exit GetCurrentTBVersion with version 0")
        Exit Function
    End If
        
    ' open the file designated as the one with the version number
    fileContents = fsoFileRead(fileWithVersionInfo)
    
    ' parse the text for the version number (build.js: "twinBASIC IDE BETA 1003")
    tempString = Mid(fileContents, InStr(fileContents, versionIndicator) + Len(versionIndicator))
    tempString = Trim$(tempString)
    GettwinBASICVersion = Val(Left$(tempString, 4))
    _tBVersion = GettwinBASICVersion
    'WriteToDebugLogFile("Exit GetCurrentTBVersion with version " & _tBVersion)
    
End Function