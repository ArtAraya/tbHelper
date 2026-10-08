VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "clsHistoryLogItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' add your member fields and procedures here
Private _logMessage As String
Private _changeLogItem As clsChangeLogItem
Private _logDateTime As String
Private _logDate As String

Public Property Let LogCLI(ByVal ChangeLogItem As clsChangeLogItem)
   Set _changeLogItem = ChangeLogItem
End Property

Public Property Get LogCLI() As clsChangeLogItem
   Set LogCLI = _changeLogItem
End Property

Public Property Let LogMessage(ByVal lMsg As String)
    _logMessage = lMsg
End Property

Public Property Get LogMessage() As String
    LogMessage = _logMessage 
End Property

Public Property Let LogDateTime(ByVal logDateTime As String)
    _logDateTime = logDateTime
    _logDate = Format(CDate(logDateTime), "MM/dd/yy")
End Property

Public Property Get LogDateTime() As String
    LogDateTime = _logDateTime 
End Property

Public Property Get LogDate() As String
    LogDate = _logDate 
End Property