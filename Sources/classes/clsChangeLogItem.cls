VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChangeLogItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6", "Yes"
Attribute VB_Ext_KEY = "Top_Level", "No"

'local variable(s) to hold property value(s)
Private _verion As Integer
Private _type As String
Private _notes As String

Public Property Let tBVersion(ByVal vData As Integer)
    _verion = vData
End Property

Public Property Get tBVersion() As Integer

    tBVersion = _verion
End Property

Public Property Let Type(ByVal vData As String)

    _type = vData
End Property

Public Property Get Type() As String

    Type = _type
End Property

Public Property Let Notes(ByVal vData As String)

    _notes = vData
End Property

Public Property Get Notes() As String

    Notes = _notes
End Property
