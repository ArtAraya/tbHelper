VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "clsChangeLogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsChangeLogEntry
Private m_Values As Collection
Private m_TypeText As String
Private m_TypeColor As Long
Private m_TypeTextColor As Long
Private m_EntryHeight As Long

' colors for the type of change log item      the names came from this site https://colornamer.robertcooper.me/
Private Const COLOR_IMPORTANT = &HA2640C          ' Raw Umber  (orange/yellow)
Private Const COLOR_IMPROVED = RGB(107, 142, 35)  ' Scallion  (light green)
Private Const COLOR_FIXED = RGB(56, 118, 29)      ' Forest Empress  (darker green)
Private Const COLOR_UPDATED = RGB(107, 45, 103)   ' Intense Mauve 
Private Const COLOR_TIP = RGB(22, 83, 126)        ' Lake Baikal (blue/grey)
Private Const COLOR_DEFAULT = &H808080            ' standard Grey
Private Const COLOR_KNOWN = vbBlack
Private Const COLOR_ADDED = RGB(75, 0, 130)       ' Indigo (purple)
Private Const COLOR_WARNING = RGB(153, 0, 0)      ' Crimson Red 
Private Const COLOR_CHANGED = RGB(198, 79, 55)    ' Burnt Sienna  (Trinidad)
Private Const COLOR_NOTE = RGB(72, 61, 139)       ' Nauseous Blue

Private Const TEXT_COLOR_WHITE = &HFFFFFF

Public Property Get EntryHeight() As Long
    EntryHeight = m_EntryHeight
End Property

Public Property Let EntryHeight(ByVal c As Long)
    m_EntryHeight = c
End Property

Public Property Get Values() As Collection
    Set Values = m_Values
End Property

Public Property Get TypeText() As String
    TypeText = m_TypeText
End Property

Public Property Let TypeText(ByVal s As String)
    m_TypeText = s
End Property

Public Property Get TypeColor() As Long
    TypeColor = m_TypeColor
End Property

Public Property Let TypeColor(ByVal c As Long)
    m_TypeColor = c
End Property

Public Property Get TypeTextColor() As Long
    TypeTextColor = m_TypeTextColor
End Property

Public Property Let TypeTextColor(ByVal c As Long)
    m_TypeTextColor = c
End Property

Public Sub Initialize(columnDefs() As clsChangeLogColumn, columnValues() As Variant)
    
    Set m_Values = New Collection
    Dim i As Integer = 0
    Dim value As String

    Dim col As clsChangeLogColumn
    For Each col In columnDefs
    
        value = columnValues(i)
        m_Values.Add value, col.Name

        If UCase(col.Name) = "TYPE" Then
            m_TypeText = UCase(value)
            m_TypeTextColor = TEXT_COLOR_WHITE

            Select Case m_TypeText
                Case "IMPORTANT": m_TypeColor = COLOR_IMPORTANT
                Case "IMPROVED":  m_TypeColor = COLOR_IMPROVED
                Case "FIXED":     m_TypeColor = COLOR_FIXED
                Case "UPDATED":   m_TypeColor = COLOR_UPDATED
                Case "TIP":       m_TypeColor = COLOR_TIP
                Case "KNOWN ISSUE": m_TypeColor = COLOR_KNOWN
                Case "ADDED":     m_TypeColor = COLOR_ADDED
                Case "WARNING":   m_TypeColor = COLOR_WARNING
                Case "CHANGED":   m_TypeColor = COLOR_CHANGED
                Case "NOTE":      m_TypeColor = COLOR_NOTE
                Case Else:        m_TypeColor = COLOR_DEFAULT
            End Select
        End If
        
        i += 1
    Next col
End Sub