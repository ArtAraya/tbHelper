Option Explicit

' Types
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

' API Declarations (gdiplus.dll)
Private Declare Function GdiplusStartup Lib "gdiplus" (token As LongPtr, input As GdiplusStartupInput, ByVal output As LongPtr) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long

' Module-level state
Private m_gdiToken As LongPtr    ' token returned by GdiplusStartup
Private m_refCount As Long       ' simple reference counter
Private m_started As Boolean

Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long

Private Const WM_MOUSEWHEEL As Long = &H20A

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const GWL_WNDPROC As Long = -4

Private CallbackOwner As Object
Public OriginalCanvasProc As Long

Public ucDictionary As New Scripting.Dictionary   ' my own dictionary to hold window handle to object (user controls) 

Public Sub RegisterScrollableCanvas(ByVal hWnd As Long, ByVal ownerCtrl As Object)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Canvas_WindowProc)
    ucDictionary.Add hWnd, ownerCtrl
End Sub

Public Const ROW_ALT_COLOR = &HF8F8F8
Public Const CUST_BTN_BCOLOR = &HA2640C

Public Sub ApplyRoundedRegion(toWhichControl As Control, borderRadius As Long)
    
    On Error GoTo AddRoundedRegionError
    
    ''WriteToDebugLogFile "   ApplyRoundedRegion: to " & toWhichControl.name
        
    Dim ctrl As Object
    Dim ctrlName As String
    Dim ctrlWidth As Long
    Dim ctrlHeight As Long
    
    ' is this a user control?
    If ucDictionary.Exists(toWhichControl.hWnd) Then
        Set ctrl = ucDictionary(toWhichControl.hWnd)
        If Not ctrl Is Nothing Then
            ctrlName = ctrl.name
        End If
    Else
        ' it is not a user control
        ctrlName = toWhichControl.name
    End If
    
    ctrlWidth = toWhichControl.Width
    ctrlHeight = toWhichControl.Height
    
    If ctrlWidth <= 0 Or ctrlHeight <= 0 Then Exit Sub
    If borderRadius <= 0 Then Exit Sub
    
    Dim w As Long: w = ctrlWidth \ Screen.TwipsPerPixelX
    Dim h As Long: h = ctrlHeight \ Screen.TwipsPerPixelY
    ' borderRadius is specified in 96-DPI pixels. The region itself is in
    ' physical pixels, so grow the corner with the same scale as the window.
    Dim r As Long: r = CLng(borderRadius * GetDPIScale())
    If r < 1 Then r = 1
    If (r * 2) > w And w > 1 Then r = w \ 2
    If (r * 2) > h And h > 1 Then r = h \ 2
        
    Dim rgn As Long
    rgn = CreateRoundRectRgn(0, 0, w + 1, h + 1, r * 2, r * 2)
        
    If rgn <> 0 Then
        SetWindowRgn toWhichControl.hWnd, rgn, True
    End If
    
    Set ctrl = Nothing
    Exit Sub
AddRoundedRegionError:

    Debug.Print "ApplyRoundedRegion: unable to apply to " & ctrlName & " due to: " & Err.Description
End Sub

Public Function Canvas_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If uMsg = WM_MOUSEWHEEL Then
        
        Dim newPos As Long
        newPos = GetScrollPos(hwnd, SB_VERT) ' Get current vertical scroll position

        UpdateScrollOwnership hwnd, newPos

        If Not CallbackOwner Is Nothing Then
            'WriteToDebugLogFile("Canvas_WindowProc: calling HandleMouseScroll in " & CallbackOwner.name)
            CallByName CallbackOwner, "HandleMouseScroll", vbMethod, wParam
        End If

        Exit Function
    End If

    Canvas_WindowProc = CallWindowProc(OriginalCanvasProc, hwnd, uMsg, wParam, lParam)
End Function

Public Sub DrawShadowBehind(ctrl As Control)
    ' Dim g As StdGraphics
    ' Set g = GetGraphics(ctrl.Parent.hWnd)

    ' Dim x As Long: x = ctrl.Left - 6
    ' Dim y As Long: y = ctrl.Top - 6
    ' Dim w As Long: w = ctrl.Width + 12
    ' Dim h As Long: h = ctrl.Height + 12

    ' g.FillRectangle ARGB(60, 0, 0, 0), x, y, w, h
    ' g.Dispose
End Sub

Public Sub FlushRedraws()
    
    If RedrawQueue.Count = 0 Then Exit Sub
    
    'WriteToDebugLogFile("     Flushing the Redraws - there are " & RedrawQueue.Count & " requiring attention")
    
    Dim req As clsRedrawRequest
    Dim hWnds As String = "|"    ' holds the handles redraw during this pass
    Dim now As Double
    Dim countRedrawItems As Integer = 1
    Dim reqProcessed As Boolean
    
    now = Timer ' Capture current timestamp once for consistency

    For Each req In RedrawQueue
        ' check that enough times has passed since last redraw, and the handle hasn't been refreshed
        ' during this loop, refresh it
        reqProcessed = False
        If ucDictionary.Exists(req.hWnd) Then
            If TypeOf ucDictionary(req.hWnd) Is ucCustomButton Then
                'WriteToDebugLogFile "       FlushRedraws, request #" & countRedrawItems & " is an update for ucCustomButton with caption: " & ucDictionary(req.hWnd).caption
            Else
                'WriteToDebugLogFile("       FlushRedraws, request #" & countRedrawItems & " RedrawQueue loop control to redraw: " & ucDictionary(req.hWnd).Name & " tag: " & ucDictionary(req.hWnd).Tag)
            End If
        End If
        
        If now - req.LastRequested > 0.05 And InStr(hWnds, "|" & req.hWnd & "|") = 0 Then
            If req.Region <> 0 Then
                RedrawWindow req.hWnd, ByVal req.Region, 0, RDW_INVALIDATE
                'RedrawWindow req.hWnd, req.Region, 0, RDW_INVALIDATE
                'WriteToDebugLogFile "       FlushRedraws: request #" & countRedrawItems & " RedrawWindow called"
                reqProcessed = True
            Else
                InvalidateRect req.hWnd, ByVal 0&, True
                'WriteToDebugLogFile "       FlushRedraws: request #" & countRedrawItems & " InvalidateRect called"
                reqProcessed = True
            End If
            
            hWnds = hWnds & CStr(req.hWnd) & "|"
            
            ' if this flush is for a custom button, then tell it the redraw via flushredraws is done.
            If ucDictionary.Exists(req.hWnd) Then
                On Error Resume Next
                CallByName ucDictionary(req.hWnd), "FlushRedrawComplete", vbMethod
                If Err.Number = 0 Then
                    'WriteToDebugLogFile("       FlushRedraws: request #" & countRedrawItems & " this is a button, so call its FlushRedrawComplete ")
                Else
                    'WriteToDebugLogFile("       FlushRedraws: request #" & countRedrawItems & " error# " & Err.Number & " error: " & Err.Description)
                End If
                On Error GoTo 0
            End If
            
            req.Processed = reqProcessed
           
        Else
            If ucDictionary.Exists(req.hWnd) Then
                'WriteToDebugLogFile "       FlushRedraws: If statement failed to run the redraw action for " & IIf(Not ucDictionary(req.hWnd) Is Nothing, ucDictionary(req.hWnd).Name, "[not in ucDictionary]")
            Else
                'WriteToDebugLogFile "       FlushRedraws: If statement failed to run the redraw action for handle " & req.hWnd
            End If
            

        End If
        countRedrawItems += 1
    Next
    
    'WriteToDebugLogFile("     FlushRedraws - count is " & CStr(countRedrawItems - 1))  ' back out the very last count at the end of the loop

    ' remove the requests that were successful
    Dim controlName As String
    Dim buttonCaption As String
    
    'WriteToDebugLogFile "       FlushRedraws: removing processed requests "
    For countRedrawItems = RedrawQueue.Count To 1 Step -1
        Set req = RedrawQueue.Item(countRedrawItems)
        If req.Processed Then
            buttonCaption = ""
            controlName = " [not in ucDictionary]"
            If ucDictionary.Exists(req.hWnd) Then
                On Error Resume Next
                controlName = ucDictionary(req.hWnd).Name
                If TypeOf ucDictionary(req.hWnd) Is ucCustomButton Then buttonCaption = ucDictionary(req.hWnd).caption
                'Debug.Print "FlushRedraws: remove the requests that were successful for " & controlName & " due to: " & Err.Description
                On Error GoTo 0
            End If
            
            RedrawQueue.Remove countRedrawItems
            'WriteToDebugLogFile "       FlushRedraws: removed processed request #" & countRedrawItems & " for " & controlName & IIf(Len(buttonCaption) > 0, " with caption: " & buttonCaption, "")
        End If
    Next
    
    Form1.tmrCatchUp.Enabled = RedrawQueue.Count > 0    ' if there are items left in the cue, run the timer, else disable it

    'WriteToDebugLogFile "     FlushRedraws - done"
End Sub
Public Function GDIPlus_EnsureStarted() As Boolean
    Dim ret As Long
    Dim s As GdiplusStartupInput

    If m_started Then
        m_refCount = m_refCount + 1
        GDIPlus_EnsureStarted = True
        Exit Function
    End If

    ' Fill startup input
    s.GdiplusVersion = 1
    s.DebugEventCallback = 0
    s.SuppressBackgroundThread = 0
    s.SuppressExternalCodecs = 0

    ' Start GDI+
    ret = GdiplusStartup(m_gdiToken, s, 0)
    If ret = 0 Then    ' Ok (Gdiplus::Ok == 0)
        m_started = True
        m_refCount = 1
        GDIPlus_EnsureStarted = True
    Else
        ' Failed to initialize GDI+. ret contains the GpStatus code.
        m_started = False
        m_refCount = 0
        GDIPlus_EnsureStarted = False
    End If
End Function

Public Sub GDIPlus_ForceShutdown()
    ' Optional: force shutdown (for debugging). Use with care.
    If m_started Then
        On Error Resume Next
        GdiplusShutdown m_gdiToken
        m_gdiToken = 0
        m_refCount = 0
        m_started = False
        On Error GoTo 0
    End If
End Sub

Public Function GDIPlus_IsStarted() As Boolean
    ' Call this before any GDI+ calls (controls can call on Create/Init)
    GDIPlus_IsStarted = m_started
End Function

Public Sub GDIPlus_Release()
    ' Call this when a control/form no longer needs GDI+ (e.g., in Terminate/Unload)
    If Not m_started Then Exit Sub

    If m_refCount > 1 Then
        m_refCount = m_refCount - 1
        Exit Sub
    End If

    ' Last release -> shutdown
    On Error Resume Next
    GdiplusShutdown m_gdiToken
    m_gdiToken = 0
    m_refCount = 0
    m_started = False
    On Error GoTo 0
End Sub

Public Sub InvalidateLabelRegion(lbl As Label)
    
    ' the update of the tB version label lags, this is an attempt 
    ' at fixing it by finding a region (the lable area) on the form
    ' and having it updated
    
    'WriteToDebugLogFile("InvalidateLabelRegion for '" & lbl.Name & "'")
    
    Dim rc As RECT
    
    With rc
        .Left = lbl.Left / Screen.TwipsPerPixelX
        .Right = (lbl.Left + lbl.Width) / Screen.TwipsPerPixelX
        .Top = lbl.Top / Screen.TwipsPerPixelY
        .Bottom = (lbl.Top + lbl.Height) / Screen.TwipsPerPixelY
    End With
    
    InvalidateRect lbl.Parent.hWnd, rc, 1
    UpdateWindow lbl.Parent.hWnd
    
    QueueRedraw(lbl.Parent.hWnd)
    
End Sub
Public Sub UpdateScrollOwnership(ByVal hWnd As LongPtr, ByVal newPos As Long)
        
    If ucDictionary.Exists(hWnd) Then

        Set CallbackOwner = ucDictionary(hWnd)
        
        On Error GoTo errorHandler   ' this is mainly used to stop the code from breaking while in design mode
        
        ' there is no need to track scrolls on frmMain controls if a panel is displayed
        If CallbackOwner.parent.name = "frmMain" Then
            If CallbackOwner.parent.isAPaneldisplayed And (CallbackOwner.tag <> "LogHistoryView" Or CallbackOwner.tag <> "RevertView") Then
                'WriteToDebugLogFile("   UpdateScrollOwnership: ignore call " & CallbackOwner.Name)
                Exit Sub
            End If
        End If
        
        'WriteToDebugLogFile("   UpdateScrollOwnership: called from " & CallbackOwner.Name)
        If Not CallbackOwner Is Nothing	Then
            ' Update scroll registry and queue redraw
            If ScrollRegistry Is Nothing Then Set ScrollRegistry = New Collection
            
            If ScrollRegistry.Exists(CStr(hWnd)) Then
                Dim reg As clsScrollRegistry
                Set reg = ScrollRegistry(CStr(hWnd))
                reg.LastScrollPos = newPos
            End If

            QueueRedraw hWnd
        Else
            'WriteToDebugLogFile("UpdateScrollOwnership: from something not in ucDictionary")
            
        End If
    End If
    
    errorHandler:
        ' just skip if there is an error 
        'If Err.Number <> 0 Then WriteToDebugLogFile "  *********************** UpdateScrollOwnership error " & Err.Description
End Sub

Public RedrawQueue As Collection

Public Sub QueueRedraw(hWnd As LongPtr)
    
    If RedrawQueue Is Nothing Then Set RedrawQueue = New Collection
    
    If Not RedrawQueue.Exists(CStr(hWnd)) Then
        Dim req As New clsRedrawRequest
        req.hWnd = hWnd
        req.LastRequested = Timer
        RedrawQueue.Add req, CStr(hWnd)
        RedrawQueue(CStr(hWnd)).LastRequested = Timer
        'WriteToDebugLogFile("     QueueRedraw: added new request")
    Else
        RedrawQueue(CStr(hWnd)).LastRequested = Timer
        'WriteToDebugLogFile("     QueueRedraw: updated last requested time")
    End If
    
    If ucDictionary.Exists(hWnd) Then
        If TypeOf ucDictionary(hWnd) Is ucCustomButton Then
            'WriteToDebugLogFile "     QueueRedraw for ucCustomButton with caption: " & ucDictionary(hWnd).caption
        Else
            'WriteToDebugLogFile("     QueueRedraw: called from " & ucDictionary(hWnd).Name)
        End If
        
    Else
        'WriteToDebugLogFile("     QueueRedraw: called from something not in ucDictionary")
    End If
    
End Sub

Private ScrollRegistry As Collection

Public Sub RegisterScrollOwner(hWnd As LongPtr)
    
    'WriteToDebugLogFile("RegisterScrollOwner hWnd=" & hWnd)

    If ucDictionary.Exists(hWnd) Then
        'WriteToDebugLogFile("   RegisterScrollOwner: called from " & ucDictionary(hWnd).Name & " on " & ucDictionary(hWnd).parent.name)
    End If
    
    Dim state As clsScrollRegistry
    state.hWnd = hWnd
    state.IsActive = True
    state.LastScrollPos = 0
    
    If ScrollRegistry Is Nothing Then Set ScrollRegistry = New Collection
    
    ScrollRegistry.Add state, CStr(hWnd)
End Sub

