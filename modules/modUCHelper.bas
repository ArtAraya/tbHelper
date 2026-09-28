Attribute VB_Name = "modUCHelper"

' add your procedures here
Public Sub ApplyRoundedCorners(ctl As Object, ByVal RadiusX As Long, ByVal RadiusY As Long)
    Dim hRgn As LongPtr
    Dim clientRect As RECT
    Dim controlName As String = ""
    
    On Error GoTo RoundedCornersError
    GetClientRect ctl.hWnd, clientRect
    
    controlName = ctl.name
    
    ' clientRect.Right IS the width in pixels.
    ' clientRect.Bottom IS the height in pixels.
    ' Radius arguments are 96-DPI pixels.
    RadiusX = CLng(RadiusX * GetDPIScale())
    RadiusY = CLng(RadiusY * GetDPIScale())
    If RadiusX < 1 Then RadiusX = 1
    If RadiusY < 1 Then RadiusY = 1
    hRgn = CreateRoundRectRgn(0, 0, clientRect.Right, clientRect.Bottom, RadiusX, RadiusY)
    
    SetWindowRgn ctl.hWnd, hRgn, True
    On Error GoTo 0
    
    Exit Sub
    RoundedCornersError:
    Debug.Print "ApplyRoundedCorners: unable to apply to " & controlName & " due to: " & Err.Description
    
End Sub

Public Sub ApplyBottomRoundedCorners(ctl As Object, ByVal RadiusX As Long, ByVal RadiusY As Long, Optional ByVal hasVerticalScrollbar As Boolean = False)
    Dim hRgn_Combined As LongPtr
    Dim hRgn_TopRect As LongPtr
    Dim hRgn_BottomRoundRect As LongPtr
    Dim clientRect As RECT
    
    ' Get the dimensions of the control's CLIENT area (the part inside borders/scrollbars)
    GetClientRect ctl.hWnd, clientRect
    
    Dim PxWidth As Long
    Dim PxHeight As Long
    PxWidth = clientRect.Right   ' For a client rect, Right IS the width
    PxHeight = clientRect.Bottom ' and Bottom IS the height

    RadiusX = CLng(RadiusX * GetDPIScale())
    RadiusY = CLng(RadiusY * GetDPIScale())
    If RadiusX < 1 Then RadiusX = 1
    If RadiusY < 1 Then RadiusY = 1
    
    ' YOUR BRILLIANT IDEA: If the control has a scrollbar, add its width back
    ' to the region's total width to avoid clipping it.
    If hasVerticalScrollbar Then
        PxWidth = PxWidth + GetSystemMetrics(SM_CXVSCROLL)
    End If
    
    ' 1. Create a sharp rectangle for the top part
    hRgn_TopRect = CreateRectRgn(0, 0, PxWidth, PxHeight - RadiusY)

    ' 2. Create a rounded rectangle for the full area
    hRgn_BottomRoundRect = CreateRoundRectRgn(0, 0, PxWidth, PxHeight, RadiusX, RadiusY)
    
    ' 3. Create an empty region to hold the result
    hRgn_Combined = CreateRectRgn(0, 0, 0, 0)
    
    ' 4. Combine the two regions
    CombineRgn hRgn_Combined, hRgn_TopRect, hRgn_BottomRoundRect, RGN_OR
    
    ' 5. Apply the final combined shape to the control's window
    SetWindowRgn ctl.hWnd, hRgn_Combined, True
    
    ' 6. Clean up
    DeleteObject hRgn_TopRect
    DeleteObject hRgn_BottomRoundRect
End Sub

''' <summary>
''' Uses the SetWindowPos API to bring a window handle to the top of the Z-Order
''' without moving or resizing it. This is more forceful than the ZOrder method.
''' </summary>
Public Sub BringWindowToTop(ByVal hWnd As LongPtr)
    SetWindowPos hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub ApplyUCStyle(ctrl As Control, styleName As String)
    ' Select Case styleName
    '     Case "Dialog"
    '         ApplyRoundedCorners ctrl, 6
    '         ApplyShadowEffect ctrl
    '     Case "Button"
    '         ApplyBottomRoundedCorners ctrl, 4
    '         ApplyHoverStyle ctrl
    ' End Select
End Sub
