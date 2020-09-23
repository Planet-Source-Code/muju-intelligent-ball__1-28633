Attribute VB_Name = "MyModule"
Const HWND_BOTTOM = 1
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Const HWND_TOPMOST = -1

Const SWP_FRAMECHANGED = &H20
Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_HIDEWINDOW = &H80
Const SWP_NOACTIVATE = &H10
Const SWP_NOCOPYBITS = &H100
Const SWP_NOMOVE = &H2
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREDRAW = &H8
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_SHOWWINDOW = &H40

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Const Pi = 3.14159265358979 / 180

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

'------------------------------------

Type MyOwnAllWindowsDataInformationType
    WindowRect() As RECT
    TotalWindows As Long '--Total windows found
End Type

Type MyOwnAllBallDataInformationType
    BPosX As Double '--Ball PositionX
    BPosY As Double '--Ball PositionY
    BPosLX As Double '--Last Ball PositionX
    BPosLY As Double '--Last Ball PositionY
    BVelX As Double '--Acceleration X of the Ball
    BVelY As Double '--Acceleration Y of the Ball
    BSpeed As Double '--Speed of the ball
    BAngle As Double '--Angle of the ball
    BSize As Integer '--Size of the ball
    Walls As RECT '--Screen Restrictions
    FricReflect As Byte '--Percentage of friction when reflected on walls and ground
End Type

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public WinInfo As MyOwnAllWindowsDataInformationType
Public BallInfo As MyOwnAllBallDataInformationType

Public Aa As Long
Public Ab As Long
Public Ac As Long
Public StrA As String
Public TempRect As RECT
Public BackBuffer As Long
Public BackBitmap As Long
Public ScreenBuffer As Long
Public ScreenBitmap As Long
Public ScreenDc As Long
Public pRgnA As Long

Public DragPos As POINTAPI
Public DragPosL As POINTAPI

Sub Main()

'----Start
    SetPrimaryVariables
    GetAllWindowInfo

'----Creating BackBuffer
    'ScreenDc = GetDC(0)
    'ScreenBuffer = CreateCompatibleDC(ScreenDc)
    'ScreenBitmap = CreateCompatibleBitmap(ScreenDc, 1024, 768)
    'DeleteObject SelectObject(ScreenBuffer, ScreenBitmap)
    'BitBlt ScreenBuffer, 0, 0, 1024, 768, ScreenDc, 0, 0, SRCCOPY
    '--
    'BackBuffer = CreateCompatibleDC(ScreenDc)
    'BackBitmap = CreateCompatibleBitmap(ScreenDc, 1024, 768)
    'DeleteObject SelectObject(BackBuffer, BackBitmap)
    'SetBkMode BackBuffer, TRANSPARENT
    'SetTextColor BackBuffer, RGB(0, 255, 0)

'----Visualizing Form (Ball)
    pRgnA = CreateEllipticRgn(0, 0, BallInfo.BSize * 2, BallInfo.BSize * 2)
    MyForm.Show
    SetWindowRgn MyForm.hWnd, pRgnA, True
    DeleteObject pRgnA
    SetWindowPos MyForm.hWnd, HWND_TOPMOST, 0, 0, 200, 200, SWP_NOSIZE Or SWP_NOMOVE
    MoveWindow MyForm.hWnd, BallInfo.BPosX - (BallInfo.BSize / 2), BallInfo.BPosY - (BallInfo.BSize / 2), (BallInfo.BSize * 2), (BallInfo.BSize * 2), True
    

'----Animation Loop
    Do
        DoEvents
        'BitBlt BackBuffer, 0, 0, 1024, 768, BackBuffer, 0, 0, 0
        GetAllWindowInfo
        If DragPosL.X = 0 And DragPosL.Y = 0 Then AnimateBall
    Loop Until GetAsyncKeyState(27) < -1

'----Releasing Back Buffer
    'BitBlt ScreenDc, 0, 0, 1024, 768, ScreenBuffer, 0, 0, SRCCOPY
    'ReleaseDC 0, ScreenDc
    'DeleteDC ScreenBuffer
    'DeleteObject ScreenBitmap
    'DeleteDC BackBuffer
    'DeleteObject BackBitmap
    DeleteObject pRgnA
    End

End Sub

Sub GetAllWindowInfo()
'----Gets the information of all the Visible Windows
    Dim GWIAa As Long
    Dim GWIAb As Long
    
    Erase WinInfo.WindowRect
    WinInfo.TotalWindows = 0
    GWIAb = 0
    
    Do
        GWIAb = GWIAb + 1
        If GWIAb = 1 Then
            'GWIAa = GetWindow(MyForm.hwnd, GW_HWNDFIRST)
            GWIAa = GetWindow(GetTopWindow(0), GW_HWNDFIRST)
            Else
            GWIAa = GetWindow(GWIAa, GW_HWNDNEXT)
        End If
            
        If GWIAa = 0 Then Exit Do
        
        GetWindowRect GWIAa, TempRect
        If IsWindowVisible(GWIAa) And TempRect.Right - TempRect.Left > 0 And TempRect.Bottom - TempRect.Top > 0 _
            And Not (TempRect.Right = 1024 And TempRect.Left = 0 And TempRect.Bottom = 768 And TempRect.Top = 0) And TempRect.Top > 0 And TempRect.Right - TempRect.Left <> BallInfo.BSize * 2 Then
            WinInfo.TotalWindows = WinInfo.TotalWindows + 1
            ReDim Preserve WinInfo.WindowRect(WinInfo.TotalWindows - 1)
            WinInfo.WindowRect(WinInfo.TotalWindows - 1) = TempRect
            'Debug.Print "Windows ="; WinInfo.TotalWindows; GWIAb, TempRect.Left; TempRect.Top; TempRect.Right; TempRect.Bottom
        End If
    Loop
    
End Sub

Sub SetPrimaryVariables()
'----Sets the first time variables
    BallInfo.Walls.Left = 0
    BallInfo.Walls.Top = 0
    BallInfo.Walls.Right = 1024
    BallInfo.Walls.Bottom = 768
    BallInfo.BAngle = 0
    BallInfo.BSpeed = 1
    BallInfo.BSize = 50
    BallInfo.BPosX = (Screen.Width / Screen.TwipsPerPixelX) / 2 '512
    BallInfo.BPosY = (Screen.Height / Screen.TwipsPerPixelY) / 2 '384 + 200
    BallInfo.BVelX = 1
    BallInfo.BVelY = 5
    '--
    BallInfo.FricReflect = 80


End Sub

Sub AnimateBall()
'----Handles the control of the ball
    'Debug.Print "anim"
    '----Init Infomation
    BallInfo.BPosLX = BallInfo.BPosX
    BallInfo.BPosLY = BallInfo.BPosY
    
    '----Draw rectangles
    For Aa = 0 To WinInfo.TotalWindows - 1
        Rectangle BackBuffer, WinInfo.WindowRect(Aa).Left, WinInfo.WindowRect(Aa).Top, WinInfo.WindowRect(Aa).Right, WinInfo.WindowRect(Aa).Bottom
    Next Aa

    '----Ball Control
    BallInfo.BPosX = BallInfo.BPosX + BallInfo.BVelX
    BallInfo.BPosY = BallInfo.BPosY + BallInfo.BVelY
    
    '--Gravitiy
    BallInfo.BVelY = BallInfo.BVelY + 0.02
    
    '--Outer Wall Reflection
    If BallInfo.BPosY + BallInfo.BSize >= BallInfo.Walls.Bottom Then
        BallInfo.BVelY = -BallInfo.BVelY
        BallInfo.BVelY = BallInfo.BVelY * BallInfo.FricReflect / 100
        BallInfo.BVelX = BallInfo.BVelX * (BallInfo.FricReflect + 400) / 500 '--Special X-Axis Friction
        BallInfo.BPosX = BallInfo.BPosLX
        BallInfo.BPosY = BallInfo.BPosLY
    End If
    If BallInfo.BPosY - BallInfo.BSize <= BallInfo.Walls.Top Then
        BallInfo.BVelY = -BallInfo.BVelY
        BallInfo.BVelY = BallInfo.BVelY * BallInfo.FricReflect / 100
        BallInfo.BPosX = BallInfo.BPosLX
        BallInfo.BPosY = BallInfo.BPosLY
    End If
    If BallInfo.BPosX + BallInfo.BSize >= BallInfo.Walls.Right Then
        BallInfo.BVelX = -BallInfo.BVelX
        BallInfo.BVelX = BallInfo.BVelX * BallInfo.FricReflect / 100
        BallInfo.BPosX = BallInfo.BPosLX
        BallInfo.BPosY = BallInfo.BPosLY
    End If
    If BallInfo.BPosX - BallInfo.BSize <= BallInfo.Walls.Left Then
        BallInfo.BVelX = -BallInfo.BVelX
        BallInfo.BVelX = BallInfo.BVelX * BallInfo.FricReflect / 100
        BallInfo.BPosX = BallInfo.BPosLX
        BallInfo.BPosY = BallInfo.BPosLY
    End If

    '----Collition With other Forms
    For Aa = 0 To WinInfo.TotalWindows - 1
        '--Top
        If BallInfo.BVelY > 0 And (BallInfo.BPosX + BallInfo.BSize >= WinInfo.WindowRect(Aa).Left And BallInfo.BPosX - BallInfo.BSize <= WinInfo.WindowRect(Aa).Right) And WinInfo.WindowRect(Aa).Top >= BallInfo.BPosY - BallInfo.BSize And WinInfo.WindowRect(Aa).Top <= BallInfo.BPosY + BallInfo.BSize Then
            BallInfo.BVelY = -BallInfo.BVelY
            BallInfo.BVelY = BallInfo.BVelY * BallInfo.FricReflect / 100
            BallInfo.BVelX = BallInfo.BVelX * (BallInfo.FricReflect + 400) / 500 '--Special X-Axis Friction
            BallInfo.BPosX = BallInfo.BPosLX
            BallInfo.BPosY = BallInfo.BPosLY
        End If
        '--Bottom
        If BallInfo.BVelY < 0 And (BallInfo.BPosX + BallInfo.BSize >= WinInfo.WindowRect(Aa).Left And BallInfo.BPosX - BallInfo.BSize <= WinInfo.WindowRect(Aa).Right) And WinInfo.WindowRect(Aa).Bottom >= BallInfo.BPosY - BallInfo.BSize And WinInfo.WindowRect(Aa).Bottom <= BallInfo.BPosY + BallInfo.BSize Then
            BallInfo.BVelY = -BallInfo.BVelY
            BallInfo.BVelY = BallInfo.BVelY * BallInfo.FricReflect / 100
            BallInfo.BPosX = BallInfo.BPosLX
            BallInfo.BPosY = BallInfo.BPosLY
        End If
        '--Left
        If BallInfo.BVelX > 0 And (BallInfo.BPosY + BallInfo.BSize >= WinInfo.WindowRect(Aa).Top And BallInfo.BPosY - BallInfo.BSize <= WinInfo.WindowRect(Aa).Bottom) And WinInfo.WindowRect(Aa).Left >= BallInfo.BPosX - BallInfo.BSize And WinInfo.WindowRect(Aa).Left <= BallInfo.BPosX + BallInfo.BSize Then
            BallInfo.BVelX = -BallInfo.BVelX
            BallInfo.BVelX = BallInfo.BVelX * BallInfo.FricReflect / 100
            BallInfo.BPosX = BallInfo.BPosLX
            BallInfo.BPosY = BallInfo.BPosLY
        End If
        '--Right
        If BallInfo.BVelX < 0 And (BallInfo.BPosY + BallInfo.BSize >= WinInfo.WindowRect(Aa).Top And BallInfo.BPosY - BallInfo.BSize <= WinInfo.WindowRect(Aa).Bottom) And WinInfo.WindowRect(Aa).Right >= BallInfo.BPosX - BallInfo.BSize And WinInfo.WindowRect(Aa).Right <= BallInfo.BPosX + BallInfo.BSize Then
            BallInfo.BVelX = -BallInfo.BVelX
            BallInfo.BVelX = BallInfo.BVelX * BallInfo.FricReflect / 100
            BallInfo.BPosX = BallInfo.BPosLX
            BallInfo.BPosY = BallInfo.BPosLY
        End If
    Next Aa


    '--DrawBall
    'Ellipse BackBuffer, BallInfo.BPosX - BallInfo.BSize, BallInfo.BPosY - BallInfo.BSize, BallInfo.BPosX + BallInfo.BSize, BallInfo.BPosY + BallInfo.BSize
    'TextOut BackBuffer, BallInfo.BPosX, BallInfo.BPosY, "Shai", 4
    MoveWindow MyForm.hWnd, BallInfo.BPosX - BallInfo.BSize, BallInfo.BPosY - BallInfo.BSize, BallInfo.BSize * 2, BallInfo.BSize * 2, True

    '----Draw to screen
    'BitBlt ScreenDc, 0, 0, 1024, 768, BackBuffer, 0, 0, SRCCOPY
    'BitBlt BackBuffer, 0, 0, 1024, 768, BackBuffer, 0, 0, 0
    'For Delay = 0 To 10000: Next Delay
End Sub
