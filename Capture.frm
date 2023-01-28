VERSION 5.00
Begin VB.Form frmCapture 
   Appearance      =   0  '2D
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'Kein
   ClientHeight    =   3360
   ClientLeft      =   11130
   ClientTop       =   6780
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   3
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Capture.frx":0000
   MousePointer    =   15  'Größenänderung alle
   ScaleHeight     =   3360
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timAlpha 
      Interval        =   150
      Left            =   2145
      Top             =   495
   End
   Begin VB.PictureBox picToSave 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1215
      HelpContextID   =   3
      Left            =   240
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "w×h"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3465
      TabIndex        =   1
      Top             =   2805
      Width           =   435
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H000040C0&
      Height          =   1830
      Left            =   30
      Top             =   30
      Width           =   2985
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Declare Function GetDC Lib "user32" ( _
    ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hDC As Long) As Long
    
Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1

'####Transparenz####
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" ( _
                 ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE As Long = -20&
Private Const WS_EX_LAYERED As Long = &H80000


'###Fenster bewegen###
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2



Private Sub Form_DblClick()
Dim lDeskDC As Long, hDeskDC As Long
Dim f As frmImage
Dim isRetry As Boolean
    On Error GoTo Form_DblClick_Error
    Me.Hide
    hDeskDC = GetDesktopWindow()
    lDeskDC = GetDC(hDeskDC)
    Sleep 100
    With picToSave
        .Width = Me.Width
        .Height = Me.Height
        .AutoRedraw = True
        .Cls
         StretchBlt picToSave.hDC, 0, 0, Me.ScaleX(.Width, vbTwips, vbPixels), Me.ScaleY(.Height, vbTwips, vbPixels), _
            lDeskDC, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(.Width, vbTwips, vbPixels), Me.ScaleY(.Height, vbTwips, vbPixels), SRCCOPY
    End With
    ReleaseDC lDeskDC, lDeskDC
    If GetAsyncKeyState(vbKeyShift) Then
Retry_Copy:
        Clipboard.Clear
        Clipboard.SetData picToSave.Image, vbCFDIB
        If Err Then
            Err.Clear
            On Error GoTo Form_DblClick_Error
            Sleep 500
            Clipboard.Clear
            Clipboard.SetData picToSave.Image, vbCFDIB
        End If
        If Not Capture Is Nothing Then Unload Capture
    Else
        Set f = New frmImage
        f.ShowCapture Me.Left, Me.Top, Me.Width, Me.Height, picToSave.Image
        If Not Capture Is Nothing Then Unload Capture
    End If
Exit Sub

Form_DblClick_Error:
Screen.MousePointer = vbDefault
If lDeskDC <> 0 Or hDeskDC <> 0 Then ReleaseDC lDeskDC, lDeskDC
If Err = 521 And Not isRetry Then
    If MsgBox("Fehler: " & Err.Number & vbCrLf & Err.Description, vbInformation Or vbRetryCancel) = vbRetry Then
        isRetry = True
        Resume Retry_Copy
    End If
Else
    Me.Show
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmCapture.Form_DblClick." & Erl & vbCrLf & Err.Source, _
     vbCritical
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tCursorPos As POINTAPI
    On Error GoTo Form_KeyDown_Error
    Select Case KeyCode
        Case vbKeyLeft:     If Shift = vbShiftMask Then Width = Abs(Width - LTwipsPerPixelX) Else Left = Left - LTwipsPerPixelX
        Case vbKeyRight:    If Shift = vbShiftMask Then Width = Abs(Width + LTwipsPerPixelX) Else Left = Left + LTwipsPerPixelX
        Case vbKeyUp:       If Shift = vbShiftMask Then Height = Abs(Height - LTwipsPerPixelX) Else Top = Top - LTwipsPerPixelY
        Case vbKeyDown:     If Shift = vbShiftMask Then Height = Abs(Height + LTwipsPerPixelX) Else Top = Top + LTwipsPerPixelY
        Case vbKeyEscape
            Unload Capture
            Set Capture = Nothing
        Case vbKeyReturn, vbKeySpace
            Call Form_DblClick
            Exit Sub
        Case vbKeyShift
            If MousePointer = vbSizeAll Then MousePointer = vbCustom
        Case vbKeyF1
            ShellExec "https://docs.ww-a.de/doku.php/pixellineal:screenshot", vbNormalFocus
            Exit Sub
    End Select
    If Not MagGlass Is Nothing Then
        DoEvents
        GetCursorPos tCursorPos
        MagGlass.PrintMagGlass tCursorPos
    End If
Exit Sub

Form_KeyDown_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmCapture.Form_KeyDown." & Erl & vbCrLf & Err.Source, _
 vbCritical
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then Form_DblClick
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift And vbKeyShift Then
        If MousePointer = vbSizeAll Then MousePointer = vbCustom
    End If
End Sub

Private Sub Form_Load()
Dim X As Long, Y As Long, w As Long, h As Long

    On Error Resume Next
    w = Abs(CInt(GetSetting(App.Title, "ScreenShot", "Width", 400)))
    h = Abs(CInt(GetSetting(App.Title, "ScreenShot", "Height", 300)))
    If w < 16 Then w = 16
    If h < 16 Then h = 16
    If w > Screen.Width \ LTwipsPerPixelX Then w = Screen.Width \ LTwipsPerPixelX
    If h > Screen.Height \ LTwipsPerPixelY Then h = Screen.Height \ LTwipsPerPixelY
    X = (Me.ScaleX(Screen.Width, vbTwips, vbPixels) - w) \ 2
    Y = (Me.ScaleX(Screen.Height, vbTwips, vbPixels) - h) \ 2
    SetWindowPos hwnd, HWND_TOPMOST, X, Y, w, h, 0&
    frmMenu.mnuScreenShot.Checked = True
    lblSize.Visible = frmMenu.mnuShowSize.Checked

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MousePosition As MousePos

    If Button = vbLeftButton Then
        MousePosition = GetMousePos(Me, X, Y)
        ReleaseCapture
        Select Case MousePosition
        Case mpTopLeft:      PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_TopLeft, 0&
        Case mpBottomLeft:   PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_BottomLeft, 0&
        Case mpTopRight:     PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_TopRight, 0&
        Case mpBottomRight:  PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_BottomRight, 0&
        Case mpLeft:         PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_Left, 0&
        Case mpTop:          PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_Top, 0&
        Case mpRight:        PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_Right, 0&
        Case mpBottom:       PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_Bottom, 0&
        Case Else:           SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        End Select
    ElseIf Button = vbRightButton Then
        frmMenu.PopupMenu frmMenu.MScreenShot
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MousePosition As MousePos
    MousePosition = GetMousePos(Me, X, Y)
    Select Case MousePosition
        Case mpTopLeft:     MousePointer = vbSizeNWSE
        Case mpBottomLeft:  MousePointer = vbSizeNESW
        Case mpTopRight:    MousePointer = vbSizeNESW
        Case mpBottomRight: MousePointer = vbSizeNWSE
        Case mpLeft:        MousePointer = vbSizeWE
        Case mpTop:         MousePointer = vbSizeNS
        Case mpRight:       MousePointer = vbSizeWE
        Case mpBottom:      MousePointer = vbSizeNS
        Case Else:          If GetAsyncKeyState(vbKeyShift) Then MousePointer = vbCustom Else MousePointer = vbSizeAll
    End Select
    
End Sub

Private Sub Form_Resize()
Dim w As Long, h As Long
    On Error Resume Next
    w = ScaleWidth
    h = ScaleHeight
    If w < 180 Then
        w = 180
        Width = w
    End If
    If h < 180 Then
        h = 180
        Height = h
    End If
    lblSize.Caption = (w \ LTwipsPerPixelX) & "×" & (h \ LTwipsPerPixelY)
    If w < 600 Or h < 600 Then
        lblSize.Visible = False
    Else
        lblSize.Visible = True
        lblSize.Move w - lblSize.Width - (5 * LTwipsPerPixelX), (4 * LTwipsPerPixelX)
    End If
    shBorder.Move 0, 0, w, h
    
End Sub

Private Sub Form_Unload(cancel As Integer)
Dim f As Form
    On Error Resume Next
    SaveSetting App.Title, "ScreenShot", "Width", Me.Width \ LTwipsPerPixelX
    SaveSetting App.Title, "ScreenShot", "Height", Me.Height \ LTwipsPerPixelY
    frmMenu.mnuScreenShot.Checked = False
    Set Capture = Nothing
    
    If frmRuler.Visible Then
        Exit Sub
    Else
        For Each f In Forms 'Prüfen ob die Anwendung geschlossen werden kann
            If TypeOf f Is frmImage Then Exit Sub
        Next
        Set f = Nothing
        modMain.CloseApp = True
    End If
    
End Sub


Private Sub timAlpha_Timer()
Static WinInfo As Long
Static alphaValue As Long
    alphaValue = alphaValue + 40&
    If alphaValue > 180& Then
        timAlpha.Enabled = False
        Exit Sub
    End If
    If WinInfo = 0 Then
        WinInfo = GetWindowLong(hwnd, GWL_EXSTYLE)
        WinInfo = WinInfo Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, WinInfo
        SetCursorPos (Me.Left + (Me.Width \ 2)) \ LTwipsPerPixelX, (Me.Top + (Me.Height \ 2)) \ LTwipsPerPixelY
    End If
    SetLayeredWindowAttributes hwnd, 0&, 255& - alphaValue, LWA_ALPHA
End Sub






