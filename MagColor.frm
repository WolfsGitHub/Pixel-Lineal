VERSION 5.00
Begin VB.Form frmMagColor 
   BorderStyle     =   0  'Kein
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "MagColor.frx":0000
   MousePointer    =   99  'Benutzerdefiniert
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picMagColor 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   240
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmMagColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lScaleWidth As Long, lScaleHeight As Long
Private lDeltaX As Long, lDeltaY As Long
Private Const XFACTOR As Long = 8
Public PipColor As Long

Friend Sub PrintMagColor(tCursorPos As POINTAPI)
Dim lDeskDC As Long, lPxColor As Long
Dim i As Integer
Const DELTA_SHIFT As Long = 8
    Me.Move (tCursorPos.X - DELTA_SHIFT) * LTwipsPerPixelX, (tCursorPos.Y - DELTA_SHIFT) * LTwipsPerPixelY
    'Bild übertragen
    lDeskDC = GetDC(0&)
    lPxColor = GetPixel(lDeskDC, tCursorPos.X, tCursorPos.Y)
    StretchBlt picMagColor.hDC, 0, 0, lScaleWidth * XFACTOR, lScaleHeight * XFACTOR, _
      lDeskDC, tCursorPos.X - (XFACTOR \ 2), tCursorPos.Y - (XFACTOR \ 2), lScaleWidth, lScaleHeight, SRCCOPY
    ReleaseDC 0&, lDeskDC
    For i = 7 To 64 Step 8
        picMagColor.Line (0, i)-(71, i)
        picMagColor.Line (i, 0)-(i, 71)
    Next i
    picMagColor.Line (30, 30)-(40, 40), vbRed, B
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim deltaX As Long, deltaY As Long
Dim tCursorPos As POINTAPI
    Select Case KeyCode
        Case vbKeyEscape:   Me.Hide
        Case vbKeyLeft:     deltaX = -1
        Case vbKeyRight:    deltaX = 1
        Case vbKeyUp:       deltaY = -1
        Case vbKeyDown:     deltaY = 1
        Case vbKeyReturn
            SetPipColor
            Me.Hide
    End Select
    If deltaX <> 0 Or deltaY <> 0 Then
        GetCursorPos tCursorPos
        If SetCursorPos(tCursorPos.X + deltaX, tCursorPos.Y + deltaY) Then
            Me.Move (tCursorPos.X + deltaX) * LTwipsPerPixelX, (tCursorPos.Y + deltaY) * LTwipsPerPixelY
        End If
    End If
End Sub

Private Sub Form_Load()
    PipColor = &H1000000
    Me.Width = (picMagColor.Left + picMagColor.Width) * LTwipsPerPixelX
    Me.Height = (picMagColor.Top + picMagColor.Height) * LTwipsPerPixelY
    Me.BackColor = vbCyan
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetPipColor Button
    Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tCursorPos As POINTAPI
    GetCursorPos tCursorPos
    PrintMagColor tCursorPos
End Sub

Private Sub Form_Resize()
    lScaleWidth = Me.ScaleWidth
    lScaleHeight = Me.ScaleHeight
End Sub

Private Sub SetPipColor(Optional Button As Integer = vbLeftButton)
Dim tCursorPos As POINTAPI
Dim lDeskDC As Long
    lDeskDC = GetDC(0&)
    GetCursorPos tCursorPos
    If Button = vbLeftButton Then
        PipColor = GetPixel(lDeskDC, tCursorPos.X, tCursorPos.Y)
    ElseIf Button = vbRightButton Then
        PipColor = GetPixel(lDeskDC, tCursorPos.X, tCursorPos.Y) * -1
    End If
End Sub




