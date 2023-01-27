VERSION 5.00
Begin VB.Form frmMagGlass 
   AutoRedraw      =   -1  'True
   Caption         =   "Pixel-Lineal"
   ClientHeight    =   6825
   ClientLeft      =   3495
   ClientTop       =   3585
   ClientWidth     =   9360
   ClipControls    =   0   'False
   Icon            =   "MagGlass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   Begin VB.PictureBox picStatusbar 
      Align           =   2  'Unten ausrichten
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'Kein
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   624
      TabIndex        =   0
      Top             =   6510
      Width           =   9360
   End
   Begin VB.Menu mnuMagGlass 
      Caption         =   "Lupe"
      Begin VB.Menu mnuCopyRGB 
         Caption         =   "RGB-Wert kopieren"
      End
      Begin VB.Menu mnuColorCollection 
         Caption         =   "Gesammelte Farben"
         Begin VB.Menu mnuColorCollectionItems 
            Caption         =   "&&H00000000&&"
            Index           =   0
         End
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Nach Online-Updates suchen"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Schließen"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "Ansicht"
      Begin VB.Menu mnuColorCode 
         Caption         =   "HTML Farbanzeige"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuColorCode 
         Caption         =   "VB Farbanzeige"
         Index           =   1
      End
      Begin VB.Menu mnuColorCode 
         Caption         =   "OLEColor"
         Index           =   2
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFakor 
         Caption         =   "Vergrößerungsfaktor"
         Begin VB.Menu mnuFaktorX 
            Caption         =   "2"
            Index           =   0
         End
         Begin VB.Menu mnuFaktorX 
            Caption         =   "4"
            Index           =   1
         End
         Begin VB.Menu mnuFaktorX 
            Caption         =   "6"
            Index           =   2
         End
         Begin VB.Menu mnuFaktorX 
            Caption         =   "8"
            Index           =   3
         End
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuInformation 
      Caption         =   "&?"
      Begin VB.Menu mnuHelp 
         Caption         =   "Hilfe"
      End
      Begin VB.Menu mnuInternet 
         Caption         =   "Pixel-Lineal im Web"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "frmMagGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xFaktor As Integer
Const lStatusHeight& = 21

Private lScaleWidth As Long, lScaleHeight As Long
Private lDeltaX As Long, lDeltaY As Long
Private lReticuleX As Long, lReticuleY As Long
Private bStatus As Boolean


Public Sub Form_Resize()
    If WindowState > 1 Then Exit Sub
    lScaleWidth = Me.ScaleWidth
    lReticuleX = lScaleWidth \ 2
    lScaleHeight = Me.ScaleHeight
    lReticuleY = lScaleHeight \ 2
    lDeltaX = lScaleWidth / (xFaktor * 2)
    lDeltaY = lScaleHeight / (xFaktor * 2)
    lReticuleX = (lDeltaX * xFaktor) + ((xFaktor - 2) / 2)
    lReticuleY = (lDeltaY * xFaktor) + ((xFaktor - 2) / 2)
    
End Sub

Public Sub SetFactorX(newFaktor As Integer)
    mnuFaktorX(0).Checked = False
    mnuFaktorX(1).Checked = False
    mnuFaktorX(2).Checked = False
    mnuFaktorX(3).Checked = False
    xFaktor = newFaktor
    Select Case newFaktor
        Case 2: mnuFaktorX(0).Checked = True
        Case 4: mnuFaktorX(1).Checked = True
        Case 6: mnuFaktorX(2).Checked = True
        Case 8: mnuFaktorX(3).Checked = True
    End Select
    ForceRefresh = FORCE_REFRESH_RES
    Call Form_Resize
End Sub

Public Property Get StatusBarVisible() As Boolean
    StatusBarVisible = bStatus
End Property

Public Property Let StatusBarVisible(ByVal vNewValue As Boolean)
  On Error GoTo StatusBarVisible_Error
  bStatus = vNewValue
  mnuStatus.Checked = bStatus
  picStatusbar.Visible = bStatus
  ForceRefresh = FORCE_REFRESH_RES
Exit Property

StatusBarVisible_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMagGlass.StatusBarVisible." & Erl & vbCrLf & Err.Source, _
 vbCritical

End Property

Public Sub mnuHelp_Click()
    On Error GoTo mnuHelp_Click_Error
    ShellExec "https://docs.ww-a.de/doku.php/pixellineal:bildschirmlupe", vbNormalFocus
    Exit Sub
    
mnuHelp_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMagGlass.mnuHelp_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Friend Sub PrintMagGlass(tCursorPos As POINTAPI)
Dim lDeskDC As Long
Dim lPxColor As Long
Dim rectLeft As Long, rectTop As Long, rectRight As Long, rectBottom As Long
Dim deltaX As Long, deltaY As Long
deltaY = xFaktor \ 2
deltaX = deltaY - 1
    'Bild übertragen
    lDeskDC = GetDC(0&)
    lPxColor = GetPixel(lDeskDC, tCursorPos.X, tCursorPos.Y)
    StretchBlt Me.hDC, 0, 0, lScaleWidth * xFaktor, lScaleHeight * xFaktor, _
      lDeskDC, tCursorPos.X - lDeltaX, tCursorPos.Y - lDeltaY, lScaleWidth, lScaleHeight, SRCCOPY
    ReleaseDC 0&, lDeskDC
    'Fadenkreuz
    Me.Line (lReticuleX, 0)-(lReticuleX, lScaleHeight), vbBlack 'V
    Me.Line (0, lReticuleY)-(lScaleWidth, lReticuleY), vbBlack  'H
    'CaptureFenster
    If Not Capture Is Nothing Then
        With Capture
            rectLeft = lReticuleX + ((.Left \ LTwipsPerPixelX) - tCursorPos.X) * xFaktor
            rectTop = lReticuleY + ((.Top \ LTwipsPerPixelY) - tCursorPos.Y) * xFaktor
            rectRight = rectLeft + ((.Width \ LTwipsPerPixelX) * xFaktor)
            rectBottom = rectTop + ((.Height \ LTwipsPerPixelY) * xFaktor)
        End With
        Me.Line (rectLeft - deltaX, rectTop - deltaX)-(rectRight - deltaY, rectBottom - deltaY), vbRed, B
    End If
    PrintStatus lPxColor, tCursorPos
End Sub

Friend Sub PrintStatus(lPxColor As Long, tCursorPos As POINTAPI, Optional isCopy As Boolean)
Dim X As Long, Y As Long, a As Single, s As String

    If Not bStatus Then Exit Sub
    picStatusbar.Cls
    
    'XY-Feld
    picStatusbar.Line (0, 0)-(lScaleWidth, 0), vbButtonShadow                   'H-oben
    picStatusbar.Line (0, lStatusHeight)-(lScaleWidth - 1, lStatusHeight), vbWhite  'H-unten
    
    picStatusbar.CurrentX = 5: picStatusbar.CurrentY = 5
    If frmRuler.Visible Then
        X = Round((tCursorPos.X - (frmRuler.Left \ LTwipsPerPixelX) + plZeroLine) * Abs(RulerScaleMulti) * 1000) / 1000
        Y = Round((tCursorPos.Y - (frmRuler.Top \ LTwipsPerPixelY) + plZeroLine) * Abs(RulerScaleMulti) * 1000) / 1000
    Else
        X = Round(tCursorPos.X * Abs(RulerScaleMulti) * 1000) / 1000
        Y = Round(tCursorPos.Y * Abs(RulerScaleMulti) * 1000) / 1000
    End If
    a = Round(Math.Sqr(X ^ 2 + Y ^ 2), 1)
    s = "X:" & X & "   Y:" & Y & "   A:" & Format$(a, "0.0")
    picStatusbar.Print s
    X = picStatusbar.TextWidth(s) + 20
    If X < 140 Then X = 140 Else X = 160
    'XY-Farbe
    If lPxColor > -1 Then picStatusbar.Line (X, 4)-(X + 14, lStatusHeight - 4), lPxColor, BF
    picStatusbar.CurrentX = X + 24: picStatusbar.CurrentY = 5
    If isCopy Then
      picStatusbar.Print "Farbe kopiert!"
    Else
      If ColorCode = PL_HEXHTML Then
        picStatusbar.Print "Farbe: " & RGBtoHTML(lPxColor, True)
      ElseIf ColorCode = PL_HEXVB Then
        picStatusbar.Print "Farbe: " & RGBtoVB(lPxColor, True)
      Else
        picStatusbar.Print "Farbe: " & lPxColor & "   RGB: (" & CStr(lPxColor And vbRed) & "," & CStr((lPxColor And vbGreen) \ &H100) & "," & CStr((lPxColor And vbBlue) \ &H10000) & ")"
      End If
    End If
    
    'Gripper
    picStatusbar.Line (lScaleWidth - 11, lStatusHeight - 3)-(lScaleWidth - 10, lStatusHeight - 2), vbWhite, BF 'H
    picStatusbar.Line (lScaleWidth - 12, lStatusHeight - 4)-(lScaleWidth - 11, lStatusHeight - 3), &HA5B5BD, BF 'D
    picStatusbar.Line (lScaleWidth - 7, lStatusHeight - 7)-(lScaleWidth - 6, lStatusHeight - 6), vbWhite, BF 'H
    picStatusbar.Line (lScaleWidth - 8, lStatusHeight - 8)-(lScaleWidth - 7, lStatusHeight - 7), &HA5B5BD, BF 'D
    picStatusbar.Line (lScaleWidth - 3, lStatusHeight - 11)-(lScaleWidth - 2, lStatusHeight - 10), vbWhite, BF 'H
    picStatusbar.Line (lScaleWidth - 4, lStatusHeight - 12)-(lScaleWidth - 3, lStatusHeight - 11), &HA5B5BD, BF 'D
    
    picStatusbar.Line (lScaleWidth - 7, lStatusHeight - 3)-(lScaleWidth - 6, lStatusHeight - 2), vbWhite, BF 'H
    picStatusbar.Line (lScaleWidth - 8, lStatusHeight - 4)-(lScaleWidth - 7, lStatusHeight - 3), &HA5B5BD, BF 'D
    picStatusbar.Line (lScaleWidth - 3, lStatusHeight - 7)-(lScaleWidth - 2, lStatusHeight - 6), vbWhite, BF 'H
    picStatusbar.Line (lScaleWidth - 4, lStatusHeight - 8)-(lScaleWidth - 3, lStatusHeight - 7), &HA5B5BD, BF 'D
    
    picStatusbar.Line (lScaleWidth - 3, lStatusHeight - 3)-(lScaleWidth - 2, lStatusHeight - 2), vbWhite, BF 'H
    picStatusbar.Line (lScaleWidth - 4, lStatusHeight - 4)-(lScaleWidth - 3, lStatusHeight - 3), &HA5B5BD, BF 'D

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyDown_Error
    If KeyCode = vbKeyF1 Then
        ShellExec "https://docs.ww-a.de/doku.php/pixellineal:bildeditor", vbNormalFocus
    Else
        frmRuler.Form_KeyDown KeyCode, Shift
    End If
    Exit Sub
    
Form_KeyDown_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMagGlass.Form_KeyDown." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub Form_Load()
Dim l As Long, t As Long, w As Long, h As Long
Dim i As Integer
    picStatusbar.Height = lStatusHeight + 2
    For i = 1 To 15: Load mnuColorCollectionItems(i): mnuColorCollectionItems(i).Visible = False: Next i
    On Error Resume Next
    l = CLng(GetSetting(App.Title, "Options", "MagGlasLeft", 50))
    t = CLng(GetSetting(App.Title, "Options", "MagGlasTop", 50))
    w = CLng(GetSetting(App.Title, "Options", "MagGlasWidth", LScreenWidth \ 4))
    h = CLng(GetSetting(App.Title, "Options", "MagGlasHeight", LScreenHeight \ 4))
    bStatus = Not CBool(GetSetting(App.Title, "Options", "Status", 1))
    xFaktor = GetSetting(App.Title, "Options", "XFaktor", 6)
    Select Case xFaktor
      Case 2
        Call mnuFaktorX_Click(0)
      Case 4
        Call mnuFaktorX_Click(1)
      Case Else
        Call mnuFaktorX_Click(2)
    End Select
    Call mnuStatus_Click
    Call mnuColorCode_Click(CInt(ColorCode))
    If l < 0 Then l = 50
    If t < 0 Then t = 50
    SetWindowPos hwnd, HWND_TOPMOST, l, t, w, h, 0&
    mnuCopyRGB.Caption = "RGB-Wert kopieren" & Chr$(9) & "Strg+Alt+C"
    Call modMenuColor.Set_MenuColor(nfoMenuBarColor, Me.hwnd, &HF0F0F0)
    Call modMenuColor.Set_MenuColor(nfoSysMenuColor, Me.hwnd, &HF0F0F0)
    Call modMenuColor.Set_MenuColor(nfoMenuColor, Me.hwnd, &HF0F0F0, 0, True)
    Call modMenuColor.Set_MenuColor(nfoMenuColor, Me.hwnd, &HF0F0F0, 1, True)
    Call modMenuColor.Set_MenuColor(nfoMenuColor, Me.hwnd, &HF0F0F0, 2, True)
    frmMenu.mnuRMagGlass.Checked = True
    frmMenu.mnuSMagGlass.Checked = True
    h = GetMenu(Me.hwnd)
    h = GetSubMenu(h, 0&)
    SetMenuItemBitmaps h, 2, MF_BYPOSITION, frmMenu.picMenuFile(5).Picture, frmMenu.picMenuFile(5).Picture
    h = GetMenu(Me.hwnd)
    h = GetSubMenu(h, 2&)
    SetMenuItemBitmaps h, 0, MF_BYPOSITION, frmMenu.picMenuFile(4).Picture, frmMenu.picMenuFile(4).Picture
    SetMenuItemBitmaps h, 1, MF_BYPOSITION, frmMenu.picMenuFile(7).Picture, frmMenu.picMenuFile(7).Picture
    SetMenuItemBitmaps h, 2, MF_BYPOSITION, frmMenu.picMenuFile(8).Picture, frmMenu.picMenuFile(8).Picture


End Sub


Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    If Me.WindowState = vbNormal Then
      SaveSetting App.Title, "Options", "MagGlasLeft", Me.Left \ LTwipsPerPixelX
      SaveSetting App.Title, "Options", "MagGlasTop", Me.Top \ LTwipsPerPixelY
      SaveSetting App.Title, "Options", "MagGlasWidth", Me.Width \ LTwipsPerPixelX
      SaveSetting App.Title, "Options", "MagGlasHeight", Me.Height \ LTwipsPerPixelY
      SaveSetting App.Title, "Options", "Status", Abs(bStatus)
      SaveSetting App.Title, "Options", "XFaktor", xFaktor
    End If
    frmMenu.mnuRMagGlass.Checked = False
    frmMenu.mnuSMagGlass.Checked = False
    Set MagGlass = Nothing
End Sub

Private Sub mnuClose_Click()
  On Error Resume Next
  Unload Me
End Sub

Private Sub mnuColorCode_Click(Index As Integer)
    If Index = 0 Then
        mnuColorCode(0).Checked = True
        mnuColorCode(1).Checked = False
        mnuColorCode(2).Checked = False
        With frmMenu
            .mnuColorCode(0).Checked = True
            .mnuColorCode(1).Checked = False
            .mnuColorCode(2).Checked = False
        End With
    ElseIf Index = 1 Then
        mnuColorCode(0).Checked = False
        mnuColorCode(1).Checked = True
        mnuColorCode(2).Checked = False
        With frmMenu
            .mnuColorCode(0).Checked = False
            .mnuColorCode(1).Checked = True
            .mnuColorCode(2).Checked = False
        End With
    Else
        mnuColorCode(0).Checked = False
        mnuColorCode(1).Checked = False
        mnuColorCode(2).Checked = True
        With frmMenu
            .mnuColorCode(0).Checked = False
            .mnuColorCode(1).Checked = False
            .mnuColorCode(2).Checked = True
        End With
    End If
    ColorCode = Index
    On Error Resume Next
    SaveSetting App.Title, "Options", "ColorCode", ColorCode
End Sub

Private Sub mnuColorCollectionItems_Click(Index As Integer)
  On Error GoTo mnuColorCollectionItems_Click_Error
  Clipboard.Clear
  Clipboard.SetText mnuColorCollectionItems(Index).Caption, vbCFText
Exit Sub

mnuColorCollectionItems_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMagGlass.mnuColorCollectionItems_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuColorCollection_Click()
    On Error GoTo mnuColorCollection_Click_Error
    Call FillMenuColorCollection(Me, 1&) '1 = Position von Menü im MagGlass-Menü
Exit Sub

mnuColorCollection_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMagGlass.mnuColorCollection_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuCopyRGB_Click()
    On Error GoTo mnuCopyRGB_Click_Error
    CopyRGB GetPxColor
  Exit Sub
  
mnuCopyRGB_Click_Error:
  Screen.MousePointer = vbDefault
  MsgBox "Fehler: " & Err.Number & vbCrLf & _
   "Beschreibung: " & Err.Description & vbCrLf & _
   "Quelle: frmMagGlass.mnuCopyRGB_Click." & Erl & vbCrLf & Err.Source, _
   vbCritical
End Sub

Private Sub mnuFaktorX_Click(Index As Integer)
    Select Case Index
        Case 0: SetFactorX 2
        Case 1: SetFactorX 4
        Case 2: SetFactorX 6
        Case 3: SetFactorX 8
    End Select
End Sub

Private Sub mnuInfo_Click()
  MsgBox modMain.GetInfo, vbInformation, "Pixel-Lineal"
End Sub

Private Sub mnuInternet_Click()
    On Error GoTo mnuInternet_Click_Error
    ShellExec "https://www.ww-a.de/pixellineal.html", vbNormalFocus
    Exit Sub
    
mnuInternet_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMagGlass.mnuInternet_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub mnuMagGlass_Click()
    If ColorCollection(0) = -1 Then
      mnuColorCollection.Enabled = False
    Else
      mnuColorCollection.Enabled = True
    End If
End Sub

Private Sub mnuReset_Click()
    On Error GoTo mnuReset_Click_Error
    frmReset.Show vbModal, Me
Exit Sub

mnuReset_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMagGlass.mnuReset_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuStatus_Click()
    StatusBarVisible = Not StatusBarVisible
End Sub


Private Sub mnuUpdates_Click()
    Call modMain.CheckVersion
End Sub



