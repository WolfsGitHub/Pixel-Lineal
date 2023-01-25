VERSION 5.00
Begin VB.Form frmRuler 
   BorderStyle     =   0  'Kein
   ClientHeight    =   510
   ClientLeft      =   8685
   ClientTop       =   5445
   ClientWidth     =   3285
   ControlBox      =   0   'False
   Icon            =   "Ruler.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   219
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerMagGlass 
      Interval        =   20
      Left            =   2475
      Top             =   0
   End
   Begin VB.PictureBox picRuler 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   0
      MouseIcon       =   "Ruler.frx":000C
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MarkerColor As Long

Private Declare Function CreateFont Lib "gdi32" Alias _
        "CreateFontA" (ByVal h As Long, ByVal w As Long, _
        ByVal e As Long, ByVal O As Long, ByVal w As _
        Long, ByVal i As Long, ByVal u As Long, ByVal s _
        As Long, ByVal c As Long, ByVal OP As Long, ByVal _
        CP As Long, ByVal Q As Long, ByVal PAF As Long, _
        ByVal f As String) As Long
 
Private Declare Function SelectObject Lib "gdi32" (ByVal _
        hDC As Long, ByVal hObject As Long) As Long
 
Private Declare Function DeleteObject Lib "gdi32" (ByVal _
        hObject As Long) As Long
 
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
        (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal lpString As String, ByVal nCount As Long) As Long
        
Private Const VK_LBUTTON = &H1
'Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4

Private mRulerOrientation As PL_Orientation
Private mRulerWidth As Single
Private mRulerHeight As Single
Private mRedrawRequired As Boolean



Public Sub DrawLabelingScaleUser()
  Dim txtMessage As String
  Dim oldSize As Long, hFont As Long, fontMem As Long, bold As Long, res As Long
  
  oldSize = picRuler.FontSize
  picRuler.FontSize = oldSize + 2
  txtMessage = "Bei gerdückter Strg-Taste auf Lineal klicken um Referenzmaß einzugeben."
  If mRulerOrientation = PL_HORIZONTAL Then
    picRuler.CurrentX = 20: picRuler.CurrentY = 6: picRuler.Print txtMessage;
    picRuler.Line (0, 0)-(0, plBREADTH), MarkerColor
  ElseIf mRulerOrientation = PL_VERTICAL Then
    Me.AutoRedraw = True
    hFont = CreateFont(oldSize + 6, 0, -900, 0, bold, _
            picRuler.FontItalic, picRuler.FontUnderline, 0, 1, 4, &H10, _
            2, 4, picRuler.FontName)
    fontMem = SelectObject(picRuler.hDC, hFont)
    res = TextOut(picRuler.hDC, 14, 10, txtMessage, Len(txtMessage))
    res = SelectObject(picRuler.hDC, fontMem)
    res = DeleteObject(hFont)
    picRuler.Line (0, 0)-(plBREADTH, 0), MarkerColor
    Me.Refresh
  End If
    picRuler.FontSize = oldSize
End Sub


Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    On Error GoTo Form_KeyDown_Error
    Select Case KeyCode
        Case vbKeyLeft:     Left = Left - LTwipsPerPixelX
        Case vbKeyRight:    Left = Left + LTwipsPerPixelX
        Case vbKeyUp:       Top = Top - LTwipsPerPixelY
        Case vbKeyDown:     Top = Top + LTwipsPerPixelY
        Case vbKeyM
            If mRulerOrientation = PL_HORIZONTAL Then
                For i = 1 To UBound(HMarker)
                    If TMarker = HMarker(i) Or TMarker = HMarker(i) - 1 Or TMarker = HMarker(i) + 1 Then
                        RemoveMarker i
                        GoTo REFRESH_MAG_GLASS
                    End If
                Next
            Else
                For i = 1 To UBound(VMarker)
                    If TMarker = VMarker(i) Or TMarker = VMarker(i) - 1 Or TMarker = VMarker(i) + 1 Then
                        RemoveMarker i
                        GoTo REFRESH_MAG_GLASS
                    End If
                Next
            End If
            SetMarker
        Case vbKeyS
            If Capture Is Nothing Then
                Set Capture = New frmCapture
                Capture.Show vbModeless, Me
            Else
                Unload Capture
                Set Capture = Nothing
            End If
        Case vbKeyF1
            ShellExec "https://docs.ww-a.de/doku.php/pixellineal:start", vbNormalFocus
            Exit Sub
        Case Else: Exit Sub
    End Select
REFRESH_MAG_GLASS:
    If Not MagGlass Is Nothing Then
        ForceRefresh = FORCE_REFRESH_RES
        TimerMagGlass_Timer
    End If
Exit Sub

Form_KeyDown_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmRuler.Form_KeyDown." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub


Public Property Get Orientation() As PL_Orientation
  Orientation = mRulerOrientation
End Property

Public Property Let Orientation(ByVal vNewValue As PL_Orientation)
  Dim tCursorPos As POINTAPI
  Dim tDeltaPos As POINTAPI
  Dim X As Long, Y As Long
  
    mRulerOrientation = vNewValue
    GetCursorPos tCursorPos
    tDeltaPos.X = tCursorPos.X * LTwipsPerPixelX - Me.Left
    tDeltaPos.Y = tCursorPos.Y * LTwipsPerPixelY - Me.Top
  
    If mRulerOrientation = PL_HORIZONTAL Then
        frmMenu.mnuOrientation.Caption = "&Vertikal"
        X = (tCursorPos.X * LTwipsPerPixelX) - tDeltaPos.Y
        Y = (tCursorPos.Y * LTwipsPerPixelY) - tDeltaPos.X
    Else
        frmMenu.mnuOrientation.Caption = "&Horizontal"
        X = (tCursorPos.X * LTwipsPerPixelX) - tDeltaPos.Y
        Y = (tCursorPos.Y * LTwipsPerPixelY) - tDeltaPos.X
    End If
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    Call ProcRefreshRuler(X, Y)
    X = GetMenu(frmMenu.hwnd)
    Y = GetSubMenu(X, 0&)
    SetMenuItemBitmaps Y, 0, MF_BYPOSITION, frmMenu.picMenuRuler(Abs(mRulerOrientation - 1)).Picture, frmMenu.picMenuRuler(Abs(mRulerOrientation - 1)).Picture
End Property


Public Sub ProcRefreshRuler(X As Long, Y As Long)
Dim i As Long
Dim iZehner As Long
Dim ubMarker As Long
Dim sBeschriftung As String

    picRuler.Cls
    If mRulerOrientation = PL_HORIZONTAL Then
      Me.Move X, Y, mRulerWidth * LTwipsPerPixelX, plBREADTH * LTwipsPerPixelY    'Breite des Lineals berechnen
      picRuler.Move 0, 0, mRulerWidth, plBREADTH
      For i = 2 To mRulerWidth Step 2    'kleine Gradierungen setzen, werden bei allen Einstellungen benötigt
          picRuler.Line (i - plZeroLine, 0)-(i - plZeroLine, 2)
      Next i
    Else
      Me.Move X, Y, plBREADTH * LTwipsPerPixelY, mRulerHeight * LTwipsPerPixelY  'Höhe des Lineals berechnen
      picRuler.Move 0, 0, plBREADTH, mRulerHeight
      For i = 2 To mRulerHeight Step 2   'kleine Gradierungen setzen, werden bei allen Einstellungen benötigt
          picRuler.Line (plBREADTH - 2, i - plZeroLine)-(plBREADTH, i - plZeroLine)
      Next i
    End If
        
    Select Case RulerScaleMode
    Case PL_PIXEL
        If mRulerOrientation = PL_HORIZONTAL Then
            For iZehner = 10 To mRulerWidth Step 10
              If iZehner Mod 100 <> 0 Then  '5-er und 10-er setzen, bei 100-er siehe else
                picRuler.Line (iZehner - plZeroLine, 0)-(iZehner - plZeroLine, 6)
                picRuler.Line (iZehner - plZeroLine - 5, 2)-(iZehner - plZeroLine - 5, 4)
                picRuler.CurrentX = iZehner - plZeroLine - 2
                picRuler.CurrentY = 5
                sBeschriftung = Left$(Right$(CStr(iZehner), 2), 1)
                picRuler.Print sBeschriftung;
              Else  'bei 100-er, langer Strich, Text tiefer
                picRuler.Line (iZehner - plZeroLine, 0)-(iZehner - plZeroLine, 10)
                picRuler.Line (iZehner - plZeroLine - 5, 2)-(iZehner - plZeroLine - 5, 4)
                sBeschriftung = CStr(iZehner)
                picRuler.CurrentX = iZehner - picRuler.TextWidth(sBeschriftung) \ 2
                picRuler.CurrentY = 10
                picRuler.Print sBeschriftung;
              End If
            Next iZehner
            'Marker
            picRuler.Line (0, 0)-(0, plBREADTH), MarkerColor
            ubMarker = UBound(HMarker)
            For i = 1 To ubMarker
                picRuler.Line (HMarker(i), 0)-(HMarker(i), plBREADTH), MarkerColor
            Next i
        ElseIf mRulerOrientation = PL_VERTICAL Then
            For iZehner = 10 To mRulerHeight Step 10
              If iZehner Mod 100 <> 0 Then '5-er und 10-er setzen, bei 100-er siehe else
                picRuler.Line (plBREADTH - 6, iZehner - plZeroLine)-(plBREADTH, iZehner - plZeroLine)
                picRuler.Line (plBREADTH - 3, iZehner - plZeroLine - 5)-(plBREADTH - 5, iZehner - plZeroLine - 5)
                sBeschriftung = Left$(Right$(CStr(iZehner), 2), 1)
                picRuler.CurrentX = 15 - picRuler.TextWidth(sBeschriftung)
                picRuler.CurrentY = iZehner - plZeroLine - 4
                picRuler.Print Left$(Right$(CStr(iZehner), 2), 1);
              Else 'bei 100-er, langer Strich, Text weiter rechts
                picRuler.Line (plBREADTH - 7, iZehner - plZeroLine)-(plBREADTH, iZehner - plZeroLine)
                picRuler.CurrentX = 0
                picRuler.CurrentY = iZehner - 4
                picRuler.Print iZehner;
              End If
            Next iZehner
            'Marker
            picRuler.Line (0, 0)-(plBREADTH, 0), MarkerColor
            ubMarker = UBound(VMarker)
            For i = 1 To ubMarker
                picRuler.Line (0, VMarker(i))-(plBREADTH, VMarker(i)), MarkerColor
            Next i
        End If
    Case PL_TWIPS
        If mRulerOrientation = PL_HORIZONTAL Then
            'optimale Beschriftung berechnen
            For i = RulerScaleMulti * 10 To mRulerWidth * RulerScaleMulti Step RulerScaleMulti
              If i Mod 100 = 0 Then Exit For
            Next i
            picRuler.CurrentX = 2: picRuler.CurrentY = 13: picRuler.Print "x100";
            i = i / RulerScaleMulti
            For iZehner = i To mRulerWidth Step i
                picRuler.Line (iZehner - plZeroLine, 0)-(iZehner - plZeroLine, 6)
                picRuler.CurrentY = 5
                sBeschriftung = CStr((iZehner * RulerScaleMulti) \ 100)
                picRuler.CurrentX = iZehner - picRuler.TextWidth(sBeschriftung) \ 2
                picRuler.Print sBeschriftung;
            Next iZehner
            'Marker
            picRuler.Line (0, 0)-(0, plBREADTH), MarkerColor
            ubMarker = UBound(HMarker)
            For i = 1 To ubMarker
                picRuler.Line (HMarker(i), 0)-(HMarker(i), plBREADTH), MarkerColor
            Next i
        ElseIf mRulerOrientation = PL_VERTICAL Then
            'optimale Beschriftung berechnen
            For i = RulerScaleMulti * 10 To mRulerHeight * RulerScaleMulti Step RulerScaleMulti
              If i Mod 100 = 0 Then Exit For
            Next i
            picRuler.CurrentX = 1: picRuler.CurrentY = 1: picRuler.Print "x100";
            i = i / RulerScaleMulti
            For iZehner = i To mRulerHeight Step i
                picRuler.Line (plBREADTH - 6, iZehner - plZeroLine)-(plBREADTH, iZehner - plZeroLine)
                sBeschriftung = CStr(Fix(iZehner * RulerScaleMulti) \ 100)
                picRuler.CurrentX = 15 - picRuler.TextWidth(sBeschriftung)
                picRuler.CurrentY = iZehner - plZeroLine - 4
                picRuler.Print sBeschriftung;
    
            Next iZehner
            'Marker
            picRuler.Line (0, 0)-(plBREADTH, 0), MarkerColor
            ubMarker = UBound(VMarker)
            For i = 1 To ubMarker
                picRuler.Line (0, VMarker(i))-(plBREADTH, VMarker(i)), MarkerColor
            Next i
        End If
    Case PL_USER
        If RulerScaleMulti = -1 Then  'Benutzerdefinierter Maßstab ist noch nicht festgelegt
              DrawLabelingScaleUser
              Exit Sub
        End If
    
        If mRulerOrientation = PL_HORIZONTAL Then
            For iZehner = 10 To mRulerWidth Step 10
              If iZehner Mod 100 <> 0 Then
                picRuler.Line (iZehner - plZeroLine, 0)-(iZehner - plZeroLine, 6)
                picRuler.Line (iZehner - plZeroLine - 5, 2)-(iZehner - plZeroLine - 5, 4)
                picRuler.CurrentX = iZehner - plZeroLine - 2
                picRuler.CurrentY = 5
              Else
                picRuler.Line (iZehner - plZeroLine, 0)-(iZehner - plZeroLine, 10)
                picRuler.Line (iZehner - plZeroLine - 5, 2)-(iZehner - plZeroLine - 5, 4)
                sBeschriftung = CStr(((iZehner * RulerScaleMulti * 100) \ 1) / 100)
                picRuler.CurrentX = iZehner - picRuler.TextWidth(sBeschriftung) \ 2
                picRuler.CurrentY = 10
                picRuler.Print sBeschriftung;
              End If
            Next iZehner
            'Marker
            picRuler.Line (0, 0)-(0, plBREADTH), MarkerColor
            ubMarker = UBound(HMarker)
            For i = 1 To ubMarker
                picRuler.Line (HMarker(i), 0)-(HMarker(i), plBREADTH), MarkerColor
            Next i
        ElseIf mRulerOrientation = PL_VERTICAL Then
            For iZehner = 10 To mRulerHeight Step 10
              If iZehner Mod 100 <> 0 Then
                picRuler.Line (plBREADTH - 6, iZehner - plZeroLine)-(plBREADTH, iZehner - plZeroLine)
                picRuler.Line (plBREADTH - 3, iZehner - plZeroLine - 5)-(plBREADTH - 5, iZehner - plZeroLine - 5)
              Else
                picRuler.Line (plBREADTH - 7, iZehner - plZeroLine)-(plBREADTH, iZehner - plZeroLine)
                picRuler.CurrentX = 0
                picRuler.CurrentY = iZehner - 4
                picRuler.Print CStr(((iZehner * RulerScaleMulti * 100) \ 1) / 100);
              End If
            Next iZehner
            'Marker
            picRuler.Line (0, 0)-(plBREADTH, 0), MarkerColor
            ubMarker = UBound(VMarker)
            For i = 1 To ubMarker
                picRuler.Line (0, VMarker(i))-(plBREADTH, VMarker(i)), MarkerColor
            Next i
        End If
    End Select
End Sub

Public Sub RemoveMarker(i As Integer)
Dim j As Long, ubMarker As Long
    If mRulerOrientation = PL_HORIZONTAL Then
        ubMarker = UBound(HMarker) - 1
        For j = i To ubMarker
            HMarker(j) = HMarker(j + 1)
        Next j
        ReDim Preserve HMarker(ubMarker)
    Else
        ubMarker = UBound(VMarker) - 1
        For j = i To ubMarker
            VMarker(j) = VMarker(j + 1)
        Next j
        ReDim Preserve VMarker(ubMarker)
    End If

    ProcRefreshRuler Left, Top
    TMarker = 0

End Sub

Public Sub SetMarker()

    If mRulerOrientation = PL_HORIZONTAL Then
        ReDim Preserve HMarker(UBound(HMarker) + 1)
        HMarker(UBound(HMarker)) = TMarker
    Else    'Vertikal
        ReDim Preserve VMarker(UBound(VMarker) + 1)
        VMarker(UBound(VMarker)) = TMarker
    End If
    ProcRefreshRuler Left, Top
    TMarker = 0

End Sub

Private Sub Form_Load()
Dim lCurrentStyle As Long
Dim iTransparencyRuler As Integer
Dim startCmd As String

    On Error GoTo Form_Load_Error
    Load frmMenu
    If CloseApp Then
        Unload Me
        Exit Sub
    End If
    LTwipsPerPixelX = Screen.TwipsPerPixelX
    LTwipsPerPixelY = Screen.TwipsPerPixelY
    LScreenWidth = Screen.Width \ LTwipsPerPixelY
    LScreenHeight = Screen.Height \ LTwipsPerPixelY
    mRulerWidth = LScreenWidth
    mRulerHeight = LScreenHeight
    
    ReDim ColorCollection(0)
    ReDim HMarker(0)
    ReDim VMarker(0)
    ColorCollection(0) = -1
    XYFieldWidth = XYFieldMinWidth
    
    On Error Resume Next
    MarkerColor = CLng(GetSetting(App.Title, "Options", "MarkColor", RGB(255, 0, 0)))
    With picRuler
        With .Font
            .Name = GetSetting(App.Title, "Options", "FontName", "Arial")
            .bold = GetSetting(App.Title, "Options", "FontBold", 0)
            .Italic = GetSetting(App.Title, "Options", "FontItalic", 0)
            .Underline = GetSetting(App.Title, "Options", "FontUnderline", 0)
            .Strikethrough = GetSetting(App.Title, "Options", "FontStrikethru", 0)
            .Size = GetSetting(App.Title, "Options", "FontSize", 6)
        End With
        .BackColor = CLng(GetSetting(App.Title, "Options", "BackColor", RGB(255, 255, 231)))
        .ForeColor = CLng(GetSetting(App.Title, "Options", "ForeColor", RGB(132, 132, 132)))
    End With
    RulerScaleMode = GetSetting(App.Title, "Options", "ScaleMode", 0)
    iTransparencyRuler = GetSetting(App.Title, "Options", "TransparencyRuler", 0)
    If RulerScaleMode > 2 Or RulerScaleMode < 0 Then RulerScaleMode = PL_PIXEL
    If iTransparencyRuler > 0 Then Call frmMenu.mnuTransparencyRuler_Click(iTransparencyRuler)
    
    On Error GoTo Form_Load_Error
    lCurrentStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    Call SetWindowLong(Me.hwnd, GWL_STYLE, lCurrentStyle And Not WS_BORDER)
    mRulerOrientation = PL_HORIZONTAL
    SetWindowPos hwnd, HWND_TOPMOST, 0, LScreenHeight \ 2, LScreenWidth, plBREADTH, 0&
    Call frmMenu.mnuScaleMode_Click(CInt(RulerScaleMode))
    Call GetAsyncKeyState(VK_LBUTTON) 'initialisieren
    startCmd = Command
    If (startCmd = "-s" And Not CBool(GetAsyncKeyState(vbKeyShift))) Or (startCmd = "" And CBool(GetAsyncKeyState(vbKeyShift))) Then
        Me.Visible = False
        Set Capture = New frmCapture
        Capture.Show vbModeless, Me
    End If
Exit Sub

Form_Load_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmRuler.Form_Load." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub Form_Resize()
    If mRulerOrientation = PL_HORIZONTAL Then
        mRulerWidth = ScaleWidth
        If mRulerWidth < 100 Then mRulerWidth = 100
        picRuler.Width = mRulerWidth
        If mRulerWidth < LScreenHeight Then mRulerHeight = mRulerWidth
    Else
        mRulerHeight = ScaleHeight
        If mRulerHeight < 100 Then mRulerHeight = 100
        picRuler.Height = mRulerHeight
        If mRulerHeight > mRulerWidth Then mRulerWidth = mRulerHeight
    End If
    mRedrawRequired = True
End Sub

Private Sub Form_Unload(cancel As Integer)
Dim f As Form
    On Error Resume Next
    TimerMagGlass.Enabled = False
    If Not MagGlass Is Nothing Then Unload MagGlass
    If Not Capture Is Nothing Then Unload Capture
    Set MagGlass = Nothing
    Set Capture = Nothing
    For Each f In Forms
        If f Is frmImage Then
            Unload f
        ElseIf f Is frmMagGlass Then
            Unload f
            Set MagGlass = Nothing
        End If
    Next
    Set f = Nothing
    gdiplus.TerminateGDI
    Unload frmMenu
    Set frmMenu = Nothing
    Set frmRuler = Nothing
End Sub


Private Sub SetScaleUser(X As Single, Y As Single)
Dim benutzerwert As Double
Dim benutzerwertStr As String
Dim eingabeok As Boolean
Dim prompt As String
    prompt = "Bitte Referenzwert eingeben:" & vbCrLf & "(1-10000)"
    While Not eingabeok
        benutzerwertStr = Trim$(InputBox(prompt, "Referenzwert", benutzerwertStr))
        If Len(benutzerwertStr) = 0 Then Exit Sub
        On Error Resume Next
        benutzerwert = CDbl(benutzerwertStr)
        If Err Or benutzerwert < 1 Or benutzerwert > 10000 Then
            Err.Clear
            prompt = "Ihre Eingabe kann nicht als gültige positive Zahl interpretiert werden. Bitte wiederholen Sie Ihre Eingabe:" & vbCrLf & "(1-10000)"
        Else
            eingabeok = True
            If mRulerOrientation = PL_HORIZONTAL Then
                RulerScaleMulti = benutzerwert / (X + 1)
            Else
                RulerScaleMulti = benutzerwert / (Y + 1)
            End If
            XYFieldWidth = XYFieldMinWidth + Len(CStr(Fix(RulerScaleMulti * 1000))) * 8
            ProcRefreshRuler frmRuler.Left, frmRuler.Top
            Exit Sub
        End If
    Wend
End Sub

Private Sub TimerMagGlass_Timer()
Dim tCursorPos As POINTAPI
Static tCursorPos0 As POINTAPI
    If modMain.CloseApp Then
        TimerMagGlass.Enabled = False
        Unload Me
    End If

    GetCursorPos tCursorPos
    'Strg+C untersuchen
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(VK_MBUTTON) Then
        frmRuler.Move tCursorPos.X * LTwipsPerPixelX, tCursorPos.Y * LTwipsPerPixelY
        Call GetAsyncKeyState(VK_LBUTTON)
    ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyC) Then
        Dim lPxColor As Long
        lPxColor = GetPxColor
        CopyRGB lPxColor
        Exit Sub
    End If
    If Not MagGlass Is Nothing Then
        If (tCursorPos0.X = tCursorPos.X And tCursorPos0.Y = tCursorPos.Y) Then 'keine Änderung der Mauszeigerposition
            ForceRefresh = ForceRefresh - 1
        Else
            ForceRefresh = FORCE_REFRESH_RES
        End If
        If ForceRefresh < 1 Then
          ForceRefresh = -1
          Sleep 50
          Exit Sub
        End If
        MagGlass.PrintMagGlass tCursorPos
    End If
    If Not MagColor Is Nothing Then MagColor.PrintMagColor tCursorPos
    
    Call CopyMemory(tCursorPos0, tCursorPos, LenB(tCursorPos))
    If mRedrawRequired Then
        mRedrawRequired = False
        ProcRefreshRuler Left, Top
    End If
End Sub

Private Sub picRuler_DblClick()
  On Error GoTo picRuler_DblClick_Error
  Orientation = Abs(Orientation - 1)
  Exit Sub
  
picRuler_DblClick_Error:
  Screen.MousePointer = vbDefault
  MsgBox "Fehler: " & Err.Number & vbCrLf & _
   "Beschreibung: " & Err.Description & vbCrLf & _
   "Quelle: frmRuler.picRuler_DblClick." & Erl & vbCrLf & Err.Source, _
   vbCritical
End Sub

Private Sub picRuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim isKeyCtrl As Boolean
Dim ubMarker As Long
Dim MousePosition As MousePos
    If Button = vbLeftButton Then
        isKeyCtrl = CBool(GetAsyncKeyState(vbKeyControl))
        If isKeyCtrl <> 0 And RulerScaleMode = PL_USER Then
          Call SetScaleUser(X, Y)
        Else
          MousePosition = GetMousePos(Me, X, Y, 5)
          ReleaseCapture
          If mRulerOrientation = PL_HORIZONTAL And (MousePosition = mpRight) Then
            PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_Right, 0&
            mRedrawRequired = True
          ElseIf mRulerOrientation = PL_VERTICAL And (MousePosition = mpBottom) Then
            PostMessage hwnd, WM_SYSCOMMAND, SC_SIZE_Bottom, 0&
            mRedrawRequired = True
          Else
            SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
          End If
        End If
    ElseIf Button = vbRightButton Then
      If ColorCollection(0) = -1 Then
        frmMenu.mnuColorCollection.Enabled = False
      Else
        frmMenu.mnuColorCollection.Enabled = True
      End If
      'Marker Menü einstellen auf setzen oder entfernen
      If mRulerOrientation = PL_HORIZONTAL Then
          ubMarker = UBound(HMarker)
          For i = 1 To ubMarker
              If X = HMarker(i) Or X = HMarker(i) - 1 Or X = HMarker(i) + 1 Then
                  frmMenu.mnuMarker.Caption = "Markierer entfernen"
                  frmMenu.mnuMarker.Tag = i
                  i = 0
                  Exit For
              End If
          Next i
          If i > 0 Then
              frmMenu.mnuMarker.Caption = "Markierer setzen          M"
              frmMenu.mnuMarker.Tag = "+"
          End If
          TMarker = X 'X-Pos zwischenspeichern, damit er in frmMenu abrufbar wird
      Else
          ubMarker = UBound(VMarker)
          For i = 1 To ubMarker
              If Y = VMarker(i) Or Y = VMarker(i) - 1 Or Y = VMarker(i) + 1 Then
                  frmMenu.mnuMarker.Caption = "Markierer entfernen"
                  frmMenu.mnuMarker.Tag = i
                  i = 0
                  Exit For
              End If
          Next i
          If i > 0 Then
              frmMenu.mnuMarker.Caption = "Markierer setzen          M"
              frmMenu.mnuMarker.Tag = "+"
          End If
          TMarker = Y 'Y-Pos zwischenspeichern, damit er in frmMenu abrufbar wird
      End If
      PopupMenu frmMenu.MRuler
        
    ElseIf Button = vbMiddleButton Then
      Dim tCursorPos As POINTAPI
      GetCursorPos tCursorPos
      Me.Move tCursorPos.X * LTwipsPerPixelX, tCursorPos.Y * LTwipsPerPixelY
    End If
End Sub

Private Sub picRuler_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MousePosition As MousePos

    MousePosition = GetMousePos(Me, X, Y, 5)
    If mRulerOrientation = PL_HORIZONTAL And (MousePosition = mpRight) Then
      picRuler.MousePointer = vbSizeWE
    ElseIf mRulerOrientation = PL_VERTICAL And (MousePosition = mpBottom) Then
      picRuler.MousePointer = vbSizeNS
    Else
        picRuler.MousePointer = vbCustom
        If RulerScaleMode = PL_USER Then
            If mRulerOrientation = PL_HORIZONTAL Then
                picRuler.ToolTipText = Round(((X + plZeroLine) * Abs(RulerScaleMulti)) * 1000) / 1000
                TMarker = X 'X-Pos zwischenspeichern, damit er mit der M-Taste abrufbar wird
            Else
                picRuler.ToolTipText = Round(((Y + plZeroLine) * Abs(RulerScaleMulti)) * 1000) / 1000
                TMarker = Y 'Y-Pos zwischenspeichern, damit er mit der M-Taste abrufbar wird
            End If
        Else
            If mRulerOrientation = PL_HORIZONTAL Then
                picRuler.ToolTipText = (X + plZeroLine) * Abs(RulerScaleMulti)
                TMarker = X 'X-Pos zwischenspeichern, damit er mit der M-Taste abrufbar wird
            Else
                picRuler.ToolTipText = (Y + plZeroLine) * Abs(RulerScaleMulti)
                TMarker = Y 'Y-Pos zwischenspeichern, damit er mit der M-Taste abrufbar wird
            End If
        End If
    End If

End Sub


