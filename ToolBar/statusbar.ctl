VERSION 5.00
Begin VB.UserControl StatusBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16215
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1081
   ToolboxBitmap   =   "statusbar.ctx":0000
   Begin VB.TextBox txtLegend 
      Alignment       =   2  'Zentriert
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10725
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      ToolTipText     =   "Nächstes Legendenzeichen"
      Top             =   165
      Width           =   375
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   180
      Index           =   27
      Left            =   9405
      Picture         =   "statusbar.ctx":0312
      ToolTipText     =   "Voll"
      Top             =   375
      Width           =   180
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   180
      Index           =   26
      Left            =   9405
      Picture         =   "statusbar.ctx":053C
      ToolTipText     =   "Voll"
      Top             =   165
      Width           =   180
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   2160
      Picture         =   "statusbar.ctx":0766
      ToolTipText     =   "Dünn"
      Top             =   90
      Width           =   240
   End
   Begin VB.Shape shFill 
      BorderColor     =   &H00404040&
      Height          =   270
      Left            =   1485
      Top             =   165
      Width           =   270
   End
   Begin VB.Shape shLine 
      BorderColor     =   &H00404040&
      Height          =   270
      Left            =   825
      Top             =   165
      Width           =   270
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   6810
      TabIndex        =   3
      ToolTipText     =   "Vordergrund-, Linie- und Rahmen-Farbe"
      Top             =   210
      Width           =   210
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H00C0C000&
      Height          =   270
      Left            =   1815
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   3
      Left            =   3075
      Picture         =   "statusbar.ctx":0AF0
      ToolTipText     =   "Dick"
      Top             =   90
      Width           =   240
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   2
      Left            =   2775
      Picture         =   "statusbar.ctx":0E7A
      ToolTipText     =   "Medium"
      Top             =   90
      Width           =   240
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   2475
      Picture         =   "statusbar.ctx":1204
      ToolTipText     =   "Dünn"
      Top             =   90
      Width           =   240
   End
   Begin VB.Image sbSeparator 
      Height          =   240
      Index           =   0
      Left            =   3300
      Picture         =   "statusbar.ctx":158E
      Top             =   90
      Width           =   240
   End
   Begin VB.Image sbSeparator 
      Height          =   240
      Index           =   1
      Left            =   6435
      Picture         =   "statusbar.ctx":1918
      Top             =   165
      Width           =   240
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   4
      Left            =   5505
      Picture         =   "statusbar.ctx":1CA2
      ToolTipText     =   "Transparent"
      Top             =   165
      Width           =   240
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   5
      Left            =   5865
      Picture         =   "statusbar.ctx":202C
      ToolTipText     =   "Durchsichtig"
      Top             =   165
      Width           =   240
   End
   Begin VB.Image sbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   6
      Left            =   6225
      Picture         =   "statusbar.ctx":23B6
      ToolTipText     =   "Voll"
      Top             =   165
      Width           =   240
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   10
      Left            =   7410
      TabIndex        =   20
      ToolTipText     =   "Schwarz"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   11
      Left            =   7650
      TabIndex        =   19
      ToolTipText     =   "Kastanienbraun"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   12
      Left            =   7890
      TabIndex        =   18
      ToolTipText     =   "Grün"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00008080&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   13
      Left            =   8130
      TabIndex        =   17
      ToolTipText     =   "Olivgrün"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   14
      Left            =   8370
      TabIndex        =   16
      ToolTipText     =   "Marineblau"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   15
      Left            =   8610
      TabIndex        =   15
      ToolTipText     =   "Lila"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   16
      Left            =   8850
      TabIndex        =   14
      ToolTipText     =   "Blaugrün"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   17
      Left            =   9090
      TabIndex        =   13
      ToolTipText     =   "Grau"
      Top             =   165
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   18
      Left            =   7410
      TabIndex        =   12
      ToolTipText     =   "Silber"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   19
      Left            =   7650
      TabIndex        =   11
      ToolTipText     =   "Rot"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   20
      Left            =   7890
      TabIndex        =   10
      ToolTipText     =   "Gelbgrün"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   21
      Left            =   8130
      TabIndex        =   9
      ToolTipText     =   "Gelb"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   22
      Left            =   8370
      TabIndex        =   8
      ToolTipText     =   "Blau"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   23
      Left            =   8610
      TabIndex        =   7
      ToolTipText     =   "Violett"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   24
      Left            =   8850
      TabIndex        =   6
      ToolTipText     =   "Aquamarin"
      Top             =   375
      Width           =   180
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   25
      Left            =   9090
      TabIndex        =   5
      ToolTipText     =   "Weiß"
      Top             =   375
      Width           =   180
   End
   Begin VB.Image sbSeparator 
      Height          =   240
      Index           =   2
      Left            =   9570
      Picture         =   "statusbar.ctx":2740
      Top             =   165
      Width           =   240
   End
   Begin VB.Label sbColor 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   6930
      TabIndex        =   4
      ToolTipText     =   "Hintergrund- und Füllfarbe"
      Top             =   330
      Width           =   210
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Legende"
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
      Left            =   9765
      TabIndex        =   2
      ToolTipText     =   "Nächstes Legendenzeichen"
      Top             =   195
      Width           =   825
   End
   Begin VB.Label lblKoordinaten 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,0"
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
      Left            =   11700
      TabIndex        =   1
      Top             =   195
      Width           =   300
   End
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Public Enum sbButtons
    sbLine0 = 0
    sbLine1 = 1
    sbLine2 = 2
    sbLine3 = 3
    sbFill0 = 4
    sbFill1 = 5
    sbFill2 = 6
    sbForeColor = 7
    sbBackColor = 8
    sbReset = 26
    sbPicker = 27
End Enum

Private mHover As Integer
Private mLine As Integer
Private mFill As Integer
Private mLegend As Boolean
Private mPalette As Integer

Public Event Click(Button As sbButtons)

Public Property Get BackColor() As Long
    BackColor = sbColor(1).BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As Long)
    If vNewValue > 16777215 Then vNewValue = 16777215
    sbColor(1).BackColor = vNewValue
End Property

Public Property Let Coordinates(ByVal vNewValue As String)
    lblKoordinaten.Caption = vNewValue
End Property

Public Property Get Fill() As Integer
    Fill = mFill
End Property

Public Property Let Fill(ByVal vNewValue As Integer)
    Select Case vNewValue
        Case 2
            mFill = vNewValue
            shFill.Move sbTool(sbFill2).Left - 1, sbTool(sbFill2).Top - 1
        Case 1
            mFill = vNewValue
            shFill.Move sbTool(sbFill1).Left - 1, sbTool(sbFill1).Top - 1
        Case Else
            mFill = 0
            shFill.Move sbTool(sbFill0).Left - 1, sbTool(sbFill0).Top - 1
        End Select
End Property

Public Property Get Palette() As Integer
    Palette = mPalette
End Property

Public Property Let Palette(ByVal vNewValue As Integer)
    mPalette = vNewValue
End Property

Public Property Get ForeColor() As Long
    ForeColor = sbColor(0).BackColor
End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
    If vNewValue > 16777215 Then vNewValue = 16777215
    sbColor(0).BackColor = vNewValue
End Property

Public Property Get Legend() As Boolean
    Legend = mLegend
End Property

Public Sub LegendIncrease()
Dim sLegend As String
    sLegend = txtLegend.Text
    If Asc(sLegend) > 160 Then Exit Sub
    Select Case sLegend
        Case "9":   txtLegend.Text = "A"
        Case "Z":   txtLegend.Text = "a"
        Case "a":   txtLegend.Text = "0"
        Case "^", "°", "!", """", "§", "$", "%", "&", "/", "(", ")", "=", "?", "+", "*", "#", "-", "_", ".", ":", ",", ";", " ", "{", "}", "[", "]", "\", "ß"
        Case Else: txtLegend.Text = Chr$(Asc(sLegend) + 1)
    End Select
End Sub

Public Property Let Legend(ByVal vNewValue As Boolean)
    mLegend = vNewValue
    txtLegend.Visible = mLegend
    lblLegend.Visible = mLegend
    If mLegend Then
        lblKoordinaten.Move txtLegend.Left + txtLegend.Width + 8, 7
    Else
        lblKoordinaten.Move lblLegend.Left
    End If
End Property

Public Property Get LegendText() As String
    LegendText = txtLegend.Text
End Property

Public Property Let LegendText(ByVal vNewValue As String)
    txtLegend.Text = Left$(vNewValue, 1)
End Property

Public Property Get Line() As Integer
    Line = mLine
End Property

Public Property Let Line(ByVal vNewValue As Integer)
    Select Case vNewValue
        Case 3
            mLine = vNewValue
            shLine.Move sbTool(sbLine3).Left - 1, sbTool(sbLine3).Top - 1
        Case 2
            mLine = vNewValue
            shLine.Move sbTool(sbLine2).Left - 1, sbTool(sbLine2).Top - 1
        Case 1
            mLine = vNewValue
            shLine.Move sbTool(sbLine1).Left - 1, sbTool(sbLine1).Top - 1
        Case Else
            mLine = 0
            shLine.Move sbTool(sbLine0).Left - 1, sbTool(sbLine0).Top - 1
        End Select
End Property


Private Sub UserControl_DblClick()
Dim col As Long
    If mHover >= 10 And mHover <= 25 Then
        shBorder.Visible = False
        col = sbColor(mHover).BackColor
        col = ShowColorDlg(UserControl.Parent.hwnd, col)
        If col > -1 Then
            sbColor(mHover).BackColor = col
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    mHover = -1
    UserControl.Height = 450
    shBorder.Move 0, 5
    sbTool(sbLine0).Move 4, 6
    sbTool(sbLine1).Move sbTool(sbLine0).Left + 20, 6
    sbTool(sbLine2).Move sbTool(sbLine1).Left + 20, 6
    sbTool(sbLine3).Move sbTool(sbLine2).Left + 20, 6
        sbSeparator(0).Move sbTool(sbLine3).Left + 16, 6
    sbTool(sbFill0).Move sbSeparator(0).Left + 16, 6
    sbTool(sbFill1).Move sbTool(sbFill0).Left + 20, 6
    sbTool(sbFill2).Move sbTool(sbFill1).Left + 20, 6
        sbSeparator(1).Move sbTool(sbFill2).Left + 16, 6
    sbColor(0).Move sbSeparator(1).Left + 14, 6
    sbColor(1).Move sbSeparator(1).Left + 20, 13
    sbColor(10).Move sbSeparator(1).Left + 40, 3
    sbColor(11).Move sbColor(10).Left + 16, 3
    sbColor(12).Move sbColor(11).Left + 16, 3
    sbColor(13).Move sbColor(12).Left + 16, 3
    sbColor(14).Move sbColor(13).Left + 16, 3
    sbColor(15).Move sbColor(14).Left + 16, 3
    sbColor(16).Move sbColor(15).Left + 16, 3
    sbColor(17).Move sbColor(16).Left + 16, 3
    sbTool(sbReset).Move sbColor(17).Left + 18, 3
    
    sbColor(18).Move sbSeparator(1).Left + 40, 16
    sbColor(19).Move sbColor(18).Left + 16, 16
    sbColor(20).Move sbColor(19).Left + 16, 16
    sbColor(21).Move sbColor(20).Left + 16, 16
    sbColor(22).Move sbColor(21).Left + 16, 16
    sbColor(23).Move sbColor(22).Left + 16, 16
    sbColor(24).Move sbColor(23).Left + 16, 16
    sbColor(25).Move sbColor(24).Left + 16, 16
    sbTool(sbPicker).Move sbColor(25).Left + 18, 16
        sbSeparator(2).Move sbTool(sbReset).Left + 16, 8
    lblLegend.Move sbSeparator(2).Left + 14, 7
    txtLegend.Move lblLegend.Left + lblLegend.Width + 4, 6
    lblKoordinaten.Move txtLegend.Left + txtLegend.Width + 8, 7
    shLine.Move sbTool(sbLine0).Left - 1, sbTool(sbLine0).Top - 1
    shFill.Move sbTool(sbFill0).Left - 1, sbTool(sbFill0).Top - 1
    shBorder.ZOrder
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    mHover = -1
    For i = 0 To sbFill2
        With sbTool(i)
            If X > .Left And X < .Left + 16 And Y > 6 And Y < 23 Then
                shBorder.Move .Left - 1, .Top - 1, .Width + 2, .Height + 2
                mHover = i
                GoTo Finaly
            End If
        End With
    Next i
    For i = 10 To 17
        With sbColor(i)
            If X > .Left And X < .Left + 16 And Y > 3 And Y < 15 Then
                shBorder.Move .Left - 1, .Top - 1, .Width + 2, .Height + 2
                mHover = i
                GoTo Finaly
            End If
        End With
    Next i
    For i = 18 To 25
        With sbColor(i)
            If X > .Left And X < .Left + 16 And Y > 15 And Y < 31 Then
                shBorder.Move .Left - 1, .Top - 1, .Width + 2, .Height + 2
                mHover = i
                GoTo Finaly
            End If
        End With
    Next i
    With sbTool(sbReset)
        If X > .Left And X < .Left + 16 And Y > 3 And Y < 15 Then
            shBorder.Move .Left - 1, .Top - 1, .Width + 2, .Height + 2
            mHover = sbReset
            GoTo Finaly
        End If
    End With
    With sbTool(sbPicker)
        If X > .Left And X < .Left + 16 And Y > 15 And Y < 31 Then
            shBorder.Move .Left - 1, .Top - 1, .Width + 2, .Height + 2
            mHover = sbPicker
            GoTo Finaly
        End If
    End With

Finaly:
    If mHover < 0 Then
        shBorder.Visible = False
        ReleaseCapture
    Else
        shBorder.Visible = True
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As sbButtons
Dim j As Integer
Dim palColors() As Long
    j = -1
    For i = 0 To sbFill2    'SCHALTER
        If X >= sbTool(i).Left And X <= sbTool(i).Left + sbTool(i).Width And Y >= 6 And Y <= 22 Then
            j = i
            Exit For
        End If
    Next
    Select Case j
        Case sbLine0:  Me.Line = 0      '0
        Case sbLine1:  Me.Line = 1      '1
        Case sbLine2:  Me.Line = 2      '2
        Case sbLine3:  Me.Line = 3      '3
        Case sbFill0:  Me.Fill = 0      '4
        Case sbFill1:  Me.Fill = 1      '5
        Case sbFill2:  Me.Fill = 2      '6
    End Select
    If j >= 0 Then
        RaiseEvent Click(i)
        Exit Sub
    End If
    For i = 10 To 17    'FARBEN OBEN
        If X > sbColor(i).Left And X < sbColor(i).Left + 16 And Y > 3 And Y < 15 Then
            If Button = vbLeftButton Then
                sbColor(0).BackColor = sbColor(i).BackColor
                RaiseEvent Click(sbForeColor)
            ElseIf Button = vbRightButton Then
                sbColor(1).BackColor = sbColor(i).BackColor
                RaiseEvent Click(sbBackColor)
            End If
            Exit Sub
        End If
    Next i
    For i = 18 To 25    'FARBEN UNTEN
        If X > sbColor(i).Left And X < sbColor(i).Left + 16 And Y > 15 And Y < 31 Then
            If Button = vbLeftButton Then
                sbColor(0).BackColor = sbColor(i).BackColor
                RaiseEvent Click(sbForeColor)
            ElseIf Button = vbRightButton Then
                sbColor(1).BackColor = sbColor(i).BackColor
                RaiseEvent Click(sbBackColor)
            End If
            Exit Sub
        End If
    Next i
    With sbTool(sbReset)
        j = 0
        If X > .Left And X < .Left + 16 And Y > 3 And Y < 15 Then
            X = sbTool(sbReset).Left * LTwipsPerPixelX
            If Parent.Top + Parent.Height + (160 * LTwipsPerPixelY) > Screen.Height Then Y = Parent.ScaleHeight - UserControl.Height Else Y = Parent.ScaleHeight - (UserControl.Height / 2)
            If frmMenu.GetPopupMenu(UserControl.Parent, X, Y, "Palette", "", j, False) Then
                palColors = frmMenu.GetPalColors(j)
                For j = 0 To 15
                    sbColor(j + 10).BackColor = palColors(j)
                Next j
            End If
            Exit Sub
        End If
    End With
    With sbTool(sbPicker)
        If X > .Left And X < .Left + 16 And Y > 15 And Y < 31 Then
            shBorder.Visible = False
            RaiseEvent Click(sbPicker)
        End If
    End With
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If UserControl.Ambient.UserMode Then Legend = mLegend
End Sub

Private Sub UserControl_Resize()
Dim lScaleWidth As Single, lScaleHeight As Single

    lScaleWidth = UserControl.ScaleWidth
    lScaleHeight = UserControl.ScaleHeight
    UserControl.Cls
    UserControl.Line (0, 0)-(lScaleWidth, 0), vbButtonShadow       'H
    'Gripper
    UserControl.Line (lScaleWidth - 11, lScaleHeight - 3)-(lScaleWidth - 10, lScaleHeight - 2), vbWhite, BF   'H
    UserControl.Line (lScaleWidth - 12, lScaleHeight - 4)-(lScaleWidth - 11, lScaleHeight - 3), &HA5B5BD, BF   'D
    UserControl.Line (lScaleWidth - 7, lScaleHeight - 7)-(lScaleWidth - 6, lScaleHeight - 6), vbWhite, BF   'H
    UserControl.Line (lScaleWidth - 8, lScaleHeight - 8)-(lScaleWidth - 7, lScaleHeight - 7), &HA5B5BD, BF   'D
    UserControl.Line (lScaleWidth - 3, lScaleHeight - 11)-(lScaleWidth - 2, lScaleHeight - 10), vbWhite, BF   'H
    UserControl.Line (lScaleWidth - 4, lScaleHeight - 12)-(lScaleWidth - 3, lScaleHeight - 11), &HA5B5BD, BF   'D
    
    UserControl.Line (lScaleWidth - 7, lScaleHeight - 3)-(lScaleWidth - 6, lScaleHeight - 2), vbWhite, BF   'H
    UserControl.Line (lScaleWidth - 8, lScaleHeight - 4)-(lScaleWidth - 7, lScaleHeight - 3), &HA5B5BD, BF   'D
    UserControl.Line (lScaleWidth - 3, lScaleHeight - 7)-(lScaleWidth - 2, lScaleHeight - 6), vbWhite, BF   'H
    UserControl.Line (lScaleWidth - 4, lScaleHeight - 8)-(lScaleWidth - 3, lScaleHeight - 7), &HA5B5BD, BF   'D
    
    UserControl.Line (lScaleWidth - 3, lScaleHeight - 3)-(lScaleWidth - 2, lScaleHeight - 2), vbWhite, BF   'H
    UserControl.Line (lScaleWidth - 4, lScaleHeight - 4)-(lScaleWidth - 3, lScaleHeight - 3), &HA5B5BD, BF   'D
End Sub

Private Sub sbColor_Click(Index As Integer)
Dim c As Long
    
    If Index <= 1 Then
        c = sbColor(0).BackColor
        sbColor(0).BackColor = sbColor(1).BackColor
        RaiseEvent Click(sbForeColor)
        sbColor(1).BackColor = c
        RaiseEvent Click(sbBackColor)
    End If
End Sub

Private Sub sbColor_DblClick(Index As Integer)
    If Index <= 1 Then
        sbColor_Click Index
    End If
End Sub
Private Sub sbColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > 1 Then SetCapture UserControl.hwnd
End Sub

Private Sub sbTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture UserControl.hwnd
End Sub

Private Sub txtLegend_GotFocus()
    txtLegend.SelStart = 0
    txtLegend.SelLength = 1
End Sub

