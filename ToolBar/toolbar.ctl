VERSION 5.00
Begin VB.UserControl ToolBar 
   Alignable       =   -1  'True
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13125
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   875
   ToolboxBitmap   =   "toolbar.ctx":0000
   Begin VB.Timer tRefresh 
      Interval        =   1000
      Left            =   10890
      Top             =   0
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   6
      Left            =   5655
      Picture         =   "toolbar.ctx":0402
      ToolTipText     =   "Verwischen"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgRuler 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   660
      Picture         =   "toolbar.ctx":078C
      ToolTipText     =   "Pixel-Lineal"
      Top             =   660
      Width           =   240
   End
   Begin VB.Image imgRuler 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   660
      Picture         =   "toolbar.ctx":0B16
      ToolTipText     =   "Pixel-Lineal"
      Top             =   450
      Width           =   240
   End
   Begin VB.Image imgSize 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   2310
      Picture         =   "toolbar.ctx":0EA0
      ToolTipText     =   "Bild zuschneiden"
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSize 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   2310
      Picture         =   "toolbar.ctx":122A
      Top             =   660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   17
      Left            =   1875
      Picture         =   "toolbar.ctx":15B4
      ToolTipText     =   "Pixel-Lineal"
      Top             =   0
      Width           =   240
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H00C0C000&
      Height          =   270
      Left            =   2970
      Top             =   495
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   13
      Left            =   165
      Picture         =   "toolbar.ctx":193E
      ToolTipText     =   "Menü"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   10
      Left            =   7005
      Picture         =   "toolbar.ctx":1CC8
      ToolTipText     =   "Text"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgUndo 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   1710
      Picture         =   "toolbar.ctx":2052
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUndo 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   1710
      Picture         =   "toolbar.ctx":23DC
      Top             =   690
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRedo 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   1950
      Picture         =   "toolbar.ctx":2766
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRedo 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   1950
      Picture         =   "toolbar.ctx":2AF0
      Top             =   690
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNew 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   990
      Picture         =   "toolbar.ctx":2E7A
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNew 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   990
      Picture         =   "toolbar.ctx":3204
      Top             =   690
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCopy 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   1230
      Picture         =   "toolbar.ctx":358E
      ToolTipText     =   "Bild in die Zwischenablage kopieren"
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCopy 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   1230
      Picture         =   "toolbar.ctx":3918
      ToolTipText     =   "Bild in die Zwischenablage kopieren"
      Top             =   690
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image tbSeparator 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   2
      Left            =   7230
      Picture         =   "toolbar.ctx":3CA2
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTextstyle 
      Appearance      =   0  '2D
      Height          =   105
      Left            =   9690
      Picture         =   "toolbar.ctx":402C
      ToolTipText     =   "Schrifteinstellungen"
      Top             =   90
      Width           =   105
   End
   Begin VB.Label lblTextstyle 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   " 8 - Verdana"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7440
      TabIndex        =   0
      ToolTipText     =   "Schrifteinstellungen"
      Top             =   0
      Width           =   2100
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   3795
      Picture         =   "toolbar.ctx":4156
      ToolTipText     =   "Zeiger"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   4095
      Picture         =   "toolbar.ctx":44E0
      ToolTipText     =   "Freihand"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   2
      Left            =   4395
      Picture         =   "toolbar.ctx":486A
      ToolTipText     =   "Linie"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   3
      Left            =   4695
      Picture         =   "toolbar.ctx":4BF4
      ToolTipText     =   "Rechteck"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   5
      Left            =   5355
      Picture         =   "toolbar.ctx":4F7E
      ToolTipText     =   "Markierer"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   8
      Left            =   6375
      Picture         =   "toolbar.ctx":5308
      ToolTipText     =   "Pfeil"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   7
      Left            =   6015
      Picture         =   "toolbar.ctx":5692
      ToolTipText     =   "Füllen"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   4
      Left            =   4995
      Picture         =   "toolbar.ctx":5A1C
      ToolTipText     =   "Kreis"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   9
      Left            =   6675
      Picture         =   "toolbar.ctx":5DA6
      ToolTipText     =   "Legende"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   20
      Left            =   3135
      Picture         =   "toolbar.ctx":6130
      ToolTipText     =   "Rückgängig"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   21
      Left            =   3375
      Picture         =   "toolbar.ctx":64BA
      ToolTipText     =   "Wiederherstellen"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbSeparator 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   1
      Left            =   3615
      Picture         =   "toolbar.ctx":6844
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   15
      Left            =   855
      Picture         =   "toolbar.ctx":6BCE
      ToolTipText     =   "Bild in die Zwischenablage kopieren"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbSeparator 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   0
      Left            =   2985
      Picture         =   "toolbar.ctx":6F58
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   14
      Left            =   495
      Picture         =   "toolbar.ctx":72E2
      ToolTipText     =   "Neu aus Screenshot"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   18
      Left            =   2175
      Picture         =   "toolbar.ctx":766C
      ToolTipText     =   "Lupe"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   12
      Left            =   1215
      Picture         =   "toolbar.ctx":79F6
      ToolTipText     =   "Bild aus Zwischenablage einfügen"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   11
      Left            =   2745
      Picture         =   "toolbar.ctx":7D80
      ToolTipText     =   "Bild zuschneiden"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   16
      Left            =   1575
      Picture         =   "toolbar.ctx":810A
      ToolTipText     =   "Rahmen/Schatten/Abrisskante(n) einfügen"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image tbTool 
      Appearance      =   0  '2D
      Height          =   240
      Index           =   19
      Left            =   2445
      Picture         =   "toolbar.ctx":8494
      ToolTipText     =   "Verkleinern"
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "ToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSelected As Integer
Private mFontSize As Single
Private mArrowStyle As Integer

Public Enum tbButtons
    tbPointer = 0
    tbFreehand = 1
    tbLine = 2
    tbRectangle = 3
    tbCyrcle = 4
    tbMarker = 5
    tbObfus = 6
    tbFill = 7
    tbArrow = 8
    tbLegend = 9
    tbText = 10
    tbCrop = 11
    tbPaste = 12
    
    tbMenu = 13
    tbNew = 14
    tbCopy = 15
    tbBorderStyle = 16
    tbLineal = 17
    tbMagGlass = 18
    tbScale = 19
    tbUndo = 20
    tbRedo = 21
    tbFont = 22
    tbTear = 23
    tbExtend = 24
End Enum

Public Enum tbBorder
    tbbBorder = 0
    tbbTearTop = 2
    tbbTearRight = 3
    tbbTearBottom = 4
    tbbTearLeft = 5
    tbbTearMiddle = 6
    tbbShadow = 7
    tbbTearTopRight = 9
    tbbTearBottomRight = 10
    tbbTearBottomLeft = 11
    tbbTearTopLeft = 12
End Enum
    

Public Event Click(ByVal Value As tbButtons, ByVal X As Long)
Public Event Change(ByVal newValue As tbButtons, ByVal OldValue As tbButtons, ByRef cancel As Boolean)
Public Event MenuClick(Name As String, Caption As String, Index As Integer, Checked As Boolean)


Public Property Get Arrow() As Integer
    Arrow = mArrowStyle
End Property

Public Property Let Arrow(ByVal vNewValue As Integer)
Dim i As Integer
    If vNewValue = 3 Or vNewValue = 8 Or vNewValue = 13 Or vNewValue = 18 Or vNewValue > 19 Then vNewValue = 0
    mArrowStyle = vNewValue
    tbTool(tbArrow).Picture = frmMenu.picMenuArrow(mArrowStyle)
    For i = frmMenu.mnuArrow.LBound To frmMenu.mnuArrow.UBound
        frmMenu.mnuArrow(i).Checked = (i = mArrowStyle)
    Next i
End Property

Public Property Get Enabled(Button As tbButtons) As Boolean
    Enabled = tbTool(Button).Enabled
End Property

Public Property Let Enabled(Button As tbButtons, ByVal vNewValue As Boolean)
    tbTool(Button).Enabled = vNewValue
    Select Case True
        Case Button = tbUndo And vNewValue
            tbTool(tbUndo).Picture = imgUndo(0).Picture
        Case Button = tbUndo And Not vNewValue
            tbTool(tbUndo).Picture = imgUndo(1).Picture
        Case Button = tbRedo And vNewValue
            tbTool(tbRedo).Picture = imgRedo(0).Picture
        Case Button = tbRedo And Not vNewValue
            tbTool(tbRedo).Picture = imgRedo(1).Picture
    End Select
End Property
Public Property Get FontBackground() As Long
    FontBackground = lblTextstyle.BackColor
End Property
Public Property Let FontBackground(ByVal vNewValue As Long)
    lblTextstyle.BackColor = vNewValue
End Property
Public Property Get FontBold() As Boolean
    FontBold = lblTextstyle.FontBold
End Property
Public Property Let FontBold(ByVal vNewValue As Boolean)
    lblTextstyle.FontBold = vNewValue
End Property
Public Property Get FontColor() As Long
    FontColor = lblTextstyle.ForeColor
End Property
Public Property Let FontColor(ByVal vNewValue As Long)
    If vNewValue > 16777215 Then vNewValue = 16777215
    lblTextstyle.ForeColor = vNewValue
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = lblTextstyle.FontItalic
End Property
Public Property Let FontItalic(ByVal vNewValue As Boolean)
    lblTextstyle.FontItalic = vNewValue
End Property


'========================================

Public Property Get FontName() As String
    FontName = lblTextstyle.FontName
End Property
Public Property Let FontName(ByVal vNewValue As String)
    lblTextstyle.FontName = vNewValue
    lblTextstyle.Caption = mFontSize & " - " & vNewValue
End Property
Public Property Get FontSize() As Single
    FontSize = mFontSize
End Property
Public Property Let FontSize(ByVal vNewValue As Single)
    mFontSize = vNewValue
    lblTextstyle.Caption = mFontSize & " - " & lblTextstyle.FontName
End Property
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = lblTextstyle.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal vNewValue As Boolean)
    lblTextstyle.FontStrikethru = vNewValue
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = lblTextstyle.FontUnderline
End Property
Public Property Let FontUnderline(ByVal vNewValue As Boolean)
    lblTextstyle.FontUnderline = vNewValue
End Property


Public Property Get Selected() As tbButtons
    Selected = mSelected
End Property

Public Property Let Selected(ByVal vNewValue As tbButtons)
Dim i As Integer
    If vNewValue > tbMenu And vNewValue <> tbTear Then vNewValue = tbPointer
    mSelected = vNewValue
    For i = tbPointer To tbMenu
        tbTool(i).BorderStyle = Abs(mSelected = i)
    Next
End Property

Public Sub SetButtonShift(Value As Boolean)
    If Value Then
        tbTool(tbNew).Picture = imgNew(1).Picture
        tbTool(tbCopy).Picture = imgCopy(1).Picture
        tbTool(tbCrop).Picture = imgSize(1).Picture
        tbTool(tbLineal).Picture = imgRuler(1).Picture
    Else
        tbTool(tbNew).Picture = imgNew(0).Picture
        tbTool(tbCopy).Picture = imgCopy(0).Picture
        tbTool(tbCrop).Picture = imgSize(0).Picture
        tbTool(tbLineal).Picture = imgRuler(0).Picture
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.Height = 375
    shBorder.Move 0, 3
    tbTool(tbMenu).Move 4, 4
    tbTool(tbNew).Move tbTool(tbMenu).Left + 20, 4
    tbTool(tbCopy).Move tbTool(tbNew).Left + 20, 4
    tbTool(tbPaste).Move tbTool(tbCopy).Left + 20, 4
    tbTool(tbBorderStyle).Move tbTool(tbPaste).Left + 20, 4
    tbTool(tbLineal).Move tbTool(tbBorderStyle).Left + 20, 4
    tbTool(tbMagGlass).Move tbTool(tbLineal).Left + 20, 4
    tbTool(tbScale).Move tbTool(tbMagGlass).Left + 20, 4
    tbTool(tbCrop).Move tbTool(tbScale).Left + 20, 4
        tbSeparator(0).Move tbTool(tbCrop).Left + 16, 4
    tbTool(tbUndo).Move tbSeparator(0).Left + 16, 4: tbTool(tbUndo).Enabled = False
    tbTool(tbRedo).Move tbTool(tbUndo).Left + 20, 4: tbTool(tbRedo).Enabled = False
        tbSeparator(1).Move tbTool(tbRedo).Left + 16, 4
    tbTool(tbPointer).Move tbSeparator(1).Left + 20, 4
    tbTool(tbFreehand).Move tbTool(tbPointer).Left + 20, 4
    tbTool(tbLine).Move tbTool(tbFreehand).Left + 20, 4
    tbTool(tbRectangle).Move tbTool(tbLine).Left + 20, 4
    tbTool(tbCyrcle).Move tbTool(tbRectangle).Left + 20, 4
    tbTool(tbMarker).Move tbTool(tbCyrcle).Left + 20, 4
    tbTool(tbObfus).Move tbTool(tbMarker).Left + 20, 4
    tbTool(tbFill).Move tbTool(tbObfus).Left + 20, 4
    tbTool(tbArrow).Move tbTool(tbFill).Left + 20, 4
    tbTool(tbLegend).Move tbTool(tbArrow).Left + 20, 4
    tbTool(tbText).Move tbTool(tbLegend).Left + 20, 4
        tbSeparator(2).Move tbTool(tbText).Left + 16, 4
    lblTextstyle.Move tbSeparator(2).Left + 20, 4
    tbTextstyle.Move lblTextstyle.Left + lblTextstyle.Width - 12, 8
    tbTextstyle.ZOrder
    mFontSize = 8
    lblTextstyle.Caption = mFontSize & " - " & lblTextstyle.FontName
    shBorder.ZOrder
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Click(-1, CLng(X))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shBorder.Visible = False
    tRefresh.Enabled = False
End Sub

Private Sub UserControl_Paint()
Dim lScaleWidth As Single, lScaleHeight As Single
    lScaleWidth = UserControl.ScaleWidth
    lScaleHeight = UserControl.ScaleHeight
    UserControl.Line (0, lScaleHeight - 1)-(lScaleWidth, lScaleHeight - 1), vbButtonShadow 'H
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If UserControl.Ambient.UserMode Then tbTool(mSelected).BorderStyle = vbFixedSingle
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 375
End Sub


'Click
Private Sub lblTextstyle_Click()
    RaiseEvent Click(tbFont, CLng(lblTextstyle.Left))
End Sub

Private Sub lblTextstyle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblTextstyle
        shBorder.Move .Left - 1, .Top - 1, lblTextstyle.Width + 2, .Height + 2
        shBorder.Visible = True
        tRefresh.Enabled = True
    End With
End Sub


Private Sub tRefresh_Timer()
Dim tCursorPos As POINTAPI, tToolbarPos As POINTAPI
    If shBorder.Visible Then
        GetCursorPos tCursorPos
        ClientToScreen UserControl.hwnd, tToolbarPos
        If tCursorPos.X > tToolbarPos.X And tCursorPos.X < tToolbarPos.X + UserControl.ScaleWidth Then
            If tCursorPos.Y > tToolbarPos.Y And tCursorPos.Y < tToolbarPos.Y + UserControl.ScaleHeight Then
                If tCursorPos.X > tToolbarPos.X + shBorder.Left And tCursorPos.X < tToolbarPos.X + shBorder.Left + shBorder.Width Then
                    Exit Sub
                End If
            End If
        End If
    End If
    shBorder.Visible = False
    tRefresh.Enabled = False
    
End Sub
Private Sub tbTextstyle_Click()
    RaiseEvent Click(tbFont, CLng(lblTextstyle.Left))
End Sub

Private Sub tbTextstyle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblTextstyle
        shBorder.Move .Left - 1, .Top - 1, lblTextstyle.Width + 2, .Height + 2
        shBorder.Visible = True
        tRefresh.Enabled = True
    End With
End Sub

Private Sub tbTool_Click(Index As Integer)
Dim i As Integer
Dim cancel As Boolean
Dim pmName As String
Dim pmCaption As String
Dim pmIndex As Integer
Dim pmChecked As Boolean

    If Index < tbMenu Then         'Selektions-Schalter
        If GetAsyncKeyState(vbKeyShift) And Index = tbCrop Then
            SetButtonShift False
            RaiseEvent Click(tbExtend, tbTool(Index).Left)
            Exit Sub
        End If
        If Index = tbArrow Then
            pmName = "Arrow"
            If frmMenu.GetPopupMenu(UserControl.Parent, tbTool(Index).Left * LTwipsPerPixelX, UserControl.Height, pmName, pmCaption, pmIndex, pmChecked) Then Me.Arrow = pmIndex
        End If
        RaiseEvent Change(Index, mSelected, cancel)
        If cancel Then
            tbTool(Index).BorderStyle = vbBSNone
        Else
            For i = 0 To tbMenu - 1
                If i = Index Then tbTool(i).BorderStyle = vbFixedSingle Else tbTool(i).BorderStyle = vbBSNone
            Next
            mSelected = Index
        End If
    ElseIf Index = tbMenu Then
        RaiseEvent Click(Index, tbTool(Index).Left)
        pmName = "File"
        If frmMenu.GetPopupMenu(UserControl.Parent, tbTool(Index).Left * LTwipsPerPixelX, UserControl.Height, pmName, pmCaption, pmIndex, pmChecked) Then
            RaiseEvent MenuClick(pmName, pmCaption, pmIndex, pmChecked)
        End If
    ElseIf Index = tbBorderStyle Then
        RaiseEvent Click(Index, tbTool(Index).Left)
        pmName = "Border"
        If frmMenu.GetPopupMenu(UserControl.Parent, tbTool(Index).Left * LTwipsPerPixelX, UserControl.Height, pmName, pmCaption, pmIndex, pmChecked) Then
            RaiseEvent MenuClick(pmName, pmCaption, pmIndex, pmChecked)
        End If
    Else
        RaiseEvent Click(Index, tbTool(Index).Left)
    End If

End Sub
Private Sub tbTool_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index >= tbMenu Then
        tbTool(Index).BorderStyle = vbFixedSingle
    End If
End Sub

Private Sub tbTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With tbTool(Index)
        shBorder.Move .Left - 1, .Top - 1, .Width + 2, .Height + 2
        shBorder.Visible = True
        tRefresh.Enabled = True
    End With
End Sub

Private Sub tbTool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index >= tbMenu Then
        tbTool(Index).BorderStyle = vbBSNone
    End If
End Sub
