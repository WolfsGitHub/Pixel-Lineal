VERSION 5.00
Begin VB.Form frmImage 
   AutoRedraw      =   -1  'True
   Caption         =   "Pixel-Lineal"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13275
   Icon            =   "Image.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picImage 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   330
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   4
      Top             =   825
      Width           =   2535
      Begin VB.TextBox txtEditBox 
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Text            =   "txtEditBox"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape shpDimension 
         BorderColor     =   &H000000FF&
         Height          =   75
         Left            =   225
         Top             =   1350
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Line linDimension 
         BorderColor     =   &H000000FF&
         Index           =   2
         Visible         =   0   'False
         X1              =   28
         X2              =   28
         Y1              =   85
         Y2              =   99
      End
      Begin VB.Line linDimension 
         BorderColor     =   &H000000FF&
         Index           =   1
         Visible         =   0   'False
         X1              =   8
         X2              =   32
         Y1              =   96
         Y2              =   96
      End
      Begin VB.Line linDimension 
         BorderColor     =   &H000000FF&
         Index           =   0
         Visible         =   0   'False
         X1              =   11
         X2              =   11
         Y1              =   85
         Y2              =   99
      End
   End
   Begin PixelLineal.ToolBar TBar 
      Align           =   1  'Oben ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   661
   End
   Begin PixelLineal.StatusBar SBar 
      Align           =   2  'Unten ausrichten
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   794
   End
   Begin VB.PictureBox picPaste 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   240
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   106
      Left            =   7590
      Picture         =   "Image.frx":038A
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   63
      Left            =   6930
      Picture         =   "Image.frx":04DC
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   62
      Left            =   6930
      Picture         =   "Image.frx":062E
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   61
      Left            =   6930
      Picture         =   "Image.frx":0780
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   60
      Left            =   6930
      Picture         =   "Image.frx":08D2
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   53
      Left            =   6270
      Picture         =   "Image.frx":0A24
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   52
      Left            =   6270
      Picture         =   "Image.frx":0B76
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   51
      Left            =   6270
      Picture         =   "Image.frx":0CC8
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   50
      Left            =   6270
      Picture         =   "Image.frx":0E1A
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   40
      Left            =   5610
      Picture         =   "Image.frx":0F6C
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   4
      Left            =   2970
      Picture         =   "Image.frx":1C36
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   9
      Left            =   3630
      Picture         =   "Image.frx":1D88
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   14
      Left            =   4290
      Picture         =   "Image.frx":1EDA
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   19
      Left            =   4950
      Picture         =   "Image.frx":202C
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   43
      Left            =   5610
      Picture         =   "Image.frx":217E
      Top             =   3135
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   42
      Left            =   5610
      Picture         =   "Image.frx":22D0
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   41
      Left            =   5610
      Picture         =   "Image.frx":2422
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   71
      Left            =   7590
      Picture         =   "Image.frx":2574
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   17
      Left            =   4950
      Picture         =   "Image.frx":26C6
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   12
      Left            =   4290
      Picture         =   "Image.frx":2818
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   7
      Left            =   3630
      Picture         =   "Image.frx":296A
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   2
      Left            =   2970
      Picture         =   "Image.frx":2ABC
      Top             =   2475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   16
      Left            =   4950
      Picture         =   "Image.frx":2C0E
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   11
      Left            =   4290
      Picture         =   "Image.frx":2D60
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   6
      Left            =   3630
      Picture         =   "Image.frx":2EB2
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   1
      Left            =   2970
      Picture         =   "Image.frx":3004
      Top             =   1815
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   15
      Left            =   4950
      Picture         =   "Image.frx":3156
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   10
      Left            =   4290
      Picture         =   "Image.frx":32A8
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   5
      Left            =   3630
      Picture         =   "Image.frx":33FA
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   0
      Left            =   2970
      Picture         =   "Image.frx":354C
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image curPointer 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Index           =   70
      Left            =   7590
      Picture         =   "Image.frx":369E
      Top             =   1155
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUndoStack As clsUndoStack
Private mTextOverhang As Long
Private mCurrentFileName As String
Private mGradingVisible As Boolean

Private Type tWorkControl
    DrawMode As Integer
    x0 As Long
    y0 As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type
Private mWorkControl As tWorkControl

Private Type tDrawStyle
    DrawStyle As Integer
    DrawMode As Integer
    DrawWidth As Integer
    FillStyle As Integer
End Type
Private mDrawStyle As tDrawStyle

Private Enum eAction
    ActionStart
    ActionEnd
End Enum
    
Private Declare Function SendMessageL Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal destHdc As Long, _
    ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, _
    ByVal srcHdc As Long, ByVal srcx As Long, ByVal srcy As Long, _
    ByVal srcwidth As Long, ByVal scrHeight As Long, ByVal BLENDFUNCT As Long) As Long

'=====Neues Bild anlegen========================================
Public Sub ShowCapture(Left As Single, Top As Single, Width As Single, Height As Single, Img As StdPicture)
Dim offsetX As Single, offsetY As Single
Dim w As Long, h As Long
    Set mUndoStack = New clsUndoStack
    offsetX = Me.Width - ScaleWidth
    offsetY = Me.Height - ScaleHeight + TBar.Height
    w = ScaleX(Img.Width, vbHimetric, vbTwips)
    h = ScaleX(Img.Height, vbHimetric, vbTwips)
    picImage.Move 0, TBar.Height, w, h
    picImage.Picture = Img
    If w < TBar.Left + TBar.Width + offsetX + 300 Then w = TBar.Left + TBar.Width + offsetX + 300
    Me.Move Left, Top, w + offsetX, h + offsetY + SBar.Height + 450
    Me.Show
    PaintGrading
    mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
    TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
End Sub


Public Sub TextStyle(Optional reset As Boolean)
Dim TmpFName As String

  Dim LFnt As LOGFONT
  Dim CF_T As CHOOSEFONT_TYPE
  On Error GoTo mnuFonts_Click_Error
      If Not reset Then
          With CF_T
            .nSizeMax = 72
            .nSizeMin = 4
            .iPointSize = 100
            .Flags = CF_SCREENFONTS Or CF_FORCEFONTEXIST Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_NOSCRIPTSEL
            .hWndOwner = Me.hwnd
            .lStructSize = Len(CF_T)
            .lpLogFont = VarPtr(LFnt)
            .hInstance = App.hInstance
            .hDC = 0
            .nFontType = SCREEN_FONTTYPE
            .rgbColors = Convert_OLEtoRBG(TBar.FontColor)
          End With
      
          TmpFName = TBar.FontName
          TmpFName = StrConv(TmpFName, vbFromUnicode)
          LFnt.lfFaceName = TmpFName & vbNullChar
          With LFnt
              .lfHeight = TBar.FontSize * -20 / LTwipsPerPixelY 'Alternativ: 'MM_TEXT mapping mode: lfHeight = -MulDiv(PointSize, GetDeviceCaps(hDC, LOGPIXELSY), 72);
              .lfWeight = IIf(TBar.FontBold, FW_BOLD, FW_NORMAL)
              .lfItalic = Abs(TBar.FontItalic)
              .lfUnderline = Abs(TBar.FontUnderline)
              .lfStrikeOut = Abs(TBar.FontStrikethru)
              .lfOutPrecision = OUT_TT_PRECIS
              .lfQuality = ANTIALIASED_QUALITY
              .lfCharSet = DEFAULT_CHARSET
              .lfPitchAndFamily = VARIABLE_PITCH
          End With
       
        ' Dialog aufrufen
        If ChooseFont(CF_T) = 0 Then GoTo FinalizeProc
        TmpFName = StrConv(LFnt.lfFaceName, vbUnicode)
    End If
    With TBar
      If reset Then
          .FontColor = vbBlack
          .FontSize = 9
          .FontName = "Verdana"
          .FontBold = False
          .FontItalic = False
          .FontUnderline = False
          .FontStrikethru = False
      Else
          .FontColor = CF_T.rgbColors
          .FontSize = CF_T.iPointSize \ 10
          .FontName = Left$(TmpFName, InStr(1, TmpFName, vbNullChar) - 1)
          .FontBold = CBool(LFnt.lfWeight >= FW_BOLD)
          .FontItalic = CBool(LFnt.lfItalic)
          .FontUnderline = CBool(LFnt.lfUnderline)
          .FontStrikethru = CBool(LFnt.lfStrikeOut)
           SaveSetting App.Title, "Textbox", "FontName", .FontName
           SaveSetting App.Title, "Textbox", "FontBold", Abs(.FontBold)
           SaveSetting App.Title, "Textbox", "FontItalic", Abs(.FontItalic)
           SaveSetting App.Title, "Textbox", "FontUnderline", Abs(.FontUnderline)
           SaveSetting App.Title, "Textbox", "FontStrikethru", Abs(.FontStrikethru)
           SaveSetting App.Title, "Textbox", "FontSize", .FontSize
           SaveSetting App.Title, "Textbox", "Color", .FontColor
      End If
      
    End With

  SyncFontAndColor
  SendMessageL txtEditBox.hwnd, EM_SETMARGINS, EC_LEFTMARGIN, 3
  txtEditBox_Change

  
FinalizeProc:
  On Error Resume Next
  Me.SetFocus
  Exit Sub

mnuFonts_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuFonts_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
 Resume FinalizeProc
End Sub

Private Sub CreateTestImage(Optional imgIndex As Integer, Optional ShowImgPaste As Boolean)
Dim i As Integer
    Me.Width = 12000: Me.Height = 6000
    With picImage
        Set .Picture = Nothing
        .Width = (imgIndex + 1) * 300 * LTwipsPerPixelX
        .Height = (imgIndex + 1) * 200 * LTwipsPerPixelX
        .DrawMode = vbCopyPen
        .DrawStyle = vbSolid
        .DrawWidth = 1
        If imgIndex = 0 Then
            picImage.Line (0, 0)-(299, 199), vbWhite, BF
            picImage.Line (0, 0)-(299, 199), vbRed, B
            picImage.Line (1, 1)-(298, 198), vbGreen, B
            For i = 9 To 80 Step 10
                picImage.Line (i, i)-(299 - i - 1, 199 - i - 1), vbRed, B
                picImage.Line (i + 1, i + 1)-(299 - i - 2, 199 - i - 2), vbGreen, B
            Next
        ElseIf imgIndex = 1 Then
            picImage.Line (0, 0)-(599, 399), vbWhite, BF
            For i = 9 To 599 Step 10
                If i Mod 99 < 9 Then
                    If i < 400 Then picImage.Line (0, i)-(600, i), &HAAAAFF
                    picImage.Line (i, 0)-(i, 400), &HAAAAFF
                Else
                    If i < 400 Then picImage.Line (0, i)-(600, i), &HFFAAAA
                    picImage.Line (i, 0)-(i, 400), &HFFAAAA
                End If
            Next
        End If
        mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(.Image)
    End With
    TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
    Call PaintGrading
    If ShowImgPaste Then
        picPaste.Move 400 * LTwipsPerPixelX, TBar.Height, picImage.Width, picImage.Height
        picPaste.Visible = True
        picPaste.ZOrder
    Else
        picPaste.Visible = False
    End If
    
End Sub

Private Sub CropOrTearImage()
Dim l&, t&, h&, w&
    With mWorkControl
        If .DrawMode = 0 Then Exit Sub
        picImage.Line (.x1, .y1)-(.x2, .y2), , B
        AdjustingWorkControlEdges
        If .x2 < .x1 Then l = .x2 Else l = .x1
        If .y2 < .y1 Then t = .y2 Else t = .y1
        w = Abs(.x2 - .x1) + 1
        h = Abs(.y2 - .y1) + 1
        .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
    End With
    If l < 0 Then l = 0
    If t < 0 Then t = 0
    If w < 2 Then w = 2
    If h < 2 Then h = 2

    If TBar.Selected = tbCrop Then
        With picPaste
            .Width = w * LTwipsPerPixelX
            .Height = h * LTwipsPerPixelY
            .PaintPicture picImage.Image, 0, 0, w * LTwipsPerPixelX, h * LTwipsPerPixelY, l * LTwipsPerPixelX, t * LTwipsPerPixelY, w * LTwipsPerPixelX, h * LTwipsPerPixelY
            picImage.Width = .Width
            picImage.Height = .Height
            Set picImage.Picture = .Image
            Set .Picture = Nothing
            .Width = 120: .Height = 120
        End With
    ElseIf TBar.Selected = tbTear Then
        If w > h Then
            TearHorizontal t, t + h
        Else
            TearVertical l, l + w
        End If
    End If
    With picImage
        .MousePointer = vbDefault
        .DrawStyle = mDrawStyle.DrawStyle
        .DrawMode = mDrawStyle.DrawMode
        .DrawWidth = mDrawStyle.DrawWidth
        .FillStyle = mDrawStyle.FillStyle
        .ForeColor = SBar.ForeColor
         mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(.Image)
         TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
    End With
    PaintGrading
    TBar.Selected = tbPointer
    
End Sub

Private Sub DrawArrow(X As Single, Y As Single)
Dim tCursorPos As POINTAPI
Dim i As Integer
Dim iPts() As POINTAPI
    With picImage
        mDrawStyle.DrawStyle = .DrawStyle
        mDrawStyle.DrawMode = .DrawMode
        mDrawStyle.DrawWidth = .DrawWidth
        mDrawStyle.FillStyle = .FillStyle
        
        .DrawMode = vbCopyPen
        .DrawStyle = vbSolid
        .DrawWidth = 1
        .FillStyle = vbFSSolid
        
        
        Select Case TBar.Arrow
            Case 0
'                ReDim iPts(6) As POINTAPI
'                iPts(0).X = X:     iPts(0).Y = Y
'                iPts(1).X = X + 4: iPts(1).Y = Y - 4
'                iPts(2).X = X + 4: iPts(2).Y = Y - 1
'                iPts(3).X = X + 9: iPts(3).Y = Y - 1
'                iPts(4).X = X + 9: iPts(4).Y = Y + 1
'                iPts(5).X = X + 4: iPts(5).Y = Y + 1
'                iPts(6).X = X + 4: iPts(6).Y = Y + 4
'                gdiplus.PaintPolygon picImage, iPts, vbBSSolid, SBar.ForeColor, 1, SBar.Fill > 0, SBar.BackColor, IIf(SBar.Fill = 1, 50, 100)
'                Exit Sub
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X + 1, Y - 1)-(X + 1, Y + 2), SBar.ForeColor
                picImage.Line (X + 2, Y - 2)-(X + 2, Y + 3), SBar.ForeColor
                picImage.Line (X + 3, Y - 3)-(X + 3, Y + 4), SBar.ForeColor
                picImage.Line (X + 4, Y - 4)-(X + 4, Y + 5), SBar.ForeColor
                picImage.Line (X + 5, Y - 1)-(X + 9, Y + 1), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X + 19 + (SBar.Line * 2), Y
            Case 1
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X + 1, Y - 1)-(X + 1, Y + 2), SBar.ForeColor
                picImage.Line (X + 2, Y - 2)-(X + 2, Y + 3), SBar.ForeColor
                picImage.Line (X + 3, Y - 3)-(X + 3, Y + 4), SBar.ForeColor
                picImage.Line (X + 4, Y - 4)-(X + 4, Y + 5), SBar.ForeColor
                picImage.Line (X + 5, Y - 5)-(X + 5, Y + 6), SBar.ForeColor
                picImage.Line (X + 6, Y - 6)-(X + 6, Y + 7), SBar.ForeColor
                picImage.Line (X + 7, Y - 2)-(X + 12, Y + 2), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X + 22 + (SBar.Line * 2), Y
            Case 2
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X + 1, Y - 1)-(X + 1, Y + 2), SBar.ForeColor
                picImage.Line (X + 2, Y - 2)-(X + 2, Y + 3), SBar.ForeColor
                picImage.Line (X + 3, Y - 3)-(X + 3, Y + 4), SBar.ForeColor
                picImage.Line (X + 4, Y - 4)-(X + 4, Y + 5), SBar.ForeColor
                picImage.Line (X + 5, Y - 5)-(X + 5, Y + 6), SBar.ForeColor
                picImage.Line (X + 6, Y - 6)-(X + 6, Y + 7), SBar.ForeColor
                picImage.Line (X + 7, Y - 7)-(X + 7, Y + 8), SBar.ForeColor
                picImage.Line (X + 8, Y - 8)-(X + 8, Y + 9), SBar.ForeColor
                picImage.Line (X + 9, Y - 3)-(X + 15, Y + 3), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X + 24 + (SBar.Line * 2), Y
            Case 5
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y + 1)-(X + 2, Y + 1), SBar.ForeColor
                picImage.Line (X - 2, Y + 2)-(X + 3, Y + 2), SBar.ForeColor
                picImage.Line (X - 3, Y + 3)-(X + 4, Y + 3), SBar.ForeColor
                picImage.Line (X - 4, Y + 4)-(X + 5, Y + 4), SBar.ForeColor
                picImage.Line (X - 1, Y + 5)-(X + 1, Y + 9), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X, Y + 19 + (SBar.Line * 2)
            Case 6
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y + 1)-(X + 2, Y + 1), SBar.ForeColor
                picImage.Line (X - 2, Y + 2)-(X + 3, Y + 2), SBar.ForeColor
                picImage.Line (X - 3, Y + 3)-(X + 4, Y + 3), SBar.ForeColor
                picImage.Line (X - 4, Y + 4)-(X + 5, Y + 4), SBar.ForeColor
                picImage.Line (X - 5, Y + 5)-(X + 6, Y + 5), SBar.ForeColor
                picImage.Line (X - 6, Y + 6)-(X + 7, Y + 6), SBar.ForeColor
                picImage.Line (X - 2, Y + 7)-(X + 2, Y + 12), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X, Y + 22 + (SBar.Line * 2)
            Case 7
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y + 1)-(X + 2, Y + 1), SBar.ForeColor
                picImage.Line (X - 2, Y + 2)-(X + 3, Y + 2), SBar.ForeColor
                picImage.Line (X - 3, Y + 3)-(X + 4, Y + 3), SBar.ForeColor
                picImage.Line (X - 4, Y + 4)-(X + 5, Y + 4), SBar.ForeColor
                picImage.Line (X - 5, Y + 5)-(X + 6, Y + 5), SBar.ForeColor
                picImage.Line (X - 6, Y + 6)-(X + 7, Y + 6), SBar.ForeColor
                picImage.Line (X - 7, Y + 7)-(X + 8, Y + 7), SBar.ForeColor
                picImage.Line (X - 8, Y + 8)-(X + 9, Y + 8), SBar.ForeColor
                picImage.Line (X - 3, Y + 9)-(X + 3, Y + 15), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X, Y + 24 + (SBar.Line * 2)
            Case 10
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y - 1)-(X + 2, Y - 1), SBar.ForeColor
                picImage.Line (X - 2, Y - 2)-(X + 3, Y - 2), SBar.ForeColor
                picImage.Line (X - 3, Y - 3)-(X + 4, Y - 3), SBar.ForeColor
                picImage.Line (X - 4, Y - 4)-(X + 5, Y - 4), SBar.ForeColor
                picImage.Line (X - 1, Y - 5)-(X + 1, Y - 9), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X, Y - 19 - (SBar.Line * 2)
            Case 11
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y - 1)-(X + 2, Y - 1), SBar.ForeColor
                picImage.Line (X - 2, Y - 2)-(X + 3, Y - 2), SBar.ForeColor
                picImage.Line (X - 3, Y - 3)-(X + 4, Y - 3), SBar.ForeColor
                picImage.Line (X - 4, Y - 4)-(X + 5, Y - 4), SBar.ForeColor
                picImage.Line (X - 5, Y - 5)-(X + 6, Y - 5), SBar.ForeColor
                picImage.Line (X - 6, Y - 6)-(X + 7, Y - 6), SBar.ForeColor
                picImage.Line (X - 2, Y - 7)-(X + 2, Y - 12), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X, Y - 22 - (SBar.Line * 2)
            Case 12
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y - 1)-(X + 2, Y - 1), SBar.ForeColor
                picImage.Line (X - 2, Y - 2)-(X + 3, Y - 2), SBar.ForeColor
                picImage.Line (X - 3, Y - 3)-(X + 4, Y - 3), SBar.ForeColor
                picImage.Line (X - 4, Y - 4)-(X + 5, Y - 4), SBar.ForeColor
                picImage.Line (X - 5, Y - 5)-(X + 6, Y - 5), SBar.ForeColor
                picImage.Line (X - 6, Y - 6)-(X + 7, Y - 6), SBar.ForeColor
                picImage.Line (X - 7, Y - 7)-(X + 8, Y - 7), SBar.ForeColor
                picImage.Line (X - 8, Y - 8)-(X + 9, Y - 8), SBar.ForeColor
                picImage.Line (X - 3, Y - 9)-(X + 3, Y - 15), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X, Y - 24 - (SBar.Line * 2)
            Case 15
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y - 1)-(X - 1, Y + 2), SBar.ForeColor
                picImage.Line (X - 2, Y - 2)-(X - 2, Y + 3), SBar.ForeColor
                picImage.Line (X - 3, Y - 3)-(X - 3, Y + 4), SBar.ForeColor
                picImage.Line (X - 4, Y - 4)-(X - 4, Y + 5), SBar.ForeColor
                picImage.Line (X - 5, Y - 1)-(X - 9, Y + 1), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X - 19 - (SBar.Line * 2), Y
            Case 16
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y - 1)-(X - 1, Y + 2), SBar.ForeColor
                picImage.Line (X - 2, Y - 2)-(X - 2, Y + 3), SBar.ForeColor
                picImage.Line (X - 3, Y - 3)-(X - 3, Y + 4), SBar.ForeColor
                picImage.Line (X - 4, Y - 4)-(X - 4, Y + 5), SBar.ForeColor
                picImage.Line (X - 5, Y - 5)-(X - 5, Y + 6), SBar.ForeColor
                picImage.Line (X - 6, Y - 6)-(X - 6, Y + 7), SBar.ForeColor
                picImage.Line (X - 7, Y - 2)-(X - 12, Y + 2), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X - 22 - (SBar.Line * 2), Y
            Case 17
                picImage.PSet (X, Y), SBar.ForeColor
                picImage.Line (X - 1, Y - 1)-(X - 1, Y + 2), SBar.ForeColor
                picImage.Line (X - 2, Y - 2)-(X - 2, Y + 3), SBar.ForeColor
                picImage.Line (X - 3, Y - 3)-(X - 3, Y + 4), SBar.ForeColor
                picImage.Line (X - 4, Y - 4)-(X - 4, Y + 5), SBar.ForeColor
                picImage.Line (X - 5, Y - 5)-(X - 5, Y + 6), SBar.ForeColor
                picImage.Line (X - 6, Y - 6)-(X - 6, Y + 7), SBar.ForeColor
                picImage.Line (X - 7, Y - 7)-(X - 7, Y + 8), SBar.ForeColor
                picImage.Line (X - 8, Y - 8)-(X - 8, Y + 9), SBar.ForeColor
                picImage.Line (X - 9, Y - 3)-(X - 15, Y + 3), SBar.ForeColor, BF
                If TBar.SelectedEx = tbLegend Then DrawLegend X - 24 - (SBar.Line * 2), Y
            Case 4, 9, 14, 19 'Mauszeiger
                ReDim iPts(10) As POINTAPI
                iPts(0).X = X + 3:  iPts(0).Y = Y + 1
                iPts(1).X = X + 3:  iPts(1).Y = Y + 16
                iPts(2).X = X + 6:  iPts(2).Y = Y + 13
                iPts(3).X = X + 6:  iPts(3).Y = Y + 13
                iPts(4).X = X + 7:  iPts(4).Y = Y + 13
                iPts(5).X = X + 11:  iPts(5).Y = Y + 20
                iPts(6).X = X + 13: iPts(6).Y = Y + 19
                iPts(7).X = X + 10:  iPts(7).Y = Y + 13
                iPts(8).X = X + 10:  iPts(8).Y = Y + 11
                iPts(9).X = X + 14: iPts(9).Y = Y + 11
                iPts(10).X = X + 4: iPts(10).Y = Y + 1
                gdiplus.PaintPolygon picImage, iPts, vbBSSolid, vbBlack, 1, True, vbBlack, 20   'Mausschatten
                For i = 0 To UBound(iPts)
                    iPts(i).X = iPts(i).X - 3
                    iPts(i).Y = iPts(i).Y - 1
                Next i
                gdiplus.PaintPolygon picImage, iPts, vbBSSolid, vbBlack, 1, True, vbWhite, 100  'Mauszeiger
                If TBar.Arrow >= 9 Then
                    picImage.Line (X - 3, Y - 1)-(X + 3, Y - 1), SBar.ForeColor
                    picImage.PSet (X - 1, Y - 2), SBar.ForeColor:   picImage.PSet (X - 2, Y - 3), SBar.ForeColor
                    picImage.PSet (X - 1, Y + 0), SBar.ForeColor:   picImage.PSet (X - 2, Y + 1), SBar.ForeColor
                    picImage.PSet (X + 1, Y - 2), SBar.ForeColor:   picImage.PSet (X + 2, Y - 3), SBar.ForeColor
                    picImage.PSet (X, Y - 3), SBar.ForeColor:       picImage.PSet (X, Y - 4), SBar.ForeColor
                    picImage.Line (X - 5, Y + 1)-(X - 5, Y + 4), vbBlack   'V
                    picImage.Line (X - 9, Y + 4)-(X - 9, Y + 13), vbBlack   'V
                    picImage.Line (X - 2, Y + 4)-(X - 2, Y + 13), vbBlack   'V
                    picImage.Line (X - 8, Y + 4)-(X - 2, Y + 4), vbBlack    'H
                    picImage.Line (X - 8, Y + 7)-(X - 2, Y + 7), vbBlack    'H
                    picImage.Line (X - 8, Y + 13)-(X - 2, Y + 13), vbBlack  'H
                    If TBar.Arrow = 9 Then
                        picImage.Line (X - 9, Y + 4)-(X - 6, Y + 7), SBar.ForeColor, BF
                    ElseIf TBar.Arrow = 14 Then
                        picImage.Line (X - 7, Y + 4)-(X - 4, Y + 7), SBar.ForeColor, BF
                    ElseIf TBar.Arrow = 19 Then
                        picImage.Line (X - 4, Y + 4)-(X - 2, Y + 7), SBar.ForeColor, BF
                    End If
                End If
        End Select
        
        .DrawMode = mDrawStyle.DrawMode
        .DrawStyle = mDrawStyle.DrawStyle
        .FillStyle = mDrawStyle.FillStyle
        .DrawWidth = mDrawStyle.DrawWidth
        If TBar.SelectedEx <> tbLegend Then
            mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(.Image)
            TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
        End If
    End With
    
    If Not MagGlass Is Nothing Then
        DoEvents
        GetCursorPos tCursorPos
        MagGlass.PrintMagGlass tCursorPos
    End If
End Sub

Private Sub DrawCyrcle(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
Dim r As Long
    If Step = ActionStart Then
        If modMain.IsLightColor(SBar.ForeColor) Then picImage.ForeColor = &HEEEEEE
        picImage.DrawMode = vbNotXorPen
        picImage.DrawStyle = vbDash
        picImage.DrawWidth = 1
        With mWorkControl
            .x2 = .x0: .y2 = .y0
            .x1 = X: .y1 = Y
            .x0 = .x1: .y0 = .y1
            .DrawMode = tbCyrcle
        End With
    Else                'Aktion-Ende
        With mWorkControl
            .x2 = .x0: .y2 = .y0
            r = Abs(.x2 - .x1)
            If Abs(.x2 - .x1) > (SBar.Line + 1) Then picImage.Circle (.x1, .y1), r

            gdiplus.PaintShape picImage, seShapeCircle, .x1, .y1, r * 2, r * 2, vbBSSolid, SBar.ForeColor, 1 + (SBar.Line * 2), SBar.Fill > 0, SBar.BackColor, IIf(SBar.Fill = 1, 50, 100)
            
            mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
            TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
            .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
        End With
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End Sub

Private Sub DrawFill(X As Single, Y As Single)
Dim hBrush As Long
Dim tCursorPos As POINTAPI
    hBrush = CreateSolidBrush(SBar.ForeColor)
    With picImage
      SelectObject .hDC, hBrush
      ExtFloodFill .hDC, X, Y, .Point(X, Y), FLOODFILLSURFACE
      mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(.Image)
      TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
    End With
    DeleteObject hBrush
    
    If Not MagGlass Is Nothing Then
        DoEvents
        GetCursorPos tCursorPos
        MagGlass.PrintMagGlass tCursorPos
    End If
End Sub

Private Sub DrawFreehand(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
    If Step = ActionStart Then
        picImage.DrawMode = vbCopyPen
        picImage.DrawStyle = vbSolid
        picImage.DrawWidth = (SBar.Line * 2) + 2
        picImage.PSet (X, Y), SBar.ForeColor
        mWorkControl.DrawMode = tbFreehand
    Else
        mWorkControl.DrawMode = 0
        mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
        TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False

        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End Sub

Private Sub DrawLegend(X As Single, Y As Single)
Dim tCursorPos As POINTAPI
Dim deltaX As Integer, deltaY As Integer, r As Long
Dim sLegend As String
    sLegend = SBar.LegendText
    r = 18 + (SBar.Line * 4)
    With picImage
        gdiplus.PaintShape picImage, seShapeCircle, X, Y, r, r, vbBSSolid, SBar.ForeColor, IIf(SBar.Line > 1, 2, 1), SBar.Fill > 0, SBar.BackColor, IIf(SBar.Fill = 0, 80, SBar.Fill * 50)
        .ForeColor = TBar.FontColor
        .FontName = TBar.FontName
        .FontBold = TBar.FontBold
        .FontItalic = TBar.FontItalic
        .FontStrikethru = TBar.FontStrikethru
        .FontUnderline = TBar.FontUnderline
        Select Case SBar.Line
            Case 0
                .FontSize = 8
                deltaX = .TextWidth(sLegend) \ 2
                deltaY = .TextHeight(sLegend) \ 2
               TextOut .hDC, X - deltaX, Y - deltaY, sLegend, 1
            Case 1
                .FontSize = 12
                deltaX = .TextWidth(sLegend) \ 2
                deltaY = .TextHeight(sLegend) \ 2
                TextOut .hDC, X - deltaX, Y - deltaY, sLegend, 1
            Case 2
                .FontSize = 14
                deltaX = .TextWidth(sLegend) \ 2
                deltaY = .TextHeight(sLegend) \ 2
                TextOut .hDC, X - deltaX, Y - deltaY, sLegend, 1
            Case 3
                .FontSize = 14
                deltaX = .TextWidth(sLegend) \ 2
                deltaY = .TextHeight(sLegend) \ 2
                TextOut .hDC, X - deltaX, Y - deltaY, sLegend, 1
        End Select
        .FillStyle = vbFSTransparent
        .FillColor = &H0&
        .ForeColor = SBar.ForeColor
        mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(.Image), sLegend
        TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
    End With
    SBar.LegendIncrease
    If Not MagGlass Is Nothing Then
        DoEvents
        GetCursorPos tCursorPos
        MagGlass.PrintMagGlass tCursorPos
    End If
End Sub

Private Sub DrawLine(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
With mWorkControl
    If Step = ActionStart Then
        If modMain.IsLightColor(SBar.ForeColor) Then picImage.ForeColor = &HEEEEEE
        picImage.DrawMode = vbNotXorPen
        picImage.DrawStyle = vbDash
        picImage.DrawWidth = 1
        If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then
            If X <= 10 Then X = 0
            If Y <= 10 Then Y = 0
        End If
        .x0 = X: .y0 = Y: .x1 = X: .y1 = Y: .x2 = X: .y2 = Y
        .DrawMode = tbLine
        If TBar.SelectedEx = tbLegend Then ResetCursor tbLegend
    ElseIf TBar.SelectedEx = tbLegend Then  'Multi-Tool
        .x2 = .x0: .y2 = .y0
        picImage.Line (.x1, .y1)-(.x2, .y2) 'aufheben
        Call CutLine(18 + (SBar.Line * 4))
        gdiplus.PaintShape picImage, seShapeLine, .x1, .y1, .x2, .y2, vbBSSolid, SBar.ForeColor, 1 + SBar.Line, , , IIf(SBar.Fill = 0, 80, SBar.Fill * 50)
        DrawLegend X, Y
        ResetCursor tbLine
    Else                'Aktion-Ende
        .x2 = .x0: .y2 = .y0
        picImage.Line (.x1, .y1)-(.x2, .y2) 'aufheben
        gdiplus.PaintShape picImage, seShapeLine, .x1, .y1, .x2, .y2, vbBSSolid, SBar.ForeColor, 1 + (SBar.Line * 2), , , (SBar.Fill * 40) + 20
    End If
    
    If Step = ActionEnd Then
        If TBar.SelectedEx <> tbLegend Then
            mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
            TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
        End If
        .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End With
End Sub

Private Sub DrawDimension(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
Dim isVDim As Boolean
Dim txtDimension As String
Dim distUd As Single
Dim distPx As Integer
Dim distStr As String
Dim decChr As String
Dim i As Integer
'###_START_PRO_###
With mWorkControl
    If Not IsPro Then Exit Sub
    If mWorkControl.DrawMode = 0 And Step = ActionStart Then
        linDimension(0).x1 = X: linDimension(0).y1 = Y: linDimension(0).x2 = X: linDimension(0).y2 = Y
        linDimension(1).x1 = X: linDimension(1).y1 = Y: linDimension(1).x2 = X: linDimension(1).y2 = Y
        linDimension(0).Visible = True: linDimension(1).Visible = True: linDimension(2).Visible = False: shpDimension.Visible = False
        .x0 = X: .y0 = Y
        .DrawMode = tbDimension     'Step1 Bemaßung
    ElseIf mWorkControl.DrawMode > 0 And Step = ActionStart Then
        mWorkControl.DrawMode = -1 * tbDimension
        linDimension(2).Visible = True
        shpDimension.Visible = True
        picImage.Font = TBar.Font
        .x2 = picImage.TextHeight("µÎ") + 1
        linDimension(2).x1 = X: linDimension(2).y1 = Y: linDimension(2).x2 = X: linDimension(2).y2 = Y
        If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then   'vertikale Bemaßung
            txtDimension = Round((Abs(Y - .y0) + 1) * Abs(SBar.RulerScaleMulti), SBar.RulerScaleDec)
            shpDimension.Move X - .x2 - 2, Y - .x2 - 1, .x2 - 1, picImage.TextWidth(txtDimension)
            .y2 = 1 'zeigt an, dass der Bemaßungstext vertikal berechnet wurde
        Else
            txtDimension = Round((Abs(X - .x0) + 1) * Abs(SBar.RulerScaleMulti), SBar.RulerScaleDec)
            shpDimension.Move X, Y - .x2 - 2, picImage.TextWidth(txtDimension), .x2 - 1
            .y2 = 0 'zeigt an, dass der Bemaßungstext horizontal berechnet wurde
        End If
        .x1 = X: .y1 = Y            'Step2 Bemaßung
    Else
        With picImage
            .DrawWidth = 1
            .DrawMode = vbCopyPen
            .DrawStyle = vbSolid
            .ForeColor = SBar.ForeColor
        End With
        isVDim = CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED)
        If isVDim Then
            distPx = (Abs(.y1 - .y0) + 1)  'vertikale Bemaßung
        Else
            distPx = (Abs(.x1 - .x0) + 1)   'horizontale Bemaßung
        End If
        TBar.SetButtonShift False
        txtDimension = InputBox("Beschriftungstext:", "Bemaßung...", Round(distPx * Abs(SBar.RulerScaleMulti), SBar.RulerScaleDec))
        If StrPtr(txtDimension) = 0 Then
            .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
            linDimension(0).Visible = False: linDimension(1).Visible = False: linDimension(2).Visible = False: shpDimension.Visible = False
            Exit Sub
        End If
        txtDimension = Trim$(txtDimension)
        'Urechnungsfaktor und Kommastellen berechnen
        If Len(txtDimension) > 0 And SBar.RulerScaleMode = PL_USER And SBar.RulerScaleMulti = -1 Then
            distStr = txtDimension
            Do While IsNumeric(distStr) = False And Len(distStr) > 0
                distStr = Right$(distStr, (Len(distStr) - 1))
            Loop
            If Len(distStr) > 0 Then distUd = CSng(distStr)
            If distUd <> 0 And distPx <> 0 Then
                distUd = Round(distUd / distPx, 2)
                SBar.RulerScaleMulti = distUd
                decChr = Mid$(CStr(1.5), 2, 1)
                distStr = CStr(distUd)
                i = InStr(distStr, decChr)
                If i > 0 Then SBar.RulerScaleDec = Len(Mid$(distStr, i + 1)) Else SBar.RulerScaleDec = 0
            End If
            Debug.Print SBar.RulerScaleMulti & vbTab & SBar.RulerScaleDec
        End If
        'Bemaßung zeichnen
        picImage.Line (linDimension(0).x1, linDimension(0).y1)-(linDimension(0).x2, linDimension(0).y2)
        picImage.Line (linDimension(1).x1, linDimension(1).y1)-(linDimension(1).x2, linDimension(1).y2)
        picImage.Line (linDimension(2).x1, linDimension(2).y1)-(linDimension(2).x2, linDimension(2).y2)
        Select Case True    'Bemaßungspfeile
            Case isVDim And Y < .y1 And .y1 < .y0   'UOO
                picImage.Line (X - 1, .y1 - 2)-(X + 2, .y1 - 2): picImage.Line (X - 2, .y1 - 3)-(X + 3, .y1 - 3): picImage.Line (X - 3, .y1 - 4)-(X + 4, .y1 - 4)
                picImage.Line (X - 1, .y0 + 2)-(X + 2, .y0 + 2): picImage.Line (X - 2, .y0 + 3)-(X + 3, .y0 + 3): picImage.Line (X - 3, .y0 + 4)-(X + 4, .y0 + 4)
                picImage.Line (X, .y0)-(X, .y0 + 10)
            Case isVDim And Y > .y0 And .y1 < .y0   'UOU
                picImage.Line (X - 1, .y1 - 2)-(X + 2, .y1 - 2): picImage.Line (X - 2, .y1 - 3)-(X + 3, .y1 - 3): picImage.Line (X - 3, .y1 - 4)-(X + 4, .y1 - 4)
                picImage.Line (X - 1, .y0 + 2)-(X + 2, .y0 + 2): picImage.Line (X - 2, .y0 + 3)-(X + 3, .y0 + 3): picImage.Line (X - 3, .y0 + 4)-(X + 4, .y0 + 4)
                picImage.Line (X, .y1)-(X, .y1 - 10)
            Case isVDim And Y < .y0 And .y0 < .y1   'OUO
                picImage.Line (X - 2, .y0 - 2)-(X + 2, .y0 - 2): picImage.Line (X - 3, .y0 - 3)-(X + 3, .y0 - 3): picImage.Line (X - 4, .y0 - 4)-(X + 4, .y0 - 4)
                picImage.Line (X - 2, .y1 + 2)-(X + 2, .y1 + 2): picImage.Line (X - 3, .y1 + 3)-(X + 3, .y1 + 3): picImage.Line (X - 4, .y1 + 4)-(X + 4, .y1 + 4)
                picImage.Line (X, .y1)-(X, .y1 + 10)
            Case isVDim And Y > .y1 And .y0 < .y1   'OUU
                picImage.Line (X - 2, .y0 - 2)-(X + 2, .y0 - 2): picImage.Line (X - 3, .y0 - 3)-(X + 3, .y0 - 3): picImage.Line (X - 4, .y0 - 4)-(X + 4, .y0 - 4)
                picImage.Line (X - 2, .y1 + 2)-(X + 2, .y1 + 2): picImage.Line (X - 3, .y1 + 3)-(X + 3, .y1 + 3): picImage.Line (X - 4, .y1 + 4)-(X + 4, .y1 + 4)
                picImage.Line (X, .y0)-(X, .y0 - 10)
            Case isVDim And .y1 < .y0   'UOM
                picImage.Line (X - 2, .y0 - 2)-(X + 2, .y0 - 2): picImage.Line (X - 3, .y0 - 3)-(X + 3, .y0 - 3): picImage.Line (X - 4, .y0 - 4)-(X + 4, .y0 - 4)
                picImage.Line (X - 2, .y1 + 2)-(X + 2, .y1 + 2): picImage.Line (X - 3, .y1 + 3)-(X + 3, .y1 + 3): picImage.Line (X - 4, .y1 + 4)-(X + 4, .y1 + 4)
            Case isVDim                 'OUM
                picImage.Line (X - 1, .y1 - 2)-(X + 2, .y1 - 2): picImage.Line (X - 2, .y1 - 3)-(X + 3, .y1 - 3): picImage.Line (X - 3, .y1 - 4)-(X + 4, .y1 - 4)
                picImage.Line (X - 1, .y0 + 2)-(X + 2, .y0 + 2): picImage.Line (X - 2, .y0 + 3)-(X + 3, .y0 + 3): picImage.Line (X - 3, .y0 + 4)-(X + 4, .y0 + 4)
            Case X < .x1 And .x1 < .x0
                picImage.Line (.x1 - 2, Y - 1)-(.x1 - 2, Y + 2): picImage.Line (.x1 - 3, Y - 2)-(.x1 - 3, Y + 3): picImage.Line (.x1 - 4, Y - 3)-(.x1 - 4, Y + 4)
                picImage.Line (.x0 + 2, Y - 1)-(.x0 + 2, Y + 2): picImage.Line (.x0 + 3, Y - 2)-(.x0 + 3, Y + 3): picImage.Line (.x0 + 4, Y - 3)-(.x0 + 4, Y + 4)
                picImage.Line (.x0, Y)-(.x0 + 10, Y)
            Case X > .x0 And .x1 < .x0
                picImage.Line (.x1 - 2, Y - 1)-(.x1 - 2, Y + 2): picImage.Line (.x1 - 3, Y - 2)-(.x1 - 3, Y + 3): picImage.Line (.x1 - 4, Y - 3)-(.x1 - 4, Y + 4)
                picImage.Line (.x0 + 2, Y - 1)-(.x0 + 2, Y + 2): picImage.Line (.x0 + 3, Y - 2)-(.x0 + 3, Y + 3): picImage.Line (.x0 + 4, Y - 3)-(.x0 + 4, Y + 4)
                picImage.Line (.x1, Y)-(.x1 - 10, Y)
            Case X < .x0 And .x0 < .x1
                picImage.Line (.x0 - 2, Y - 1)-(.x0 - 2, Y + 2): picImage.Line (.x0 - 3, Y - 2)-(.x0 - 3, Y + 3): picImage.Line (.x0 - 4, Y - 3)-(.x0 - 4, Y + 4)
                picImage.Line (.x1 + 2, Y - 1)-(.x1 + 2, Y + 2): picImage.Line (.x1 + 3, Y - 2)-(.x1 + 3, Y + 3): picImage.Line (.x1 + 4, Y - 3)-(.x1 + 4, Y + 4)
                picImage.Line (.x1, Y)-(.x1 + 10, Y)
            Case X > .x1 And .x0 < .x1
                picImage.Line (.x0 - 2, Y - 1)-(.x0 - 2, Y + 2): picImage.Line (.x0 - 3, Y - 2)-(.x0 - 3, Y + 3): picImage.Line (.x0 - 4, Y - 3)-(.x0 - 4, Y + 4)
                picImage.Line (.x1 + 2, Y - 1)-(.x1 + 2, Y + 2): picImage.Line (.x1 + 3, Y - 2)-(.x1 + 3, Y + 3): picImage.Line (.x1 + 4, Y - 3)-(.x1 + 4, Y + 4)
                picImage.Line (.x0, Y)-(.x0 - 10, Y)
            Case .x1 < .x0  'Innen
                picImage.Line (.x0 - 2, Y - 1)-(.x0 - 2, Y + 2): picImage.Line (.x0 - 3, Y - 2)-(.x0 - 3, Y + 3): picImage.Line (.x0 - 4, Y - 3)-(.x0 - 4, Y + 4)
                picImage.Line (.x1 + 2, Y - 1)-(.x1 + 2, Y + 2): picImage.Line (.x1 + 3, Y - 2)-(.x1 + 3, Y + 3): picImage.Line (.x1 + 4, Y - 3)-(.x1 + 4, Y + 4)
            Case Else       'Innen
                picImage.Line (.x0 + 2, Y - 1)-(.x0 + 2, Y + 2): picImage.Line (.x0 + 3, Y - 2)-(.x0 + 3, Y + 3): picImage.Line (.x0 + 4, Y - 3)-(.x0 + 4, Y + 4)
                picImage.Line (.x1 - 2, Y - 1)-(.x1 - 2, Y + 2): picImage.Line (.x1 - 3, Y - 2)-(.x1 - 3, Y + 3): picImage.Line (.x1 - 4, Y - 3)-(.x1 - 4, Y + 4)
        End Select
        .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
        linDimension(0).Visible = False: linDimension(1).Visible = False: linDimension(2).Visible = False: shpDimension.Visible = False
        If isVDim Then
            Dim hFont As Long, fontMem As Long, res As Long
            hFont = CreateFont(picImage.FontSize * 1.55, 0, 900, 0, IIf(picImage.FontBold, 700, 0), _
                    picImage.FontItalic, picImage.FontUnderline, picImage.FontStrikethru, 1, 4, &H10, 2, 4, picImage.FontName)
            fontMem = SelectObject(picImage.hDC, hFont)
            res = TextOut(picImage.hDC, shpDimension.Left, shpDimension.Top + (1.5 * shpDimension.Width), txtDimension, Len(txtDimension))
            res = SelectObject(picImage.hDC, fontMem)
            res = DeleteObject(hFont)
        Else
            TextOut picImage.hDC, shpDimension.Left, shpDimension.Top, txtDimension, Len(txtDimension)
        End If
        mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
        TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
        
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End With
'###_END_PRO_###
End Sub


Private Sub DrawMarker(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
    If Step = ActionStart Then
        picImage.DrawMode = vbMaskPen
        picImage.DrawStyle = vbSolid
        picImage.DrawWidth = (SBar.Line * 6) + 6
        picImage.PSet (X, Y), SBar.ForeColor
        mWorkControl.DrawMode = tbMarker
        If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then mWorkControl.y0 = Y
    Else                'Aktion-Ende
        mWorkControl.DrawMode = 0
        mWorkControl.y0 = 0
        picImage.DrawWidth = SBar.Line + 1
        mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
        TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End Sub

Private Sub DrawRectangle(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
    If Step = ActionStart Then
        picImage.DrawMode = vbNotXorPen
        picImage.DrawStyle = vbDash
        If modMain.IsLightColor(SBar.ForeColor) Then picImage.ForeColor = &HEEEEEE
        picImage.DrawMode = vbNotXorPen
        picImage.DrawStyle = vbDash
        picImage.DrawWidth = 1
        With mWorkControl
            .x0 = X: .y0 = Y: .x1 = X: .y1 = Y: .x2 = X: .y2 = Y
            .DrawMode = tbRectangle
        End With
    Else                'Aktion-Ende
        With mWorkControl
            .x2 = .x0: .y2 = .y0
            picImage.Line (.x1, .y1)-(.x2, .y2), , B    'aufheben
            If .x0 <> .x1 And .y0 <> .y1 Then
                gdiplus.PaintShape picImage, seShapeRectangle, .x1, .y1, .x2 - .x1, .y2 - .y1, vbBSSolid, SBar.ForeColor, 1 + (SBar.Line * 2), SBar.Fill > 0, SBar.BackColor, IIf(SBar.Fill = 1, 50, 100)
                mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
                TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
            End If
            .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
        End With
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End Sub

Private Sub DrawObfus(X As Single, Y As Single, Optional Step As eAction)
Dim tCursorPos As POINTAPI
Dim p As StdPicture
Dim l&, t&, h&, w&
    If Step = ActionStart Then
        picImage.DrawMode = vbNotXorPen
        picImage.DrawStyle = vbDash
        If modMain.IsLightColor(SBar.ForeColor) Then picImage.ForeColor = &HEEEEEE
        picImage.DrawMode = vbNotXorPen
        picImage.DrawStyle = vbDash
        picImage.DrawWidth = 1
        With mWorkControl
            .x0 = X: .y0 = Y: .x1 = X: .y1 = Y: .x2 = X: .y2 = Y
            .DrawMode = tbObfus
        End With
    Else                'Aktion-Ende
        With mWorkControl
            .x2 = .x0: .y2 = .y0
            picImage.Line (.x1, .y1)-(.x2, .y2), , B    'aufheben
            If .x0 <> .x1 And .y0 <> .y1 Then
                AdjustingWorkControlEdges
                If .x2 < .x1 Then l = .x2 Else l = .x1
                If .y2 < .y1 Then t = .y2 Else t = .y1
                w = Abs(.x2 - .x1) + 1
                h = Abs(.y2 - .y1) + 1
                With picPaste
                    .Width = w * LTwipsPerPixelX
                    .Height = h * LTwipsPerPixelY
                    .PaintPicture picImage.Image, 0, 0, w * LTwipsPerPixelX, h * LTwipsPerPixelY, l * LTwipsPerPixelX, t * LTwipsPerPixelY, w * LTwipsPerPixelX, h * LTwipsPerPixelY
                    Set p = gdiplus.BlurPicture(.Image, SBar.Line + 8)
                    .Picture = p
                    picImage.PaintPicture .Image, x1:=l, y1:=t
                    Set .Picture = Nothing
                    .Width = 120: .Height = 120
                End With
                mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
                TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
            End If
            .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
        End With
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End Sub

Private Sub DrawText(Optional X As Single, Optional Y As Single, Optional Step As eAction = ActionEnd)
Dim tCursorPos As POINTAPI
Dim l As Integer
Static eingabe As String
    If Step = ActionStart Then
        eingabe = InputBox("Text:", "Text erfassen...", eingabe)
        If StrPtr(eingabe) = 0 Then Exit Sub
        With txtEditBox
            .Text = eingabe
            .Move X, Y, picImage.TextWidth(eingabe) + mTextOverhang
            .Visible = True
            .SetFocus
            SendMessageL .hwnd, EM_SETMARGINS, EC_LEFTMARGIN, 3&
            picImage.MousePointer = vbDefault
        End With
    Else    'Aktion-Ende
        With txtEditBox
            If Len(.Text) Then
                If SBar.Fill > 0 Then
                    If SBar.Line = 0 Then
                        gdiplus.PaintShape picImage, seShapeRectangle, .Left, .Top, .Width - 2, .Height + SBar.Line, vbTransparent, SBar.ForeColor, SBar.Line, SBar.Fill > 0, SBar.BackColor, IIf(SBar.Fill = 1, 50, 100)
                        picImage.DrawWidth = 1
                        picImage.DrawStyle = vbDot
                        picImage.Line (.Left, .Top)-(.Left + .Width - 2, .Top + .Height + SBar.Line), SBar.ForeColor, B
                    Else
                        gdiplus.PaintShape picImage, seShapeRectangle, .Left, .Top, .Width - 2, .Height + SBar.Line, vbBSSolid, SBar.ForeColor, SBar.Line, SBar.Fill > 0, SBar.BackColor, IIf(SBar.Fill = 1, 50, 100)
                    End If
                End If
                picImage.ForeColor = TBar.FontColor
                picImage.Font = txtEditBox.Font
                picImage.FontSize = txtEditBox.FontSize
                l = SBar.Line - 1: If l < 1 Then l = 1
                TextOut picImage.hDC, .Left + 3, .Top, .Text, Len(.Text)
                mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
                TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
            End If
            .Visible = False
        End With
        Call ResetCursor
        If Not MagGlass Is Nothing Then
            DoEvents
            GetCursorPos tCursorPos
            MagGlass.PrintMagGlass tCursorPos
        End If
    End If
End Sub

Private Sub Extend()
Dim ret As String, s As String
Dim e(3) As Integer, i As Integer, j As Integer

    ret = Trim$(InputBox("Anzahl Pixel mit bieliebigen Trennzeichen für" & vbCrLf & "Oben, Rechts, Unten, Links eingeben:", "Bild erweitern...", "10,10,10,10"))
    If Len(ret) = 0 Then Exit Sub
    Do
        i = i + 1
        If i > Len(ret) Then Exit Do
        If IsNumeric(Mid$(ret, i, 1)) Then
            s = s & Mid$(ret, i, 1)
        Else
            e(j) = Val(s)
            s = ""
            j = j + 1
            If j > UBound(e) Then Exit Do
            If i > Len(ret) Then Exit Do
            Do Until IsNumeric(Mid$(ret, i + 1, 1))
                i = i + 1
            Loop
        End If
    Loop
    If j <= UBound(e) And IsNumeric(s) Then e(j) = Val(s)
    If e(0) = 0 And e(1) = 0 And e(2) = 0 And e(3) = 0 Then Exit Sub
    picPaste.Picture = picImage.Image
    With picImage
        .Cls
        .Width = .Width + (e(1) + e(3)) * LTwipsPerPixelX
        .Height = .Height + (e(0) + e(2)) * LTwipsPerPixelY
        .DrawMode = vbCopyPen
        .DrawStyle = vbSolid
        picImage.Line (0, 0)-(.Width, .Height), SBar.BackColor, BF
    End With
    picImage.PaintPicture picPaste.Image, e(3), e(0)
    Call PaintGrading
    Set picPaste.Picture = Nothing
    mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
    TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False

End Sub



Private Sub FixPaste()
    If Not picPaste.Visible Then Exit Sub
    With mWorkControl
        'Berechnen der Einfügepunktes
        .x0 = picPaste.Left
        .y0 = picPaste.Top - TBar.Height
        'Berechnen der Ziel-Größe
        .x1 = .x0 + picPaste.Width
        .y1 = .y0 + picPaste.Height
        'Berechnen der verfügbaren Fenster-Breite
        .x2 = Me.ScaleWidth
        .y2 = Me.ScaleHeight - TBar.Height - SBar.Height
        'Ggf. Ziehlgröße auf die Fenstergröße reduzieren
        If .x1 > .x2 Then .x1 = .x2
        If .y1 > .y2 Then .y1 = .y2
        'Anpassen des Ziels
        If picImage.Width < .x1 Then picImage.Width = .x1
        If picImage.Height < .y1 Then picImage.Height = .y1
        If picImage.Width > .x2 Then picImage.Width = .x2
        If picImage.Height > .y2 Then picImage.Height = .y2
        Set picImage.Picture = Nothing
        picImage.PaintPicture picPaste.Image, x1:=.x0 \ LTwipsPerPixelX, y1:=.y0 \ LTwipsPerPixelY, Width2:=.x2, Height2:=.y2
        .x0 = 0: .x1 = 0: .x2 = 0: .y0 = 0: .y1 = 0: .y2 = 0
    End With

    With picPaste
        .Visible = False
         Set .Picture = Nothing
        .Width = 1
        .Height = 1
        .Visible = False
    End With
    Call PaintGrading
    mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
    TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
    TBar.Selected = tbPointer
End Sub

'=====FORM========================================
Private Sub Form_Activate()
Dim cancel As Boolean
    TBar.SetButtonShift CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED)
    TBar_Change TBar.Selected, tbPointer, cancel
End Sub

Private Sub Form_Click()
    On Error GoTo Form_Click_Error
    If TBar.Selected = tbPaste Then FixPaste
    Exit Sub
    
Form_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.Form_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub Form_Deactivate()
    TBar.SetButtonShift False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyDown_Error
    If KeyCode = vbKeyShift Then
        TBar.SetButtonShift True
    ElseIf KeyCode = vbKeyF1 Then
        ShellExec "https://docs.ww-a.de/doku.php/pixellineal:bildeditor", vbNormalFocus
    ElseIf KeyCode = vbKeyEscape Then
        With mWorkControl
            '###_START_PRO_###
            If Abs(.DrawMode) = tbDimension Then  'Abbruch Bemaßung
                .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
                linDimension(0).Visible = False: linDimension(1).Visible = False: linDimension(2).Visible = False: shpDimension.Visible = False
            End If
            '###_END_PRO_###
        End With
    End If
    Exit Sub
    
Form_KeyDown_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.Form_KeyDown." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then TBar.SetButtonShift False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modMenuColor.Set_MenuColor(nfoSysMenuColor, Me.hwnd, &HF0F0F0)
    Call ReadSettings
    
    txtEditBox.Text = ""
    TBar.Enabled(tbUndo) = False
    TBar.Enabled(tbRedo) = False
    picImage.ZOrder 1
    picPaste.ZOrder
    SBar.ZOrder
    TBar.ZOrder
End Sub

Public Sub ReadSettings()
Dim i As Integer
    On Error Resume Next
    mGradingVisible = CBool(Val(GetSetting(App.Title, "Editor", "Grading", 1)))
    i = Abs(CInt(GetSetting(App.Title, "Editor", "Tool", 0)))
    If i > tbMenu Then i = 0
    TBar.Selected = i
    TBar.FontColor = CLng(GetSetting(App.Title, "Textbox", "Color", 0))
    TBar.FontBold = CBool(GetSetting(App.Title, "Textbox", "FontBold", False))
    TBar.FontItalic = GetSetting(App.Title, "Textbox", "FontItalic", 0)
    TBar.FontUnderline = GetSetting(App.Title, "Textbox", "FontUnderline", 0)
    TBar.FontStrikethru = GetSetting(App.Title, "Textbox", "FontStrikethru", 0)
    TBar.FontSize = CInt(GetSetting(App.Title, "Textbox", "FontSize", 9))
    TBar.FontName = GetSetting(App.Title, "Textbox", "FontName", "Verdana")
    TBar.Arrow = Abs(CInt(GetSetting(App.Title, "Editor", "Arrow", 0)))
    SBar.Line = Abs(CInt(GetSetting(App.Title, "Editor", "LineWidth", 0)))
    SBar.ForeColor = CLng(GetSetting(App.Title, "Editor", "ForeColor", 0))
    SBar.BackColor = CLng(GetSetting(App.Title, "Editor", "BackColor", vbWhite))
    SBar.Fill = CLng(GetSetting(App.Title, "Editor", "Fill", 0))
    i = CLng(GetSetting(App.Title, "Editor", "Palette", 0))
    If i >= 8 Then i = 0
    SBar.Palette = i
    '###_START_PRO_###
    linDimension(0).BorderColor = SBar.ForeColor
    linDimension(1).BorderColor = SBar.ForeColor
    linDimension(2).BorderColor = SBar.ForeColor
    shpDimension.BorderColor = SBar.ForeColor
    '###_END_PRO_###
    Call SyncFontAndColor
End Sub

Public Sub SaveSettings()
    On Error Resume Next
    If mGradingVisible <> CBool(Val(GetSetting(App.Title, "Editor", "Grading", 1))) Then SaveSetting App.Title, "Editor", "Grading", mGradingVisible
    With TBar
        If .Selected <> Abs(CInt(GetSetting(App.Title, "Editor", "Tool", 0))) Then SaveSetting App.Title, "Editor", "Tool", .Selected
        If .FontColor <> CLng(GetSetting(App.Title, "Textbox", "Color", 0)) Then SaveSetting App.Title, "Textbox", "Color", .FontColor
        If .FontBold <> CBool(GetSetting(App.Title, "Textbox", "FontBold", False)) Then SaveSetting App.Title, "Textbox", "FontBold", Abs(.FontBold)
        If .FontItalic <> GetSetting(App.Title, "Textbox", "FontItalic", 0) Then SaveSetting App.Title, "Textbox", "FontItalic", Abs(.FontItalic)
        If .FontUnderline <> GetSetting(App.Title, "Textbox", "FontUnderline", 0) Then SaveSetting App.Title, "Textbox", "FontUnderline", Abs(.FontUnderline)
        If .FontStrikethru <> GetSetting(App.Title, "Textbox", "FontStrikethru", 0) Then SaveSetting App.Title, "Textbox", "FontStrikethru", Abs(.FontStrikethru)
        If .FontSize <> CInt(GetSetting(App.Title, "Textbox", "FontSize", 9)) Then SaveSetting App.Title, "Textbox", "FontSize", .FontSize
        If .FontName <> GetSetting(App.Title, "Textbox", "FontName", "Verdana") Then SaveSetting App.Title, "Textbox", "FontName", .FontName
        If .Arrow <> Abs(CInt(GetSetting(App.Title, "Editor", "Arrow", 0))) Then SaveSetting App.Title, "Editor", "Arrow", .Arrow
    End With
    With SBar
        If .Line <> Abs(CInt(GetSetting(App.Title, "Editor", "LineWidth", 0))) Then SaveSetting App.Title, "Editor", "LineWidth", .Line
        If .ForeColor <> CLng(GetSetting(App.Title, "Editor", "ForeColor", 0)) Then SaveSetting App.Title, "Editor", "ForeColor", .ForeColor
        If .BackColor <> CLng(GetSetting(App.Title, "Editor", "BackColor", vbWhite)) Then SaveSetting App.Title, "Editor", "BackColor", .BackColor
        If .Fill <> CLng(GetSetting(App.Title, "Editor", "Fill", 0)) Then SaveSetting App.Title, "Editor", "Fill", .Fill
        If .Palette <> CLng(GetSetting(App.Title, "Editor", "Palette", 0)) Then SaveSetting App.Title, "Editor", "Palette", .Palette
    End With

End Sub



Private Sub Form_Unload(cancel As Integer)
Dim f As Form
    On Error Resume Next
    Call SaveSettings

    If frmRuler.Visible Then
        Exit Sub
    Else
        For Each f In Forms 'Prüfen ob die Anwendung geschlossen werden kann
            If TypeOf f Is frmCapture Then Exit Sub
            If TypeOf f Is frmImage And Not f Is Me Then Exit Sub
        Next
        If Not MagGlass Is Nothing Then
            Unload MagGlass
            Set MagGlass = Nothing
        End If
        Set f = Nothing
        modMain.CloseApp = True
    End If
End Sub


Private Sub MakeBorder(ByVal Index As tbBorder, Optional crUndoStep As Boolean = True)
Dim w As Long, h As Long, tw As Long, th As Long
Dim i As Integer, c As Long
Dim iPts(2) As POINTAPI
    w = picImage.Width
    h = picImage.Height
    c = SBar.ForeColor
    picPaste.BackColor = vbWhite
    Select Case Index
        
        Case tbbBorder  'Rahmen
            w = w + (2 * (SBar.Line + 1) * LTwipsPerPixelX)
            h = h + 2 * ((SBar.Line + 1) * LTwipsPerPixelY)
            With picPaste
                .Cls
                .Width = w: .Height = h
                .DrawWidth = 1: .DrawStyle = vbSolid
                .PaintPicture picImage.Image, (1 + SBar.Line) * LTwipsPerPixelX, (1 + SBar.Line) * LTwipsPerPixelY
                picPaste.Line (0, 0)-(w - LTwipsPerPixelX, h - LTwipsPerPixelY), c, B
                If SBar.Line > 0 Then picPaste.Line (LTwipsPerPixelX, LTwipsPerPixelY)-(w - (2 * LTwipsPerPixelX), h - (2 * LTwipsPerPixelY)), c, B
                If SBar.Line > 1 Then picPaste.Line (2 * LTwipsPerPixelX, 2 * LTwipsPerPixelX)-(w - (3 * LTwipsPerPixelX), h - (3 * LTwipsPerPixelY)), c, B
                picImage.Width = .Width
                picImage.Height = .Height
                picImage.Picture = .Image
            End With
            Call PaintGrading
        Case tbbShadow  'Schatten
            With picPaste
                .Cls
                .Width = w + (2 * (SBar.Line + 2) * LTwipsPerPixelX)
                .Height = h + 2 * ((SBar.Line + 2) * LTwipsPerPixelY)
                .DrawWidth = 1: .DrawStyle = vbSolid
                .PaintPicture picImage.Image, LTwipsPerPixelX, LTwipsPerPixelY
                picPaste.Line (0, 0)-(w + LTwipsPerPixelX, h + LTwipsPerPixelY), c, B                    'R
                c = modMain.Lighten(c)
                picPaste.Line (2 * LTwipsPerPixelX, h + (2 * LTwipsPerPixelY))-(w + (3 * LTwipsPerPixelX), h + (2 * LTwipsPerPixelY)), c 'H1
                picPaste.Line (w + (2 * LTwipsPerPixelX), LTwipsPerPixelY)-(w + (2 * LTwipsPerPixelX), h + (2 * LTwipsPerPixelY)), c   'V1
                c = modMain.Lighten(c)
                picPaste.Line (2 * LTwipsPerPixelX, h + (3 * LTwipsPerPixelY))-(w + (4 * LTwipsPerPixelX), h + (3 * LTwipsPerPixelY)), c 'H1
                picPaste.Line (w + (3 * LTwipsPerPixelX), 2 * LTwipsPerPixelY)-(w + (3 * LTwipsPerPixelX), h + (3 * LTwipsPerPixelY)), c 'V1
                If SBar.Line > 0 Then
                    c = modMain.Lighten(c)
                    picPaste.Line (4 * LTwipsPerPixelX, h + (4 * LTwipsPerPixelY))-(w + (5 * LTwipsPerPixelX), h + (4 * LTwipsPerPixelY)), c 'H1
                    picPaste.Line (w + (4 * LTwipsPerPixelX), 3 * LTwipsPerPixelY)-(w + (4 * LTwipsPerPixelX), h + (4 * LTwipsPerPixelY)), c 'V1
                    c = modMain.Lighten(c)
                    picPaste.Line (5 * LTwipsPerPixelX, h + (5 * LTwipsPerPixelY))-(w + (6 * LTwipsPerPixelX), h + (5 * LTwipsPerPixelY)), c 'H1
                    picPaste.Line (w + (5 * LTwipsPerPixelX), 5 * LTwipsPerPixelY)-(w + (5 * LTwipsPerPixelX), h + (5 * LTwipsPerPixelY)), c 'V1
                End If
                If SBar.Line > 1 Then
                    c = modMain.Lighten(c)
                    picPaste.Line (8 * LTwipsPerPixelX, h + (6 * LTwipsPerPixelY))-(w + (4 * LTwipsPerPixelX), h + (6 * LTwipsPerPixelY)), c 'H1
                    picPaste.Line (w + (6 * LTwipsPerPixelX), 7 * LTwipsPerPixelY)-(w + (6 * LTwipsPerPixelX), h + (4 * LTwipsPerPixelY)), c 'V1
                    c = modMain.Lighten(c)
                    picPaste.Line (9 * LTwipsPerPixelX, h + (7 * LTwipsPerPixelY))-(w + (2 * LTwipsPerPixelX), h + (7 * LTwipsPerPixelY)), c 'H1
                    picPaste.Line (w + (7 * LTwipsPerPixelX), 10 * LTwipsPerPixelY)-(w + (7 * LTwipsPerPixelX), h + (2 * LTwipsPerPixelY)), c 'V1
                End If
                picImage.Width = .Width
                picImage.Height = .Height
                picImage.Picture = .Image
            End With
            Call PaintGrading
        Case tbbTearTop  'Abriss oben
            tw = frmMenu.picTearOff(0).Width * LTwipsPerPixelX: th = frmMenu.picTearOff(0).Height * LTwipsPerPixelY
            For i = 0 To w + tw Step tw
                TransparentBlt hDC:=picImage.hDC, X:=i \ LTwipsPerPixelX, Y:=0, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                               hSrcDC:=frmMenu.picTearOff(2).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
            Next
        Case tbbTearRight  'Abriss rechts
            tw = frmMenu.picTearOff(1).Width * LTwipsPerPixelX: th = frmMenu.picTearOff(1).Height * LTwipsPerPixelY
            For i = 0 To h + th Step th
                TransparentBlt hDC:=picImage.hDC, X:=(w - tw) \ LTwipsPerPixelX, Y:=i \ LTwipsPerPixelY, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                               hSrcDC:=frmMenu.picTearOff(1).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
            Next
        Case tbbTearBottom  'Abriss unten
            tw = frmMenu.picTearOff(0).Width * LTwipsPerPixelX: th = frmMenu.picTearOff(0).Height * LTwipsPerPixelY
            For i = 0 To w + tw Step tw
                TransparentBlt hDC:=picImage.hDC, X:=i \ LTwipsPerPixelX, Y:=(h - th) \ LTwipsPerPixelY, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                               hSrcDC:=frmMenu.picTearOff(0).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
            Next
        Case tbbTearLeft    'Abriss links
            tw = frmMenu.picTearOff(1).Width * LTwipsPerPixelX: th = frmMenu.picTearOff(1).Height * LTwipsPerPixelY
            For i = 0 To h + th Step th
                TransparentBlt hDC:=picImage.hDC, X:=0, Y:=i \ LTwipsPerPixelY, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                               hSrcDC:=frmMenu.picTearOff(3).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
            Next
        Case tbbTearMiddle    'Abriss mitte
            picImage.MousePointer = vbCrosshair
            TBar.Selected = tbTear
        Case tbbTearTopRight  'Abriss oben-rechts
            MakeBorder tbbTearTop, False
            MakeBorder tbbTearRight, False
            iPts(0).X = w \ LTwipsPerPixelX - 6:   iPts(0).Y = 0
            iPts(1).X = w \ LTwipsPerPixelX:       iPts(1).Y = 0
            iPts(2).X = w \ LTwipsPerPixelX:       iPts(2).Y = 6
            gdiplus.PaintPolygon picImage, iPts, vbBSSolid, &HDDDDDD, 0, True, vbWhite, 70
        Case tbbTearBottomRight  'Abriss unten-rechts
            MakeBorder tbbTearBottom, False
            MakeBorder tbbTearRight, False
            iPts(0).X = w \ LTwipsPerPixelX - 8:   iPts(0).Y = h \ LTwipsPerPixelY
            iPts(1).X = w \ LTwipsPerPixelX:       iPts(1).Y = h \ LTwipsPerPixelY
            iPts(2).X = w \ LTwipsPerPixelX:       iPts(2).Y = h \ LTwipsPerPixelY - 8
            gdiplus.PaintPolygon picImage, iPts, vbBSSolid, &HDDDDDD, 0, True, vbWhite, 70
        Case tbbTearBottomLeft   'Abriss unten-links
            MakeBorder tbbTearBottom, False
            MakeBorder tbbTearLeft, False
            iPts(0).X = 0:   iPts(0).Y = h \ LTwipsPerPixelY - 8
            iPts(1).X = 8:   iPts(1).Y = h \ LTwipsPerPixelY
            iPts(2).X = 0:   iPts(2).Y = h \ LTwipsPerPixelY
            gdiplus.PaintPolygon picImage, iPts, vbBSSolid, &HDDDDDD, 0, True, vbWhite, 70
        Case tbbTearTopLeft   'Abriss oben-links
            MakeBorder tbbTearTop, False
            MakeBorder tbbTearLeft, False
            iPts(0).X = 0:   iPts(0).Y = 0
            iPts(1).X = 8:   iPts(1).Y = 0
            iPts(2).X = 0:   iPts(2).Y = 8
            gdiplus.PaintPolygon picImage, iPts, vbBSSolid, &HDDDDDD, 0, True, vbWhite, 70
    End Select
    If crUndoStep Then
        mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
        TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
    End If
End Sub

Private Sub PaintGrading()
Dim i As Long, w  As Long, h As Long, v As Long, t As Long
    Me.Cls
    If mGradingVisible Then
        t = TBar.Height / LTwipsPerPixelY
        w = (picImage.Width / LTwipsPerPixelX) + 1
        h = (picImage.Height / LTwipsPerPixelY) + t + 1
        With Me
            .Cls
            .ScaleMode = vbPixels
            .ForeColor = &H8000000C
            .DrawStyle = vbSolid
            .DrawWidth = 1
            .DrawMode = vbCopyPen
            .CurrentY = h + 10
            .FontName = "Arial"
            .FontSize = 6
            'Horizontal
            v = (.TextWidth("000") \ 2) + 1
            For i = 1 To w Step 2
                Line (i, h)-(i, h + 2)
            Next
            For i = 9 To w Step 10
                If (i + 1) Mod 100 = 0 Then
                     Line (i, h + 2)-(i, h + 10)
                    .CurrentX = i - v
                    Print Round((i + 1) * Abs(SBar.RulerScaleMulti), SBar.RulerScaleDec)
                Else
                    Line (i, h + 2)-(i, h + 7)
                End If
            Next
            'Vertikal
            .CurrentX = w + 10
            v = (.TextHeight("0") \ 2)
            For i = t + 1 To h Step 2
                Line (w, i)-(w + 2, i)
            Next
            h = h - t
            For i = 9 To h Step 10
                Line (w, i + t)-(w + 7, i + t)
                If (i + 1) Mod 100 = 0 Then
                    Line (w, i + t)-(w + 10, i + t)
                    .CurrentY = i + t - v
                    Print Round((i + 1) * Abs(SBar.RulerScaleMulti), SBar.RulerScaleDec)
                Else
                    Line (w, i + t)-(w + 7, i + t)
                End If
            Next
            .ScaleMode = vbTwips
        End With
    End If
End Sub


Private Sub ResetCursor(Optional ByVal newValue As tbButtons = -1)
    If newValue < 0 Then newValue = TBar.Selected
    Select Case newValue
        Case tbFreehand, tbLine, tbRectangle, tbCyrcle
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(50 + SBar.Line).Picture
        Case tbMarker
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(60 + SBar.Line).Picture
        Case tbText
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(70).Picture
        Case tbFill
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(71).Picture
        Case tbArrow
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(TBar.Arrow).Picture
        Case tbLegend
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(40 + SBar.Line).Picture
        Case tbCrop, tbObfus, tbDimension
            picImage.MousePointer = vbCustom
            picImage.MouseIcon = curPointer(106).Picture
        Case Else
            picImage.MousePointer = vbDefault
    End Select
End Sub


Private Sub SBar_ChangeScaleMode()
    On Error Resume Next
    PaintGrading
End Sub

Private Sub SBar_Click(Button As sbButtons)
Dim isMagGlass As Boolean
    On Error GoTo SBar_Click_Error
    Select Case Button
        Case sbLine0, sbLine1, sbLine2, sbLine3
            picImage.DrawWidth = (SBar.Line * 2) + 2
            SaveSetting App.Title, "Editor", "LineWidth", SBar.Line
            Call ResetCursor
        Case sbForeColor
            SaveSetting App.Title, "Editor", "ForeColor", SBar.ForeColor
            '###_START_PRO_###
            linDimension(0).BorderColor = SBar.ForeColor
            linDimension(1).BorderColor = SBar.ForeColor
            linDimension(2).BorderColor = SBar.ForeColor
            shpDimension.BorderColor = SBar.ForeColor
            '###_END_PRO_###
        Case sbBackColor
            SaveSetting App.Title, "Editor", "BackColor", SBar.BackColor
            SyncFontAndColor
        Case sbFill0, sbFill1, sbFill2
            SaveSetting App.Title, "Editor", "Fill", SBar.Fill
            SyncFontAndColor
        Case sbPicker
            If Not MagGlass Is Nothing Then
                isMagGlass = True
                MagGlass.Visible = False
                Set MagGlass = Nothing
            End If
            Set MagColor = New frmMagColor
            Set MagColor.PicColorTarget = SBar.GetPickerColor
            MagColor.Show vbModal, Me
            SBar.GetPickerColor reset:=True
            Unload MagColor
            If MagColor.PipColor <> &H1000000 Then
                If MagColor.PipColor > 0 Then
                    SBar.ForeColor = MagColor.PipColor
                    CopyRGB MagColor.PipColor, False
                Else
                    SBar.BackColor = MagColor.PipColor * -1
                    CopyRGB MagColor.PipColor
                End If
                SyncFontAndColor
            End If
            Set MagColor = Nothing
        Case sbPalette
           SaveSetting App.Title, "Editor", "Palette", SBar.Palette
    End Select
    
SBar_Click_Resume:
    On Error Resume Next
    If isMagGlass Then
        Set MagGlass = frmMagGlass
        MagGlass.Visible = True
    End If
    Exit Sub
    
SBar_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.SBar_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
     Resume SBar_Click_Resume
End Sub

Private Sub Shrink()
Dim w As Long, h As Long
Dim tCursorPos As POINTAPI

    w = (picImage.Width * 0.9): h = (picImage.Height * 0.9)
    If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then    'verkleinern ohne GDI+
        picPaste.Width = w: picPaste.Height = h
        picPaste.PaintPicture picImage.Image, 0, 0, w, h, 0, 0, picImage.Width, picImage.Height
        Set picImage.Picture = picPaste.Image
        Set picPaste.Picture = Nothing
        picImage.Width = w: picImage.Height = h
    Else                                    'verkleinern mit GDI+
        picImage.Picture = picImage.Image
        w = w \ LTwipsPerPixelX: h = h \ LTwipsPerPixelY
        gdiplus.ResizePicture picImage, w, h
        picImage.Width = w * LTwipsPerPixelX: picImage.Height = h * LTwipsPerPixelY
        picImage.Picture = picImage.Image
    End If
    PaintGrading
    
    mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
    TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False

    If Not MagGlass Is Nothing Then
        DoEvents
        GetCursorPos tCursorPos
        MagGlass.PrintMagGlass tCursorPos
    End If
End Sub

Private Sub SyncFontAndColor()

    With TBar
        picImage.FontName = .FontName
        picImage.FontBold = .FontBold
        picImage.FontItalic = .FontItalic
        picImage.FontUnderline = .FontUnderline
        picImage.FontStrikethru = .FontStrikethru
        picImage.FontSize = .FontSize
        If .Selected = tbText Then picImage.ForeColor = TBar.FontColor Else picImage.ForeColor = SBar.ForeColor
        
        txtEditBox.FontName = .FontName
        txtEditBox.FontBold = .FontBold
        txtEditBox.FontItalic = .FontItalic
        txtEditBox.FontUnderline = .FontUnderline
        txtEditBox.FontStrikethru = .FontStrikethru
        txtEditBox.FontSize = .FontSize
        txtEditBox.ForeColor = .FontColor
        txtEditBox.Height = picImage.TextHeight("µÎ") + 1
        mTextOverhang = picImage.TextWidth("W")
    End With
    
    If SBar.Fill = 0 Then   'bei transparenten Hintergrund Vordergrundfarbe analysieren
        If modMain.IsLightColor(TBar.FontColor, 500) Then
            TBar.FontBackground = &HE0E0E0
            txtEditBox.BackColor = &HE0E0E0
        Else
            TBar.FontBackground = vbWhite
            txtEditBox.BackColor = vbWhite
        End If
    Else
        TBar.FontBackground = SBar.BackColor
        txtEditBox.BackColor = SBar.BackColor
    End If
    If txtEditBox.Visible Then txtEditBox.SetFocus
    
    
End Sub


Private Sub TBar_Change(ByVal newValue As tbButtons, ByVal OldValue As tbButtons, cancel As Boolean)
    On Error GoTo TBar_Change_Error
    With mWorkControl
        Select Case OldValue
            Case tbPaste:  If picPaste.Visible Then FixPaste
            Case tbText:   If txtEditBox.Visible Then DrawText
            Case tbDimension
                '###_START_PRO_###
                If Abs(.DrawMode) = tbDimension Then  'Abbruch Bemaßung
                    linDimension(0).Visible = False: linDimension(1).Visible = False: linDimension(2).Visible = False: shpDimension.Visible = False
                End If
                '###_END_PRO_###
        End Select
        .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
    End With
    Select Case True
        Case newValue = tbPaste
            If Clipboard.GetFormat(vbCFBitmap) Then
                Set picPaste.Picture = Clipboard.GetData(vbCFBitmap)
            ElseIf Clipboard.GetFormat(vbCFDIB) Then
                Set picPaste.Picture = Clipboard.GetData(vbCFDIB)
            Else
                MsgBox "Keine gültigen Bilddaten in der Zwischenablage gefunden.", vbInformation, "Einfügen"
                cancel = True
                Exit Sub
            End If
            TBar.Enabled(tbUndo) = True
            TBar.Enabled(tbRedo) = False
            With picPaste
                .MousePointer = vbSizeAll
                .Move 0, TBar.Height
                .Visible = True
                .SetFocus
            End With
        Case newValue = tbLegend Or TBar.SelectedEx = tbLegend
            SBar.Legend = True
        Case newValue = tbTear
            cancel = False
    End Select
    SBar.Legend = newValue = tbLegend Or TBar.SelectedEx = tbLegend
    If Not cancel Then Call ResetCursor(newValue)
    Exit Sub
    
TBar_Change_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.TBar_Change." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub


Private Sub TBar_Click(ByVal value As tbButtons, ByVal X As Long)
Dim w As Long, h As Long
Dim isRetry As Boolean
Dim p As StdPicture
Dim stackLgndChar As String

    On Error GoTo TBar_Click_Error
    Select Case value
        Case tbMenu
            If TBar.Selected = tbPaste And picPaste.Visible Then FixPaste
        Case tbCopy
            If TBar.Selected = tbPaste And picPaste.Visible Then FixPaste
Retry_Copy:
            On Error Resume Next
            Clipboard.Clear
            Clipboard.SetData picImage.Image, vbCFDIB
            If Err Then
                Err.Clear
                On Error GoTo TBar_Click_Error
                Sleep 500
                Clipboard.Clear
                Clipboard.SetData picImage.Image, vbCFDIB
            End If
            If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then Unload Me
        Case tbMagGlass
            frmMenu.ToogleMagGlass
        Case tbLineal
            If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then
                mGradingVisible = Not mGradingVisible
                Call PaintGrading
                SaveSetting App.Title, "Editor", "Grading", Abs(mGradingVisible)
            Else
                frmRuler.Visible = Not frmRuler.Visible
            End If
        Case tbNew
            Set Capture = New frmCapture
            If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then
                If TBar.Selected = tbPaste And picPaste.Visible Then FixPaste
                Capture.Show vbModeless, Me
            Else
                Capture.Move Me.Left, Me.Top, picImage.Width, picImage.Height
                Capture.Show
                Unload Me
            End If
        Case tbBorderStyle
            If TBar.Selected = tbPaste And picPaste.Visible Then FixPaste
        Case tbScale
            If TBar.Selected = tbPaste And picPaste.Visible Then FixPaste
            Call Shrink
        Case tbExtend
            Call Extend
        Case tbFont
            Call TextStyle
        Case tbUndo
            If TBar.Selected = tbPaste And picPaste.Visible Then
                picPaste.Visible = False
                Set picPaste.Picture = Nothing
                TBar.Enabled(tbUndo) = mUndoStack.CanUndo
                TBar.Enabled(tbRedo) = mUndoStack.CanRedo
                TBar.Selected = tbPointer
                Exit Sub
            End If
            If txtEditBox.Visible Then
                If SendMessage(txtEditBox.hwnd, EM_CANUNDO, 0&, 0&) <> 0& Then
                    Call SendMessage(txtEditBox.hwnd, EM_UNDO, 0&, 0&)
                    Call SendMessage(txtEditBox.hwnd, EM_EMPTYUNDOBUFFER, 0&, 0&)
                    TBar.Enabled(tbUndo) = SendMessage(txtEditBox.hwnd, EM_CANUNDO, 0&, 0&) <> 0& Or mUndoStack.CanUndo
                    Exit Sub
                End If
            End If
            If mUndoStack.CanUndo Then
                If mUndoStack.GetUndo(p, stackLgndChar) Then
                    w = ScaleX(p.Width, vbHimetric, vbTwips)
                    h = ScaleY(p.Height, vbHimetric, vbTwips)
                    If picImage.Width <> w Or h <> picImage.Height Then
                        picImage.Move 0, TBar.Height, w, h
                        Call PaintGrading
                    End If
                    Set picImage.Picture = p
                    TBar.Enabled(tbRedo) = True
                End If
                TBar.Enabled(tbUndo) = mUndoStack.CanUndo
                If Len(stackLgndChar) Then SBar.LegendText = stackLgndChar
            End If
        Case tbRedo
            If mUndoStack.CanRedo Then
                If mUndoStack.GetRedo(p, stackLgndChar) Then
                    w = ScaleX(p.Width, vbHimetric, vbTwips)
                    h = ScaleY(p.Height, vbHimetric, vbTwips)
                    If picImage.Width <> w Or h <> picImage.Height Then
                        picImage.Move 0, TBar.Height, w, h
                        Call PaintGrading
                    End If
                    Set picImage.Picture = p
                    TBar.Enabled(tbUndo) = True
                    TBar.Enabled(tbRedo) = mUndoStack.CanRedo
                    If Len(stackLgndChar) Then SBar.LegendText = stackLgndChar
                End If
            End If
        Case tbArrow
            If TBar.Selected = tbPaste And picPaste.Visible Then FixPaste
            PopupMenu frmMenu.MArrow, vbPopupMenuLeftAlign, X * LTwipsPerPixelX, TBar.Height

        Case -1 'Test-Grafik einfügen
            If App.LogMode = 0 Then
                If TBar.Selected = tbPaste And picPaste.Visible Then
                    picPaste.Visible = False
                    Set picPaste.Picture = Nothing
                End If
                If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then
                    CreateTestImage 0, True
                ElseIf CBool(GetAsyncKeyState(vbKeyControl) And KEY_PRESSED) Then
                    CreateTestImage 1, False
                End If
            End If
    End Select

Exit Sub
 
TBar_Click_Error:
 Screen.MousePointer = vbDefault
 If Err = 521 And value = tbCopy And Not isRetry Then
    If MsgBox("Fehler: " & Err.Number & vbCrLf & Err.Description, vbInformation Or vbRetryCancel) = vbRetry Then
        isRetry = True
        Resume Retry_Copy
    End If
Else
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.TBar_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End If
End Sub


Private Sub TBar_MenuClick(Name As String, Caption As String, Index As Integer, Checked As Boolean)
Dim Filter As String, InitialDir As String, Extension As String, FileName As String
Dim Flags As Long
Dim i As Integer
Dim isPaste As Boolean, isSave As Boolean
Dim p As StdPicture
Dim faktorSW As Single, faktorSH As Single

    On Error GoTo TBar_MenuClick_Error
    If Name = "mnuFileOpen" Then
        InitialDir = GetSetting(App.Title, "Editor", "FileDir", "C:\")
        Extension = GetSetting(App.Title, "Editor", "Extension", "*.png")
    ElseIf Name = "mnuFilePaste" Then
        isPaste = True
        InitialDir = GetSetting(App.Title, "Editor", "PasteDir", "C:\")
        Extension = GetSetting(App.Title, "Editor", "Extension", "*.png")
    ElseIf Name = "mnuFileSave" Then
        isSave = True
        Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_EXTENSIONDIFFERENT
        InitialDir = GetSetting(App.Title, "Editor", "FileDir", "C:\")
        Extension = GetSetting(App.Title, "Editor", "Extension", "*.png")
        Filter = "JPG (*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & _
                 "BMP (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0) & _
                 "GIF (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & _
                 "PNG (*.png)" & Chr$(0) & "*.png" & Chr$(0) & Chr$(0)
        FileName = modMain.ShowSaveFileDlg(Filter, Flags, Me.hwnd, FileName, InitialDir, Extension)
    ElseIf Name = "mruFile" Then
        FileName = Caption
    ElseIf Name = "mruPaste" Then
        isPaste = True
        FileName = Caption
    ElseIf Name = "mnuBorder" Then
        MakeBorder Index
    ElseIf Name = "mnuReset" Then
        frmReset.Show vbModal, Me
    ElseIf Name = "mnuPrint" Then
        If Not Printer Is Nothing Then
            Set frmPrint.Picture = Nothing
            faktorSW = Printer.ScaleWidth / picImage.Width
            faktorSH = Printer.ScaleHeight / picImage.Height
            If faktorSW < 1 Or faktorSH < 1 Then
                If faktorSH < faktorSW Then faktorSW = faktorSH
                frmPrint.Width = CLng(picImage.Width * faktorSW)
                frmPrint.Height = CLng(picImage.Height * faktorSW)
                gdiplus.ResizePicture picImage, CLng((picImage.Width * faktorSW) \ LTwipsPerPixelX), CLng((picImage.Height * faktorSW) \ LTwipsPerPixelX), frmPrint.hDC
                frmPrint.Picture = frmPrint.Image
                
            Else
                frmPrint.Width = picImage.Width
                frmPrint.Height = picImage.Height
                frmPrint.Picture = picImage.Image
            End If
        Else
            MsgBox "Der Standard-Drucker ist nicht verfügbar!", vbCritical, "Drucken..."
            Exit Sub
        End If
        frmPrint.PrintForm
        DoEvents
        Unload frmPrint
    End If
    
    '====File-Aktionen===
    If Name = "mnuFileOpen" Or Name = "mnuFilePaste" Then
        Flags = modMain.OFN_FILEMUSTEXIST Or modMain.OFN_PATHMUSTEXIST Or modMain.OFN_EXPLORER Or OFN_EXTENSIONDIFFERENT
        Filter = "Alle Bilddateien" & Chr$(0) & "*.jpg;*.jpe;*.jfif;*.gif;*.png;*.jpeg;*.bmp;*.tif;*.tiff;*.ico" & Chr$(0) & _
                "JPG (*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & _
                "JPE (*.jpe)" & Chr$(0) & "*.jpe" & Chr$(0) & _
                "JFIF (*.jfif)" & Chr$(0) & "*.jfif" & Chr$(0) & _
                "GIF (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & _
                "PNG (*.png)" & Chr$(0) & "*.png" & Chr$(0) & _
                "JPEG (*.jpeg)" & Chr$(0) & "*.jepg" & Chr$(0) & _
                "BMP (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0) & _
                "TIF (*.tif)" & Chr$(0) & "*.tif" & Chr$(0) & _
                "TIFF (*.tiff)" & Chr$(0) & "*.tiff" & Chr$(0) & _
                "ICON (*.ico)" & Chr$(0) & "*.ico" & Chr$(0) & Chr$(0)
        FileName = Trim$(modMain.ShowOpenFileDlg(Filter, Flags, Me.hwnd, "Bild laden...", InitialDir))
    End If
    If Len(FileName) Then
        i = InStrRev(FileName, "\")
        If i > 0 Then
            InitialDir = Left$(FileName, i)
            Extension = LCase$(GetFileExtension(FileName))
            If isPaste Then
                SaveSetting App.Title, "Editor", "PasteDir", InitialDir
                SaveSetting App.Title, "Editor", "Extension", Extension
            Else
                SaveSetting App.Title, "Editor", "FileDir", InitialDir
                SaveSetting App.Title, "Editor", "Extension", Extension
            End If
        End If

        If isPaste Then
            Set p = gdiplus.OpenPicture(FileName)
            picPaste.Move 0, TBar.Height, ScaleX(p.Width, vbHimetric, vbTwips), ScaleX(p.Height, vbHimetric, vbTwips)
            Set picPaste.Picture = p
            TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
            With picPaste
                .MousePointer = vbSizeAll
                .Visible = True
                .SetFocus
            End With
            TBar.Selected = tbPaste
            frmMenu.UpdatePasteMru FileName
        ElseIf isSave Then
            picImage.Picture = picImage.Image
            If gdiplus.SavePicture(picImage.Image, FileName) Then
                mCurrentFileName = FileName
                Me.Caption = "Pixel-Lineal - " & mCurrentFileName
                frmMenu.UpdateFileMru FileName
            End If
        Else
            Set p = gdiplus.OpenPicture(FileName)
            If Not p Is Nothing Then
                picImage.Move 0, TBar.Height, ScaleX(p.Width, vbHimetric, vbTwips), ScaleX(p.Height, vbHimetric, vbTwips)
                Set picImage = p
            End If
            mCurrentFileName = FileName
            Me.Caption = "Pixel-Lineal - " & mCurrentFileName
            mUndoStack.CreateUndoStep gdiplus.CopyStdPicture(picImage.Image)
            TBar.Enabled(tbUndo) = True: TBar.Enabled(tbRedo) = False
            frmMenu.UpdateFileMru FileName
            Call PaintGrading
        End If
    End If
    Exit Sub

TBar_MenuClick_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.TBar_MenuClick." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub TearHorizontal(y1 As Long, y2 As Long)
Dim tw As Long, th As Long, w As Long, h As Long
Dim i As Integer
    'hier Twips
    tw = frmMenu.picTearOff(0).Width * LTwipsPerPixelX: th = frmMenu.picTearOff(0).Height * LTwipsPerPixelY
    w = picImage.Width
    h = picImage.Height
    y1 = y1 * LTwipsPerPixelY + th
    y2 = y2 * LTwipsPerPixelY - th
    
    picPaste.Width = w
    picPaste.Height = y1 + (h - y2) + th
    picPaste.PaintPicture picImage.Image, 0, 0, w, y1, 0, 0, w, y1
    picPaste.PaintPicture picImage.Image, 0, y1 + th, w, h - y2, 0, y2, w, h - y2
    
    picImage.Height = picPaste.Height
    picImage.Picture = picPaste.Image
    Set picPaste.Picture = Nothing
    'ab hier Pixel
    For i = 0 To w + tw Step tw
        TransparentBlt hDC:=picImage.hDC, X:=i \ LTwipsPerPixelX, Y:=(y1 - th) \ LTwipsPerPixelY, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                       hSrcDC:=frmMenu.picTearOff(0).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
    Next
    For i = 0 To w + tw Step tw
        TransparentBlt hDC:=picImage.hDC, X:=i \ LTwipsPerPixelX, Y:=(y1 + th) \ LTwipsPerPixelX, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                       hSrcDC:=frmMenu.picTearOff(2).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
    Next
    Call PaintGrading
    
End Sub

Private Sub TearVertical(x1 As Long, x2 As Long)
Dim tw As Long, th As Long, w As Long, h As Long
Dim i As Integer
    'hier Twips
    tw = frmMenu.picTearOff(1).Width * LTwipsPerPixelX: th = frmMenu.picTearOff(1).Height * LTwipsPerPixelY
    w = picImage.Width
    h = picImage.Height
    x1 = x1 * LTwipsPerPixelX + tw
    x2 = x2 * LTwipsPerPixelX - tw
    
    picPaste.Width = x1 + (w - x2) + tw
    picPaste.Height = h
    picPaste.PaintPicture picImage.Image, 0, 0, x1, h, 0, 0, x1, h
    picPaste.PaintPicture picImage.Image, x1 + tw, 0, w - x2, h, x2, 0, w - x2, h
    
    picImage.Width = picPaste.Width
    picImage.Picture = picPaste.Image
    Set picPaste.Picture = Nothing
    'ab hier Pixel
    For i = 0 To h + th Step th
        TransparentBlt hDC:=picImage.hDC, X:=(x1 - tw) \ LTwipsPerPixelX, Y:=i \ LTwipsPerPixelY, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                       hSrcDC:=frmMenu.picTearOff(1).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)

    Next
    For i = 0 To h + th Step th
        TransparentBlt hDC:=picImage.hDC, X:=(x1 + tw) \ LTwipsPerPixelX, Y:=i \ LTwipsPerPixelY, nWidth:=tw \ LTwipsPerPixelX, nHeight:=th \ LTwipsPerPixelY, _
                       hSrcDC:=frmMenu.picTearOff(3).hDC, xSrc:=0, ySrc:=0, nSrcWidth:=tw \ LTwipsPerPixelX, nSrcHeight:=th \ LTwipsPerPixelY, crTransparent:=frmMenu.picTearOff(0).Point(0, 0)
    Next
    Call PaintGrading
End Sub


'=====picImage========================================
Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo picImage_MouseDown_Error
    If Button = vbLeftButton Then
        Select Case TBar.Selected
            Case tbFreehand:  DrawFreehand X, Y
            Case tbLine:      DrawLine X, Y
            Case tbRectangle: DrawRectangle X, Y
            Case tbCyrcle:    DrawCyrcle X, Y
            Case tbMarker:    DrawMarker X, Y
            Case tbObfus:     DrawObfus X, Y
            Case tbFill:      DrawFill X, Y
            Case tbText:      If txtEditBox.Visible Then DrawText Else DrawText X, Y, ActionStart
            Case tbArrow:     DrawArrow X, Y
            Case tbLegend:    DrawLegend X, Y
            Case tbDimension: DrawDimension X, Y
            Case tbCrop, tbTear
                With picImage
                     mDrawStyle.DrawStyle = .DrawStyle
                     mDrawStyle.DrawMode = .DrawMode
                     mDrawStyle.DrawWidth = .DrawWidth
                     mDrawStyle.FillStyle = .FillStyle
                     If modMain.IsLightColor(SBar.ForeColor) Then .ForeColor = &HEEEEEE
                    .DrawStyle = vbDash
                    .DrawMode = vbNotXorPen
                    .ForeColor = vbBlack
                    .DrawWidth = 1
                End With
                With mWorkControl
                    If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then
                        If X <= 10 Then X = 0
                        If Y <= 10 Then Y = 0
                    End If
                    .x2 = .x0: .y2 = .y0
                     picImage.Line (.x1, .y1)-(.x2, .y2), , B
                    .x1 = X: .y1 = Y
                    .x0 = .x1: .y0 = .y1
                    .DrawMode = tbCrop
                End With
            Case tbLegend
                SBar.Legend = True
            Case tbPaste
                FixPaste
        End Select
    End If

Exit Sub

picImage_MouseDown_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmImage.picImage_MouseDown." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With mWorkControl
        '###_START_PRO_###
        If .DrawMode = tbDimension Then  'Bemaßung1
            If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then   'vertikale Bemaßung
                linDimension(0).x2 = X: linDimension(0).y2 = .y0
                linDimension(1).x1 = X: linDimension(1).y1 = .y0:  linDimension(1).x2 = X: linDimension(1).y2 = Y
            Else                                                'horizontale Bemaßung
                linDimension(0).x2 = .x0: linDimension(0).y2 = Y
                linDimension(1).x1 = .x0: linDimension(1).y1 = Y:  linDimension(1).x2 = X: linDimension(1).y2 = Y
            End If
            SBar.Coordinates X + 1, Y + 1
            Exit Sub
        ElseIf Abs(.DrawMode) = tbDimension Then  'Bemaßung2
            If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then   'vertikale Bemaßung
                linDimension(0).x2 = X: linDimension(0).y2 = .y0
                Select Case True
                    Case Y < .y0 And .y0 < .y1: linDimension(1).y1 = .y1: linDimension(1).y2 = Y - .x2 - 10
                    Case Y > .y1 And .y0 < .y1: linDimension(1).y1 = Y:   linDimension(1).y2 = .y0 - 1
                    Case Y < .y1 And .y1 < .y0: linDimension(1).y1 = .y0: linDimension(1).y2 = Y - .x2 - 10
                    Case Y > .y0 And .y1 < .y0: linDimension(1).y1 = .y1: linDimension(1).y2 = Y
                    Case .y1 < .y0:             linDimension(1).y1 = .y0: linDimension(1).y2 = .y1 - 1
                    Case .y1 > .y0:             linDimension(1).y1 = .y0: linDimension(1).y2 = .y1 + 1
                End Select
                linDimension(1).x1 = X: linDimension(1).x2 = X
                linDimension(2).x2 = X: linDimension(2).y2 = .y1
                If .y2 = 0 Then
                    .y2 = shpDimension.Width
                    shpDimension.Width = shpDimension.Height
                    shpDimension.Height = .y2
                    .y2 = 1
                End If
                shpDimension.Move X - .x2 - 2, Y - .x2
            Else                                                'horizontale Bemaßung
                linDimension(0).x2 = .x0: linDimension(0).y2 = Y
                Select Case True
                    Case X < .x0 And .x1 > .x0:   linDimension(1).x1 = X:   linDimension(1).x2 = .x1 + 1
                    Case X > .x1 And .x1 > .x0:   linDimension(1).x1 = .x0: linDimension(1).x2 = X + .x2 + 10
                    Case X < .x1 And .x1 < .x0:   linDimension(1).x1 = X:   linDimension(1).x2 = .x0 + 1
                    Case X > .x0 And .x1 < .x0:   linDimension(1).x1 = .x1: linDimension(1).x2 = X + .x2 + 10
                    Case .x1 > .x0:               linDimension(1).x1 = .x0: linDimension(1).x2 = .x1 + 1
                    Case .x1 < .x0:               linDimension(1).x1 = .x0: linDimension(1).x2 = .x1 - 1
                End Select
                linDimension(1).y1 = Y:   linDimension(1).y2 = Y
                linDimension(2).x2 = .x1: linDimension(2).y2 = Y
                If .y2 = 1 Then
                    .y2 = shpDimension.Width
                    shpDimension.Width = shpDimension.Height
                    shpDimension.Height = .y2
                    .y2 = 0
                End If
                shpDimension.Move X, Y - .x2 - 2
            End If
            SBar.Coordinates X + 1, Y + 1
            Exit Sub
        End If
        '###_END_PRO_###
        
        If Button = vbLeftButton Then
            If .DrawMode = tbFreehand Or .DrawMode = tbMarker Then 'Punkt oder Marker
                If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) And mWorkControl.y0 <> 0 Then Y = mWorkControl.y0
                picImage.PSet (X, Y), SBar.ForeColor
                SBar.Coordinates X + 1, Y + 1
            ElseIf .DrawMode = tbLine Then  'Linie
                If .x1 <> X Or .y1 <> Y Then
                    .x2 = .x0: .y2 = .y0
                    picImage.Line (.x1, .y1)-(.x2, .y2)
                    If CBool(GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) Then
                        If Abs(.x1 - X) > Abs(.y1 - Y) Then
                            .x2 = X: .y2 = .y1
                        Else
                            .x2 = .x1: .y2 = Y
                        End If
                    Else
                        .x2 = X: .y2 = Y
                    End If
                    picImage.Line (.x1, .y1)-(.x2, .y2)
                End If
                .x0 = .x2: .y0 = .y2
                SBar.Coordinates X + 1, Y + 1
            ElseIf .DrawMode = tbRectangle Or .DrawMode = tbObfus Or .DrawMode = tbCrop Or .DrawMode = tbTear Then   'Rechteck oder Ausschneiden
                If .x1 <> X Or .y1 <> Y Then
                    .x2 = .x0: .y2 = .y0
                    picImage.Line (.x1, .y1)-(.x2, .y2), , B
                    .x2 = X: .y2 = Y
                    picImage.Line (.x1, .y1)-(.x2, .y2), , B
                    .x0 = .x2: .y0 = .y2
                End If
                SBar.Coordinates X + 1, Y + 1, Abs(.x2 - .x1), Abs(.y2 - .y1)
            ElseIf .DrawMode = tbCyrcle Then  'Kreis
                If .x1 <> X Or .y1 <> Y Then
                    .x2 = .x0: .y2 = .y0
                    If Abs(.x2 - .x1) > (SBar.Line + 1) Then picImage.Circle (.x1, .y1), Abs(.x2 - .x1)
                    .x2 = X: .y2 = Y
                    If Abs(.x2 - .x1) > (SBar.Line + 1) Then picImage.Circle (.x1, .y1), Abs(.x2 - .x1)
                    .x0 = .x2: .y0 = .y2
                End If
                SBar.Coordinates X + 1, Y + 1
            End If
        Else
            SBar.Coordinates X + 1, Y + 1
        End If
    End With
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With mWorkControl
        If Button = vbLeftButton Then
            On Error GoTo picImage_MouseUp_Error
            Select Case Abs(mWorkControl.DrawMode)
                Case tbFreehand:   DrawFreehand X, Y, ActionEnd
                Case tbLine:       DrawLine X, Y, ActionEnd
                Case tbRectangle:  DrawRectangle X, Y, ActionEnd
                Case tbCyrcle:     DrawCyrcle X, Y, ActionEnd
                Case tbMarker:     DrawMarker X, Y, ActionEnd
                Case tbObfus:      DrawObfus X, Y, ActionEnd
                Case tbDimension:  Exit Sub
                Case tbCrop, tbTear
                    With mWorkControl
                        If .DrawMode = 0 Then Exit Sub
                        .x2 = .x0: .y2 = .y0
                    End With
                    Call CropOrTearImage
            End Select
            .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
            TBar.Enabled(tbUndo) = mUndoStack.CanUndo
            TBar.Enabled(tbRedo) = mUndoStack.CanRedo
        '###_START_PRO_###
        ElseIf Button = vbRightButton And Abs(.DrawMode) = tbDimension Then     'Abbruch Bemaßung
            .x0 = 0: .y0 = 0: .x1 = 0: .y1 = 0: .x2 = 0: .y2 = 0: .DrawMode = 0
            linDimension(0).Visible = False: linDimension(1).Visible = False: linDimension(2).Visible = False: shpDimension.Visible = False
        '###_END_PRO_###
        End If
    End With
Exit Sub

picImage_MouseUp_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmImage.picImage_MouseUp." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub picPaste_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tCursorPos As POINTAPI
    On Error GoTo Form_KeyDown_Error
    With picPaste
        Select Case KeyCode
            Case vbKeyLeft:     .Left = .Left - LTwipsPerPixelX
            Case vbKeyRight:    .Left = .Left + LTwipsPerPixelX
            Case vbKeyUp:       .Top = .Top - LTwipsPerPixelY
            Case vbKeyDown:     .Top = .Top + LTwipsPerPixelY
            Case vbKeyEscape, vbKeyDelete, vbKeyBack
                .Visible = False
                 Set .Picture = Nothing
                .Width = 1
                .Height = 1
                TBar.Selected = tbPointer
            Case vbKeyReturn, vbKeySpace
                TBar_Click tbPointer, True
        End Select
    End With
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
 "Quelle: frmImage.Form_KeyDown." & Erl & vbCrLf & Err.Source, _
 vbCritical
    
End Sub


Private Sub picPaste_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then
        If TBar.Selected = tbPaste Then
            With mWorkControl
                .DrawMode = tbPaste
                .x0 = X
                .y0 = Y
            End With
        End If
    End If
End Sub

Private Sub picPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xD&, yD&
    If Button = vbLeftButton Or Button = vbMiddleButton Then
        If TBar.Selected = tbPaste Then
            xD = picPaste.Left + (X - mWorkControl.x0)
            yD = picPaste.Top + (Y - mWorkControl.y0)
            picPaste.Move xD, yD
            SBar.Coordinates xD \ LTwipsPerPixelX, yD \ LTwipsPerPixelX
        End If
    End If
End Sub

Private Sub picPaste_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then
        If mWorkControl.DrawMode = tbPaste Then mWorkControl.DrawMode = 0
    End If
End Sub

Private Sub txtEditBox_Change()
Dim txWidth As Long
    On Error GoTo txtEditBox_Change_Error
    With txtEditBox
        If Not TBar.Enabled(tbUndo) Then
            TBar.Enabled(tbUndo) = Not (SendMessage(.hwnd, EM_CANUNDO, 0&, 0&) = 0&)
        End If
        txWidth = picImage.TextWidth(.Text) + mTextOverhang
        .Width = txWidth
    End With
    Exit Sub
    
txtEditBox_Change_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.txtEditBox_Change." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub txtEditBox_KeyDown(KeyCode As Integer, Shift As Integer)
Dim txWidth As Long
    On Error GoTo txtEditBox_KeyDown_Error
    Select Case KeyCode
        Case vbKeyEscape
            txtEditBox.Visible = False
            ResetCursor
        Case vbKeyReturn
            Call DrawText
        Case 32 To 255
            txWidth = picImage.TextWidth(txtEditBox.Text & Chr$(KeyCode)) + mTextOverhang
            If txtEditBox.Width < txWidth Then txtEditBox.Width = txWidth
    End Select
    Exit Sub
    
txtEditBox_KeyDown_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmImage.txtEditBox_KeyDown." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub txtEditBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbMiddleButton Then
        txtEditBox.MousePointer = vbSizeAll
        With mWorkControl
            .DrawMode = tbText
            .x0 = X \ LTwipsPerPixelX
            .y0 = Y \ LTwipsPerPixelY
        End With
    End If
End Sub

Private Sub txtEditBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xD&, yD&
    If Button = vbMiddleButton Then
        If mWorkControl.DrawMode = tbText Then
            xD = txtEditBox.Left + ((X \ LTwipsPerPixelX) - mWorkControl.x0)
            yD = txtEditBox.Top + ((Y \ LTwipsPerPixelY) - mWorkControl.y0)
            If xD < 0 Then xD = 0: If yD < 0 Then yD = 0
            txtEditBox.Move xD, yD
            SBar.Coordinates CSng(xD), CSng(yD)
        End If
    End If
End Sub

Private Sub txtEditBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mWorkControl.DrawMode = tbText Then
        mWorkControl.DrawMode = 0
        txtEditBox.MousePointer = vbDefault
    End If
End Sub



Private Sub AdjustingWorkControlEdges()
    With mWorkControl
        If .x2 > picImage.ScaleWidth - 1 Then .x2 = picImage.ScaleWidth - 1     'über rechten Rand verhindern
        If .y2 > picImage.ScaleHeight - 1 Then .y2 = picImage.ScaleHeight - 1   'über unteren Rand verhindern
        If .x2 < 0 Then .x2 = 0                                                 'über linken Rand verhindern
        If .y2 < 0 Then .y2 = 0                                                 'über oberen Rand verhindern
    End With
End Sub

Private Sub CutLine(r As Double)
Dim f As Double
    With mWorkControl
        f = Math.Sqr(Abs(.x2 - .x1) ^ 2 + Abs(.y2 - .y1) ^ 2)
        f = (f - (r / 2)) / f
        Select Case True
            Case .x2 >= .x1 And .y2 >= .y1  'SO
                .x2 = ((.x2 - .x1) * f) + .x1
                .y2 = ((.y2 - .y1) * f) + .y1
            Case .x2 < .x1 And .y2 >= .y1   'SW
                .x2 = .x1 - ((.x1 - .x2) * f)
                .y2 = ((.y2 - .y1) * f) + .y1
            Case .x2 >= .x1 And .y2 < .y1   'NO
                .x2 = ((.x2 - .x1) * f) + .x1
                .y2 = .y1 - ((.y1 - .y2) * f)
            Case .x2 < .x1 And .y2 < .y1    'NW
                .x2 = .x1 - ((.x1 - .x2) * f)
                .y2 = .y1 - ((.y1 - .y2) * f)
        End Select
    End With
End Sub
