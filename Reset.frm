VERSION 5.00
Begin VB.Form frmReset 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Einstellungen zurücksetzen..."
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Reset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chkResetAll 
      Appearance      =   0  '2D
      Caption         =   "&Alle Einstellungen zurücksetzen"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   390
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox lstReset 
      Appearance      =   0  '2D
      Height          =   2055
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Kontrollkästchen
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With lstReset
        .AddItem "Bildschirlupe-Vergrößerungsfaktor":       .ItemData(.NewIndex) = 0
        .AddItem "Bildschirlupe-Statusbar":                 .ItemData(.NewIndex) = 1
        .AddItem "Gesamelte Farben":                        .ItemData(.NewIndex) = 2
        .AddItem "Lineal-Farben":                           .ItemData(.NewIndex) = 3
        .AddItem "Lineal-Schrift":                          .ItemData(.NewIndex) = 4
        .AddItem "Lineal-Transparenz":                      .ItemData(.NewIndex) = 5
        .AddItem "Liste der zuletzt geöffneten Bilder":     .ItemData(.NewIndex) = 6
        .AddItem "Liste der zuletzt eingefügten Bilder":    .ItemData(.NewIndex) = 7
        .AddItem "Text-Einstellungen im Bildeditor":        .ItemData(.NewIndex) = 8
        .AddItem "Suche nach Online-Updates":               .ItemData(.NewIndex) = 9
    End With
End Sub
Private Sub ResetColorCollection()
Dim f As Form
Dim i As Integer
    ReDim ColorCollection(0)
    ColorCollection(0) = -1
    For Each f In Forms
        If TypeOf f Is frmMagGlass Then
            For i = 1 To f.mnuColorCollectionItems.UBound: f.mnuColorCollectionItems(i).Visible = False: Next i
        ElseIf TypeOf f Is frmMenu Then
            For i = 1 To f.mnuColorCollectionItems.UBound: f.mnuColorCollectionItems(i).Visible = False: Next i
            f.mnuPal(f.mnuPal.UBound).Visible = False
        ElseIf TypeOf f Is frmImage Then
        
        End If
    Next
    
End Sub

Private Sub ResetMagFaktorX()
    On Error Resume Next
    DeleteSetting App.Title, "Options", "XFaktor"
    If Not MagGlass Is Nothing Then MagGlass.SetFactorX 6
End Sub
Private Sub ResetMagStatus()
    On Error Resume Next
    DeleteSetting App.Title, "Options", "Status"
    If Not MagGlass Is Nothing Then MagGlass.StatusBarVisible = True
End Sub
Private Sub ResetMruFile()
Dim i As Integer
    On Error Resume Next
    For i = 0 To 9
        DeleteSetting App.Title, "Editor", "LastFile" & (i + 1)
        frmMenu.mruFile(i).Visible = False
        frmMenu.mruFile(i).Caption = ""
    Next i
    
End Sub
Private Sub ResetMruPaste()
Dim i As Integer
    On Error Resume Next
    For i = 0 To 9
        DeleteSetting App.Title, "Editor", "LastPaste" & (i + 1)
        frmMenu.mruPaste(i).Visible = False
        frmMenu.mruPaste(i).Caption = ""
    Next i
End Sub
Private Sub ResetRulerColors()
    On Error Resume Next
    DeleteSetting App.Title, "Options", "BackColor"
    DeleteSetting App.Title, "Options", "ForeColor"
    DeleteSetting App.Title, "Options", "MarkColor"
    With frmRuler
        .picRuler.BackColor = RGB(255, 255, 231)
        .picRuler.ForeColor = RGB(132, 132, 132)
        .MarkerColor = RGB(255, 0, 0)
        .ProcRefreshRuler .Left, .Top
    End With
End Sub
Private Sub ResetRulerFont()
    On Error Resume Next
    DeleteSetting App.Title, "Options", "FontBold"
    DeleteSetting App.Title, "Options", "FontItalic"
    DeleteSetting App.Title, "Options", "FontName"
    DeleteSetting App.Title, "Options", "FontSize"
    DeleteSetting App.Title, "Options", "FontStrikethru"
    DeleteSetting App.Title, "Options", "FontUnderline"
    
    With frmRuler
        With .picRuler.Font
            .Name = "Arial"
            .bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
            .Size = 6
        End With
        .ProcRefreshRuler .Left, .Top
    End With
End Sub
Private Sub ResetRulerTransparenz()
Dim i As Integer
    On Error Resume Next
    DeleteSetting App.Title, "Options", "TransparencyRuler"
    With frmMenu
        .mnuTransparencyRuler(0).Checked = True
        For i = 0 To .mnuTransparencyRuler.UBound
            .mnuTransparencyRuler(i).Checked = False
        Next i
    End With
    Call TransparencyRuler(frmRuler.hwnd, 255&)
End Sub
Private Sub ResetText()
Dim f As Form
    On Error Resume Next
    DeleteSetting App.Title, "Textbox"
    For Each f In Forms
        If TypeOf f Is frmImage Then
            f.TextStyle reset:=True
        End If
    Next f
    
End Sub
Private Sub ResetUpdates()
    On Error Resume Next
    DeleteSetting App.Title, "Options", "VerInfo"
End Sub

Private Sub cmdDialog_Click(Index As Integer)
Dim i As Integer
Dim f As Form
    If Index = 0 Then
        For i = 0 To lstReset.ListCount - 1
            If lstReset.Selected(i) Or CBool(chkResetAll.Value) Then
                Select Case lstReset.ItemData(i)
                    Case 0: ResetMagFaktorX
                    Case 1: ResetMagStatus
                    Case 2: ResetColorCollection
                    Case 3: ResetRulerColors
                    Case 4: ResetRulerFont
                    Case 5: ResetRulerTransparenz
                    Case 6: ResetMruFile
                    Case 7: ResetMruPaste
                    Case 8: ResetText
                    Case 9: ResetUpdates
                End Select
            End If
        Next i
    End If
    If CBool(chkResetAll.Value) Then
        On Error Resume Next
        For Each f In Forms
            If TypeOf f Is frmImage Then
                With f
                    .SBar.Line = 0
                    .SBar.ForeColor = vbBlack
                    .SBar.BackColor = vbWhite
                    .SBar.Palette = 0
                    .TBar.Selected = 0
                    .TBar.Arrow = 0
                    .SBar.Fill = 0
                End With
            End If
        Next f
        DeleteSetting App.Title, "Options"
        DeleteSetting App.Title, "ScreenShot"
        DeleteSetting App.Title, "Editor"
        DeleteSetting App.Title, "Textbox"
    End If
    Unload Me
Exit Sub

cmdDialog_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmReset.cmdDialog_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

