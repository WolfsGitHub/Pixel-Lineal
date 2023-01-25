Attribute VB_Name = "modMain"
Option Explicit

Public MagGlass As frmMagGlass
Public MagColor As frmMagColor
Public Capture As frmCapture
Public CloseApp As Boolean


Public ForceRefresh As Integer
Public Const FORCE_REFRESH_RES As Integer = 10  'nicht keiner machen, weil sonst die Anzeige in der Lupe verfälscht wird

Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)




Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112&
Public Const HTCAPTION = 2

Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_EMPTYUNDOBUFFER = &HCD


Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Const WS_BORDER = &H800000
Public Const GWL_STYLE = -16

Public Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_SHOWWINDOW = &H40

'Public Declare Function GetDesktopWindow Lib "user32" () As Long
'Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function StretchBlt Lib "gdi32" ( _
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
Public Const SRCCOPY = &HCC0020

Public Declare Function GetCursorPos Lib "user32" (ByRef lngPunkte As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum MousePos
    mpOther
    mpLeft
    mpTop
    mpRight
    mpBottom
    mpTopLeft
    mpBottomLeft
    mpTopRight
    mpBottomRight
End Enum

Public Const SC_SIZE_BottomRight As Long = &HF008&
Public Const SC_SIZE_BottomLeft As Long = &HF007&
Public Const SC_SIZE_Bottom As Long = &HF006&
Public Const SC_SIZE_TopRight As Long = &HF005&
Public Const SC_SIZE_TopLeft As Long = &HF004&
Public Const SC_SIZE_Top As Long = &HF003&
Public Const SC_SIZE_Right As Long = &HF002&
Public Const SC_SIZE_Left As Long = &HF001&

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
        
Private Declare Function ShellExecuteA Lib "shell32.dll" ( _
  ByVal hwnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long
        

'Farbauswahl
Private Declare Function ChooseColor Lib "comdlg32.dll" _
  Alias "ChooseColorA" (pChoosecolor As LPCHOOSECOLOR) As Long

Private Type LPCHOOSECOLOR
  lStructSize As Long
  hwnd As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4

Public Enum PL_Orientation
    PL_HORIZONTAL = 0
    PL_VERTICAL = 1
End Enum

'Einheit
Public Enum PL_ScaleMode
  PL_PIXEL
  PL_TWIPS
  PL_USER
End Enum
Public RulerScaleMode As PL_ScaleMode
Public RulerScaleMulti As Double


Public Enum PL_ColorCode
  PL_HEXHTML = 0
  PL_HEXVB = 1
  PL_OLE = 2
End Enum

Public Const plZeroLine = 1
Public Const plBREADTH = 22

'Für Menümanipulation
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" ( _
        ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long, ByVal wIDNewItem As Long, _
        ByVal lpString As Any) As Long

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_BYPOSITION = &H400&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public ColorCollection() As Long
Public ColorCode As PL_ColorCode

'####Transparenz####
Declare Function SetLayeredWindowAttributes Lib "user32.dll" ( _
                 ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long

Public Const GWL_EXSTYLE As Long = -20&
Public Const WS_EX_LAYERED As Long = &H80000
'Macht nur eine Farbe transparent
Public Const LWA_COLORKEY = &H1
'Macht das ganze Fenster transparent
Public Const LWA_ALPHA = &H2

'#### Ende Transparenz ###

Public Const XYFieldMinWidth = 80
Public XYFieldWidth As Long
Public LTwipsPerPixelX As Long
Public LTwipsPerPixelY As Long
Public LScreenWidth As Long
Public LScreenHeight As Long
Public VMarker() As Single
Public HMarker() As Single
Public TMarker As Single

'#### FILE ####
Public Declare Function GetOpenFileNameW Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" _
        Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Type OPENFILENAME
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXPLORER = &H80000
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_EXTENSIONDIFFERENT = &H400


'#### FONT ####
Public Declare Function ChooseFont Lib "comdlg32.dll" _
  Alias "ChooseFontA" ( _
  lpcf As CHOOSEFONT_TYPE) As Long
 
Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type
 
Public Type CHOOSEFONT_TYPE
    lStructSize As Long
    hWndOwner As Long ' caller's window handle
    hDC As Long ' printer DC/IC or NULL
    lpLogFont As Long ' ptr. to a LOGFONT struct
    iPointSize As Long ' 10 * size in points of selected font
    Flags As Long ' enum. private Type flags
    rgbColors As Long ' returned text color
    lCustData As Long ' data passed to hook fn.
    lpfnHook As Long ' ptr. to hook function
    lpTemplateName As String ' custom template name
    hInstance As Long ' instance handle of.EXE that
    lpszStyle As String ' return the style field here
    nFontType As Integer ' same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long ' minimum pt size allowed &
    nSizeMax As Long ' max pt size allowed if
End Type
 
'' Zeigt nur Windows- oder Unicode-Fonts
'Public Const CF_ANSIONLY = &H400
'' Zeigt einen "Übernehmen" Button an
'Public Const CF_APPLY = &H200
'' Listet Drucker- und Bildschirm-Fonts
'Public Const CF_BOTH = &H3
'' Erlaubt Font-Besonderheiten wie
'' Unterstreichen, Farbe und Durchgestrichen
Public Const CF_EFFECTS = &H100
'' Aktiviert die Callback-Funktion
'Public Const CF_ENABLEHOOK = &H8
'' Der Dialog benutzt Template's die von
'' TemplateNames festgelegt sind
'Public Const CF_ENABLETEMPLATE = &H10
'' Verwendet den durch hInstance festgelegten Dialog
'Public Const CF_ENABLETEMPLATEHANDLE = &H20
'' Listet nur Fixed-Pitch Fonts
'Public Const CF_FIXEDPITCHONLY = &H4000
'' Verweigert die Eingabe nicht aufgeführter Fonts
Public Const CF_FORCEFONTEXIST = &H10000
'' Setzt die Startwerte, welche über die
'' LOGFONT-Struktur angegeben wurden
Public Const CF_INITTOLOGFONTSTRUCT = &H40
'' Erlaubt nur Schriftgrößen im Bereich "nSizeMin" und "nSizeMax"
Public Const CF_LIMITSIZE = &H2000
'' Zeigt keine OEM Fonts
'Public Const CF_NOOEMFONTS = &H800
'' Kein Standard Facenamen selektieren
'Public Const CF_NOFACESEL = &H80000
'' Kein Standard Script selektieren
Public Const CF_NOSCRIPTSEL = &H800000
'' keine Standardgröße setzen
'Public Const CF_NOSIZESEL = &H200000
'' Kein Beispiel (Vorschau) anzeigen
'Public Const CF_NOSIMULATIONS = &H1000
'' kein Standard-Stil setzen
'Public Const CF_NOSTYLESEL = &H100000
'' keine Vector-Fonts anzeigen
'Public Const CF_NOVECTORFONTS = &H800
'' keine vertikal ausgerichtete Fonts anzeigen
'Public Const CF_NOVERTFONTS = &H1000000
'' Listet Drucker-Fonts
'Public Const CF_PRINTERFONTS = &H2
'' Listet nur skalierbare Fonts
'Public Const CF_SCALABLEONLY = &H20000
'' Listet Bildschirm-Fonts
Public Const CF_SCREENFONTS = &H1
'' Listet nur Windows- oder Unicode-Fonts
'Public Const CF_SCRIPTSONLY = &H400
'' Listet nur Script-Fonts
'Public Const CF_SELECTSCRIPT = &H400000
'' Zeigt den Hilfe-Button an
'Public Const CF_SHOWHELP = &H4
'' Listet nur TrueType-Schriftarten
'Public Const CF_TTONLY = &H40000
'' Verwendet die in "lpStyle" angegebenen Werte
'Public Const CF_USESTYLE = &H80
'' Listet nur Fonts, die Drucker- und Bildschirm-Fonts gleichzeitig sind
'' (muss benutzt werden mit CF_BOTH und CF_SCALABLEONLY)
'Public Const CF_WYSIWYG = &H8000
'
Public Const SCREEN_FONTTYPE = &H2000 ' Bildschirm-Fonts


' lfWeight Konstanten
' ===================
Public Const LF_FACESIZE = 32
Public Const ANTIALIASED_QUALITY = 4
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const DEFAULT_CHARSET = 1
Public Const OUT_TT_PRECIS = 4
Public Const VARIABLE_PITCH = 2
'Public Const FW_DONTCARE = 0       ' Standard
'Public Const FW_THIN = 100         ' super dünn
'Public Const FW_EXTRALIGHT = 200   ' extra dünn
'Public Const FW_LIGHT = 300        ' dünn
'Public Const FW_NORMAL = 400       ' normal
'Public Const FW_MEDIUM = 500       ' mittel
'Public Const FW_SEMIBOLD = 600     ' etwas dicker
'Public Const FW_EXTRABOLD = 800    ' extra fett
'Public Const FW_HEAVY = 900        ' super fett

'Flächen befüllen
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal _
      hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal _
      crColor As Long, ByVal wFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1

Public Sub AddColor(ColorItem As Long)
DoEvents
Dim i As Integer, u As Integer
    If UBound(ColorCollection) < 15 Then
        If ColorCollection(0) > -1 Then
            ReDim Preserve ColorCollection(UBound(ColorCollection) + 1)
            If Not MagGlass Is Nothing Then MagGlass.mnuColorCollectionItems(UBound(ColorCollection)).Visible = True
            frmMenu.mnuColorCollectionItems(UBound(ColorCollection)).Visible = True
        End If
    End If
    
    For i = UBound(ColorCollection) To 1 Step -1
        ColorCollection(i) = ColorCollection(i - 1)
    Next i
    ColorCollection(0) = ColorItem
    
    With frmMenu
        u = .mnuPal.UBound
        Set .picMenuPal(u).Picture = Nothing
        For i = 0 To UBound(ColorCollection)
            .picMenuPal(u).Line (i * 10, 0)-((i + 1) * 10, 18), ColorCollection(i), BF
        Next i
        .picMenuPal(u).Picture = .picMenuPal(u).Image
        .mnuPal(u).Visible = True
        SetMenuItemBitmaps GetSubMenu(GetMenu(.hwnd), 5&), u, MF_BYPOSITION, .picMenuPal(u).Picture, .picMenuPal(u).Picture
    End With
    
End Sub

Public Function GetPxColor() As Long
Dim lDeskDC As Long
Dim tCursorPos As POINTAPI
    
    lDeskDC = GetDC(0&)
    GetCursorPos tCursorPos
    GetPxColor = GetPixel(lDeskDC, tCursorPos.X, tCursorPos.Y)
    ReleaseDC 0&, lDeskDC
    Exit Function
    
GetPxColor_Error:
    If lDeskDC Then ReleaseDC 0&, lDeskDC
End Function

Public Sub CopyRGB(lPxColor As Long, Optional cpy As Boolean = True)
Dim tCursorPos As POINTAPI
Dim i As Integer
    If cpy Then
        Clipboard.Clear
        If ColorCode = PL_HEXHTML Then
            Clipboard.SetText RGBtoHTML(lPxColor), vbCFText
        ElseIf ColorCode = PL_HEXVB Then
            Clipboard.SetText RGBtoVB(lPxColor), vbCFText
        Else
            Clipboard.SetText lPxColor, vbCFText
        End If
        If Not MagGlass Is Nothing Then
            GetCursorPos tCursorPos
            MagGlass.PrintStatus lPxColor, tCursorPos, True
        End If
    End If
    'auf doppelte Farben prüfen
    For i = 0 To UBound(ColorCollection)
        If ColorCollection(i) = lPxColor Then Exit Sub
    Next i
    AddColor lPxColor
End Sub

Public Function FileExists(FileName As String) As Boolean
Dim res As Long
  res = GetFileAttributes(StrPtr(FileName))
  If res = -1 Then Exit Function
  FileExists = Not CBool(res And (vbDirectory Or vbVolume))
End Function

Public Sub FillMenuColorCollection(f As Form, Id As Long)
Dim i As Integer
Dim mnuID As Long, h1 As Long, h2 As Long, h3 As Long
    h1 = GetMenu(f.hwnd)
    h2 = GetSubMenu(h1, 0)
    h3 = GetSubMenu(h2, Id)
    For i = 0 To UBound(ColorCollection)
        If ColorCode = PL_HEXHTML Then
            f.mnuColorCollectionItems(i).Caption = RGBtoHTML(ColorCollection(i))
        ElseIf ColorCode = PL_HEXVB Then
            f.mnuColorCollectionItems(i).Caption = RGBtoVB(ColorCollection(i))
        Else
            f.mnuColorCollectionItems(i).Caption = ColorCollection(i)
        End If
        With frmMenu
            .picMenuColor(i).Line (0, 0)-(18, 18), ColorCollection(i), BF
            .picMenuColor(i).Picture = .picMenuColor(i).Image
            mnuID = GetMenuItemID(h3, CLng(i))
            SetMenuItemBitmaps h3, i, MF_BYPOSITION, .picMenuColor(i).Picture, .picMenuColor(i).Picture
        End With
    Next i
End Sub

Public Function GetFileExtension(FilePath As String)
  GetFileExtension = Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
End Function


Public Function GetMousePos(Parent As Object, X As Single, Y As Single, Optional Border As Single = 100) As MousePos
Dim IsLeft As Boolean, IsTop As Boolean, IsRight As Boolean, IsBottom As Boolean
    IsLeft = X < Border
    IsTop = Y < Border
    IsRight = X > Parent.ScaleWidth - Border
    IsBottom = Y > Parent.ScaleHeight - Border
    Select Case True
        Case IsTop And IsLeft:      GetMousePos = mpTopLeft
        Case IsBottom And IsLeft:   GetMousePos = mpBottomLeft
        Case IsTop And IsRight:     GetMousePos = mpTopRight
        Case IsBottom And IsRight:  GetMousePos = mpBottomRight
        Case IsLeft:    GetMousePos = mpLeft
        Case IsTop:     GetMousePos = mpTop
        Case IsRight:   GetMousePos = mpRight
        Case IsBottom:  GetMousePos = mpBottom
        Case Else:      GetMousePos = mpOther
      End Select
End Function

Public Function IsLightColor(c As Long, Optional ref As Long = 700) As Boolean
Dim b As Long, g As Long, r As Long
    b = c \ 65536
    g = (c - b * 65536) \ 256
    r = c - (b * 65536) - (g * 256)
    IsLightColor = (b + g + r > ref)
End Function

Public Function Lighten(c As Long) As Long
Dim b As Long, g As Long, r As Long
    b = c \ 65536
    g = (c - b * 65536) \ 256
    r = c - (b * 65536) - (g * 256)
    b = b + ((255 - b) \ 3)
    g = g + ((255 - g) \ 3)
    r = r + ((255 - r) \ 3)
    Lighten = RGB(r, g, b)
End Function


Public Function RGBtoHTML(RGBColor As Long, Optional extended As Boolean) As String
Dim s As String
  s = Right$("00000" & Hex$(RGBColor), 6)
  RGBtoHTML = "#000000"
  Mid$(RGBtoHTML, 6&, 2&) = Left$(s, 2&)
  Mid$(RGBtoHTML, 4&, 2&) = Mid$(s, 3&, 2&)
  Mid$(RGBtoHTML, 2&, 2&) = Right$(s, 2&)
  If extended Then
    RGBtoHTML = RGBtoHTML & "   RGB: (" & CStr(RGBColor And vbRed) & "," & CStr((RGBColor And vbGreen) \ &H100) & "," & CStr((RGBColor And vbBlue) \ &H10000) & ")"
  End If

End Function

Public Function RGBtoVB(RGBColor As Long, Optional extended As Boolean) As String
Dim s As String
Dim l As Long
Const c = 9&
  s = Hex$(RGBColor): l = Len(s)
  RGBtoVB = "&H000000"
  Mid$(RGBtoVB, c - l, l) = s
  If extended Then
    RGBtoVB = RGBtoVB & "      RGB: (" & CStr(RGBColor And vbRed) & "," & CStr((RGBColor And vbGreen) \ &H100) & "," & CStr((RGBColor And vbBlue) \ &H10000) & ")"
  End If
End Function

Public Function RTrimNull(s As String) As String
    RTrimNull = s
    Do
        If Right$(RTrimNull, 1&) <> vbNullChar Then Exit Do
        RTrimNull = Left$(RTrimNull, Len(RTrimNull) - 1&)
    Loop
End Function

Public Function ShellExec(ByVal DokPfad As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    On Error Resume Next
    ShellExec = (ShellExecuteA(0&, "open", DokPfad, vbNullString, vbNullString, WindowStyle) > 32)
End Function

Public Function ShowColorDlg(ByVal hwnd As Long, ByVal col As Long) As Long
Dim nResult As Long
Dim pDialog As LPCHOOSECOLOR
  On Error GoTo ShowColorDlg_Error
  With pDialog
    .lStructSize = Len(pDialog)
    .hwnd = hwnd
    .hInstance = App.hInstance
    .Flags = CC_SOLIDCOLOR Or CC_RGBINIT Or CC_ANYCOLOR Or CC_FULLOPEN
    .lpCustColors = String$(16 * 4, 0)
    .rgbResult = col

    nResult = ChooseColor(pDialog)
    If nResult <> 0 Then
      ShowColorDlg = .rgbResult
    Else
      ShowColorDlg = -1
    End If
  End With
Exit Function

ShowColorDlg_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: modMain.ShowColorDlg." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Function

Public Function ShowOpenFileDlg(Filter As String, _
                                Flags As Long, _
                                hwnd As Long, _
                                Optional Title As String = "Bild laden...", _
                                Optional InitialDir As String, _
                                Optional nFilterIndex As Long = 1) As String
Dim Buffer$, Result&
Dim ComDlgOpenFileName As OPENFILENAME
Dim Files() As String
Dim i As Integer
  
  If Flags And OFN_ALLOWMULTISELECT Then
    Buffer = String$(32768, 0)
  Else
    Buffer = String$(256, 0)
  End If
   
    With ComDlgOpenFileName
      .lStructSize = Len(ComDlgOpenFileName)
      .lpstrTitle = Title & vbNullChar
      .hWndOwner = hwnd
      .Flags = Flags
      .nFilterIndex = nFilterIndex
      .nMaxFile = Len(Buffer)
      .lpstrFile = Buffer
      .lpstrFilter = Filter
      .lpstrInitialDir = InitialDir
    End With
    Result = GetOpenFileNameW(ComDlgOpenFileName)
    
    If Result <> 0 Then
      Files = Split(RTrimNull(ComDlgOpenFileName.lpstrFile), Chr$(0))
      If UBound(Files) = 0 Then
        ShowOpenFileDlg = Files(0)
      Else
        ShowOpenFileDlg = Files(0) & "\" & Files(1)
        For i = 2 To UBound(Files)
            ShowOpenFileDlg = ShowOpenFileDlg & "##" & Files(0) & "\" & Files(i)
        Next
      End If
    End If

End Function

Public Function ShowSaveFileDlg(Filter As String, Flags As Long, hwnd As Long, _
                                FileName As String, Optional InitialDir As String, Optional Extension As String) As String
Dim Buffer As String
Dim res As Long
Dim ComDlgOpenFileName As OPENFILENAME
  
    Buffer = FileName & String$(256 - Len(FileName), 0)
    With ComDlgOpenFileName
      .lStructSize = Len(ComDlgOpenFileName)
      .hWndOwner = hwnd
      .Flags = Flags
      .nFilterIndex = 1
      .nMaxFile = Len(Buffer)
      .lpstrFile = Buffer
      .lpstrFilter = Filter
      .lpstrInitialDir = InitialDir
      .lpstrDefExt = Extension
      Select Case Extension
        Case "bmp": .nFilterIndex = 2
        Case "gif": .nFilterIndex = 3
        Case "png": .nFilterIndex = 4
        Case Else: .nFilterIndex = 0
      End Select
    End With

    res = GetSaveFileNameW(ComDlgOpenFileName)
    
    If res <> 0 Then
        ShowSaveFileDlg = Left$(ComDlgOpenFileName.lpstrFile, _
                        InStr(ComDlgOpenFileName.lpstrFile, _
                        Chr$(0)) - 1)
        Select Case ComDlgOpenFileName.nFilterIndex
            Case 1: Extension = "jpg"
            Case 2: Extension = "bmp"
            Case 3: Extension = "gif"
            Case 4: Extension = "png"
            Case Else: Extension = "jpg"
        End Select
        If LCase$(GetFileExtension(ShowSaveFileDlg)) <> Extension Then ShowSaveFileDlg = ShowSaveFileDlg & "." & Extension
            
    End If
End Function

Public Sub TransparencyRuler(hwnd As Long, Rate As Byte)
    'Rate: 254 = normal 0 = ganz transparent (also unsichtbar)
    Dim WinInfo As Long
    On Error Resume Next
    WinInfo = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    If Rate < 255 Then
        WinInfo = WinInfo Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, WinInfo
        SetLayeredWindowAttributes hwnd, 0, Rate, LWA_ALPHA
    Else
        'Wenn als Rate 255 angegeben wird, so wird der Ausgangszustand wiederhergestellt
        WinInfo = WinInfo Xor WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, WinInfo
    End If
End Sub

Public Function GetInfo() As String
  Dim s As String
  s = "P I X E L - L I N E A L" & vbCrLf & vbCrLf
  s = s & "Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
  s = s & "-Freeware-" & vbCrLf & vbCrLf
  s = s & "Autor: WW-Anwendungsentwicklung" & vbCrLf
  s = s & "https://www.ww-a.de"
  GetInfo = s
End Function

Public Sub CheckVersion(Optional autoCheck As Boolean)
Dim xmlhttp As Object
Dim lastVerInfo As String, lastVerDate As String
Dim lastVerArray() As String
    On Error GoTo CheckVersion_Error
    If autoCheck Then
        lastVerDate = GetSetting(App.Title, "Options", "VerInfo", "")
        If IsDate(lastVerDate) Then
            If DateDiff("d", Now, CDate(lastVerDate)) < 2 Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    With xmlhttp
        .Open "GET", "https://docs.ww-a.de/lib/exe/fetch.php/pixellineal:verinfo.txt", False
        .setRequestHeader "Content-Type", "text/plain"
        .setRequestHeader "Connection", "keep-alive"
        .send
        lastVerInfo = .responseText
    End With
    If Len(lastVerInfo) Then
        lastVerArray = Split(lastVerInfo, vbTab)
        If LCase(lastVerArray(0)) = "pixlin.exe" Then
            lastVerDate = lastVerArray(2)
            lastVerInfo = lastVerArray(1)
            lastVerArray = Split(lastVerInfo, ".")
            If App.Major < CInt(lastVerArray(0)) Or _
               App.Minor < CInt(lastVerArray(1)) Or _
               App.Revision < CInt(lastVerArray(3)) Then
                    If MsgBox("Neue Version " & lastVerInfo & " vom " & lastVerDate & " gefunden." & vbCrLf & "Möchtest du jetzt die neue Version herunterladen?", vbYesNo Or vbInformation Or vbDefaultButton1, "Pixel-Lineal V" & App.Major & "." & App.Minor & ".0." & App.Revision) = vbYes Then
                        On Error GoTo ShellExec_Error
                        CloseApp = True
                        ShellExec "https://docs.ww-a.de/doku.php/pixellineal:installation", vbNormalFocus
                    End If
            Else
                If autoCheck = False Then
                    If MsgBox("Die Version " & App.Major & "." & App.Minor & ".0." & App.Revision & " ist aktuell." & vbCrLf & _
                              "Soll zukünftig automatisch auf neue Versionen geprüft werden?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Online-Update") = vbYes Then
                        SaveSetting App.Title, "Options", "VerInfo", Format$(Now, "dd.mm.yyyy")
                    Else
                        On Error Resume Next
                        DeleteSetting App.Title, "Options", "VerInfo"
                    End If
                End If
            End If
            If autoCheck Then SaveSetting App.Title, "Options", "VerInfo", Format$(Now, "dd.mm.yyyy")
        End If
    End If
    
    
    Exit Sub
    
CheckVersion_Error:
    Exit Sub
    
ShellExec_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: CheckVersion." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

