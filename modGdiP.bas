Attribute VB_Name = "modGdiP"
Option Explicit

Private Const GDI_PLUS_VERSION As Long = 1
Private Const MIME_JPG As String = "image/jpeg"
Private Const MIME_BMP As String = "image/bmp"
Private Const MIME_PNG As String = "image/png"
Private Const MIME_GIF As String = "image/gif"
Private Const ENCODER_PARAMETER_VALUE_TYPE_LONG As Long = 4
Private Const ENCODER_QUALITY As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

' ----==== Sonstige Types ====----
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' ----==== GDIPlus Types ====----
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    Guid As Guid
    NumberOfValues As Long
    Type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(15) As EncoderParameter
End Type

Private Type ImageCodecInfo
    Clsid As Guid
    FormatID As Guid
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As Long
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

Private Type COLORMATRIX
    m(0 To 4, 0 To 4) As Single
End Type

' ----==== GDIPlus Enums ====----
Private Enum eGdiStatus 'GDI+ eGdiStatus
    gdiOk = 0
    gdiGenericError = 1
    gdiInvalidParameter = 2
    gdiOutOfMemory = 3
    gdiObjectBusy = 4
    gdiInsufficientBuffer = 5
    gdiNotImplemented = 6
    gdiWin32Error = 7
    gdiWrongState = 8
    gdiAborted = 9
    gdiFileNotFound = 10
    gdiValueOverflow = 11
    gdiAccessDenied = 12
    gdiUnknownImageFormat = 13
    gdiFontFamilyNotFound = 14
    gdiFontStyleNotFound = 15
    gdiNotTrueTypeFont = 16
    gdiUnsupportedGdiplusVersion = 17
    gdiNotInitialized = 18
    gdiPropertyNotFound = 19
    gdiPropertyNotSupported = 20
    gdiProfileNotFound = 21
End Enum

Public Enum RotateFlipType
    RotateNoneFlipNone = 0
    Rotate90FlipNone = 1
    Rotate180FlipNone = 2
    Rotate270FlipNone = 3
    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7
    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX
    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum

Private Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Private Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap = 1
    ColorAdjustTypeBrush = 2
    ColorAdjustTypePen = 3
    ColorAdjustTypeText = 4
    ColorAdjustTypeCount = 5
    ColorAdjustTypeAny = 6
End Enum

Public Enum Unit
    UnitWorld = 0
    UnitDisplay = 1
    UnitPixel = 2
    UnitPoint = 3
    UnitInch = 4
    UnitDocument = 5
    UnitMillimeter = 6
End Enum

' ----==== GDI+ API Declarationen ====----

Private Declare Function GdipGetImageDimension Lib "gdiplus" _
    (ByVal Image As Long, ByRef sngWidth As Single, _
    ByRef sngHeight As Single) As eGdiStatus

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef Image As Long) As eGdiStatus
    
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" ( _
    ByVal hDC As Long, ByRef graphics As Long _
    ) As eGdiStatus
    
Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef Token As Long, ByRef lpInput As GdiplusStartupInput, _
    Optional ByRef lpOutput As Any) As eGdiStatus

Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal Token As Long) As eGdiStatus

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef BITMAP As Long) As eGdiStatus

Private Declare Function GdipSaveImageToFile Lib "gdiplus" _
    (ByVal Image As Long, ByVal FileName As Long, _
    ByRef clsidEncoder As Guid, _
    ByRef encoderParams As Any) As eGdiStatus

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal BITMAP As Long, ByRef hbmReturn As Long, _
    ByVal background As Long) As eGdiStatus

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" _
    (ByVal hbm As Long, ByVal hPal As Long, _
    ByRef BITMAP As Long) As eGdiStatus

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" _
    (ByRef numEncoders As Long, ByRef Size As Long) As eGdiStatus

Private Declare Function GdipGetImageEncoders Lib "gdiplus" _
    (ByVal numEncoders As Long, ByVal Size As Long, _
    ByRef Encoders As Any) As eGdiStatus

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal Image As Long) As eGdiStatus
    
Private Declare Function GdipImageRotateFlip Lib "gdiplus" _
    (ByVal Image As Long, ByVal rfType As RotateFlipType) As eGdiStatus
    
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" _
    (ByRef imageattr As Long) As eGdiStatus
    
Private Declare Function GdipDeleteGraphics Lib "gdiplus" _
    (ByVal graphics As Long) As eGdiStatus

Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" _
    (ByVal imageattr As Long) As eGdiStatus
    
Private Declare Function GdipGetImageThumbnail Lib "gdiplus" _
    (ByVal Image As Long, ByVal thumbWidth As Long, _
    ByVal thumbHeight As Long, ByRef thumbImage As Long, _
    ByVal callback As Long, ByVal callbackData As Long) _
    As eGdiStatus

Private Declare Function GdipSetImagePalette Lib "gdiplus" _
    (ByVal pImage As Long, ByRef Palette As ColorPalette) As Long
    
Private Type ColorPalette            ' GDI+ palette object
   Flags As Long
   Count As Long
   Entries(0 To 255) As Long
End Type

' ----==== OLE API Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32" _
    (ByVal str As Long, id As Guid) As Long

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
    (lpPictDesc As PictDesc, riid As IID, ByVal fOwn As Boolean, _
    lplpvObj As Object)

' ----==== Kernel API Declarations ====----
Private Declare Function lstrlenW Lib "kernel32" _
    (lpString As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" _
    (lpString1 As Any, lpString2 As Any) As Long
        
Private Declare Function GdipSetImageAttributesColorMatrix _
    Lib "gdiplus" (ByVal imageattr As Long, _
    ByVal ColorAdjust As ColorAdjustType, _
    ByVal EnableFlag As Boolean, _
    ByRef MatrixColor As COLORMATRIX, _
    ByRef MatrixGray As COLORMATRIX, _
    ByVal Flags As ColorMatrixFlags) As eGdiStatus
    
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" _
    (ByVal graphics As Long, ByVal Image As Long, _
    ByVal dstx As Single, ByVal dsty As Single, _
    ByVal dstwidth As Single, ByVal dstheight As Single, _
    ByVal srcx As Single, ByVal srcy As Single, _
    ByVal srcwidth As Single, ByVal srcheight As Single, _
    ByVal srcUnit As Unit, ByVal imageAttributes As Long, _
    ByVal callback As Long, ByVal callbackData As Long) As eGdiStatus
    
    
    
' ----==== Variablen ====----
Private GdipToken As Long
Private GdiPInitialized As Boolean

'------------------------------------------------------
' Funktion     : StartUpGDIPlus
' Beschreibung : Initialisiert GDI+ Instanz
' Übergabewert : GDI+ Version
' Rückgabewert : GDI+ eGdiStatus
'------------------------------------------------------
Private Function StartUpGDIPlus() As eGdiStatus
    If GdiPInitialized Then
        StartUpGDIPlus = gdiOk
    Else
        ' Initialisieren der GDI+ Instanz
        Dim GdipStartupInput As GdiplusStartupInput
        GdipStartupInput.GdiplusVersion = GDI_PLUS_VERSION
        StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
        If StartUpGDIPlus <> gdiOk Then
            Err.Raise vbObjectError, "GDI_PLUS", "GDI+ not inizialized."
        Else
            GdiPInitialized = True
        End If
    End If
End Function

'------------------------------------------------------
' Funktion     : ShutdownGDIPlus
' Beschreibung : Beendet die GDI+ Instanz
' Rückgabewert : GDI+ eGdiStatus
'------------------------------------------------------
Public Sub ShutdownGDIPlus()
    ' Beendet GDI+ Instanz
    If GdiPInitialized Then
        GdiplusShutdown GdipToken
        GdiPInitialized = False
    End If
End Sub

'------------------------------------------------------
' Funktion     : Execute
' Beschreibung : Gibt im Fehlerfall die entsprechende
'                GDI+ Fehlermeldung aus
' Übergabewert : GDI+ eGdiStatus
' Rückgabewert : GDI+ eGdiStatus
'------------------------------------------------------
Private Function Execute(ByVal lReturn As eGdiStatus) As eGdiStatus
    Dim lCurErr As eGdiStatus
    If lReturn = gdiOk Then
        lCurErr = gdiOk
    Else
        lCurErr = lReturn
        MsgBox GdiErrorString(lReturn) & " GDI+ Error: " & lReturn, _
            vbOKOnly, "GDI Error"
    End If
    Execute = lCurErr
End Function

'------------------------------------------------------
' Funktion     : GdiErrorString
' Beschreibung : Umwandlung der GDI+ Statuscodes in Stringcodes
' Übergabewert : GDI+ eGdiStatus
' Rückgabewert : Fehlercode als String
'------------------------------------------------------
Private Function GdiErrorString(ByVal lError As eGdiStatus) As String
    Dim s As String
    
    Select Case lError
    Case gdiGenericError:              s = "Allgemeiner Fehler: Beim Methodenaufruf ist ein unbekannter Fehler aufgetreten."
    Case gdiInvalidParameter:          s = "Ungültiger Parameter: Eines der an die Methode übergebenen Argumente war nicht gültig."
    Case gdiOutOfMemory:               s = "Out Of Memory: Das Betriebssystem hat nicht genügend Arbeitsspeicher und konnte keinen Speicher zum Verarbeiten des Methodenaufrufs zuweisen."
    Case gdiObjectBusy:                s = "Object Busy. Eines der im API-Aufruf angegebenen Argumente wird bereits in einem anderen Thread verwendet."
    Case gdiInsufficientBuffer:        s = "Insufficient Buffer: Ein Puffer, der im API-Aufruf als Argument angegeben wurde, ist nicht groß genug, um die zu empfangenden Daten zu speichern."
    Case gdiNotImplemented:            s = "Nicht implementiert: Die Methode ist nicht implementiert"
    Case gdiWin32Error:                s = "Win32 Fehler: Die Methode hat einen Win32-Fehler generiert. Eventuell ist ein Zugriff auf den Zielpfad nicht möglich."
    Case gdiWrongState:                s = "Falscher eGdiStatus: Das Objekt befindet sich in einem ungültigen eGdiStatus, um den API-Aufruf zu erfüllen."
    Case gdiAborted:                   s = "Abgebrochen: Die Methode wurde abgebrochen."
    Case gdiFileNotFound:              s = "Datei nicht gefunden: Die angegebene Bilddatei oder Metadatei konnte nicht gefunden werden."
    Case gdiValueOverflow:             s = "Wertüberlauf: Die Methode hat eine arithmetische Operation ausgeführt, die zu einem numerischen Überlauf geführt hat."
    Case gdiAccessDenied:              s = "Zugriff abgelehnt:  Eine Schreiboperation für die angegebene Datei ist nicht zulässig."
    Case gdiUnknownImageFormat:        s = "Unbekanntes Bildformat: Das angegebene Bilddateiformat ist nicht bekannt."
    Case gdiFontFamilyNotFound:        s = "Schriftfamilie nicht gefunden: Die angegebene Schriftfamilie konnte nicht gefunden werden. Entweder ist der Name der Schriftart falsch oder die Schriftfamilie ist nicht installiert."
    Case gdiFontStyleNotFound:         s = "Stil nicht verfügbar: Der angegebene Stil für die angegebene Schriftfamilie ist nicht verfügbar."
    Case gdiNotTrueTypeFont:           s = "Keine TrueType-Schrift: Die abgerufene Schriftart ist keine TrueType-Schriftart und kann nicht mit GDI+ verwendet werden."
    Case gdiUnsupportedGdiplusVersion: s = "Nicht unterstützte Gdi+ Version: Die Version von GDI+, die auf dem System installiert ist, ist nicht mit der Version kompatibel, mit der die Anwendung kompiliert wurde."
    Case gdiNotInitialized:            s = "Gdiplus Not Initialized. Gibt an, dass sich die GDI+ API nicht in einem initialisierten Zustand befindet, die GdiplusStartup-Methode wurde nicht durchgeführt"
    Case gdiPropertyNotFound:          s = "Eigenschaft nicht gefunden: Die angegebene Eigenschaft ist nicht im Bild vorhanden."
    Case gdiPropertyNotSupported:      s = "Eigenschaft nicht unterstützt: Die angegebene Eigenschaft wird nicht vom Bildformat unterstützt und kann daher nicht festgelegt werden."
    Case Else:                         s = "Unbekannter GDI+ Fehler."
    End Select
    
    GdiErrorString = s
End Function

'------------------------------------------------------
' Funktion     : LoadPicturePlus
' Beschreibung : Lädt ein Bilddatei per GDI+
' Übergabewert : Pfad\Dateiname der Bilddatei
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function LoadPicturePlus(ByVal FileName As String) As StdPicture
    Dim retStatus As eGdiStatus
    Dim lBitmap As Long
    Dim hBitmap As Long
    
    StartUpGDIPlus
    ' Öffnet die Bilddatei in lBitmap
    retStatus = Execute(GdipCreateBitmapFromFile(StrPtr(FileName), _
        lBitmap))
    
    If retStatus = gdiOk Then
        
        ' Erzeugen einer GDI Bitmap lBitmap -> hBitmap
        retStatus = Execute(GdipCreateHBITMAPFromBitmap(lBitmap, _
            hBitmap, 0))
        
        If retStatus = gdiOk Then
            ' Erzeugen des StdPicture Objekts von hBitmap
            Set LoadPicturePlus = HandleToPicture(hBitmap, _
                vbPicTypeBitmap)
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

'------------------------------------------------------
' Funktion     : SavePictureAsPNG
' Beschreibung : Speichert ein StdPicture Objekt
'                per GDI+ als PNG
' Übergabewert : Pic = StdPicture Objekt
'                FileName = Pfad\Dateiname.png
' Rückgabewert : True = speichern erfolgreich
'                False = speichern fehlgeschlagen
'------------------------------------------------------
Public Function SavePictureAsPNG(ByVal pic As StdPicture, _
    ByVal sFileName As String) As Boolean
    
    Dim lBitmap As Long
    Dim tPicEncoder As Guid
    
    StartUpGDIPlus
    ' Erzeugt eine GDI+ Bitmap vom
    ' StdPicture Handle -> lBitmap
    If Execute(GdipCreateBitmapFromHBITMAP( _
    pic.Handle, 0, lBitmap)) = gdiOk Then
        
        ' Ermitteln der CLSID vom mimeType Encoder
        If GetEncoderClsid(MIME_PNG, tPicEncoder) = True Then
            
            ' Speichert lBitmap als PNG
            If Execute(GdipSaveImageToFile(lBitmap, _
            StrPtr(sFileName), tPicEncoder, ByVal 0)) = gdiOk Then
                
                ' speichern erfolgreich
                SavePictureAsPNG = True
            Else
                ' speichern nicht erfolgreich
                SavePictureAsPNG = False
            End If
        Else
            ' speichern nicht erfolgreich
            SavePictureAsPNG = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", _
            vbOKOnly, "Encoder Error"
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

'------------------------------------------------------
' Funktion     : SavePictureAsGIF
' Beschreibung : Speichert ein StdPicture Objekt
'                per GDI+ als GIF
' Übergabewert : Pic = StdPicture Objekt
'                FileName = Pfad\Dateiname.gif
' Rückgabewert : True = speichern erfolgreich
'                False = speichern fehlgeschlagen
'------------------------------------------------------
Public Function SavePictureAsGIF(ByVal pic As StdPicture, _
    ByVal sFileName As String) As Boolean
    
    Dim lBitmap As Long
    Dim tPicEncoder As Guid
    
    StartUpGDIPlus
    ' Erzeugt eine GDI+ Bitmap vom
    ' StdPicture Handle -> lBitmap
    If Execute(GdipCreateBitmapFromHBITMAP( _
    pic.Handle, 0, lBitmap)) = gdiOk Then
        
        ' Ermitteln der CLSID vom mimeType Encoder
        If GetEncoderClsid(MIME_GIF, tPicEncoder) = True Then
            
            ' Speichert lBitmap als GIF
            If Execute(GdipSaveImageToFile(lBitmap, _
            StrPtr(sFileName), tPicEncoder, ByVal 0)) = gdiOk Then
                
                ' speichern erfolgreich
                SavePictureAsGIF = True
            Else
                ' speichern nicht erfolgreich
                SavePictureAsGIF = False
            End If
        Else
            ' speichern nicht erfolgreich
            SavePictureAsGIF = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", vbOKOnly, "Encoder Error"
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

'------------------------------------------------------
' Funktion     : SavePictureAsJPG
' Beschreibung : Speichert ein StdPicture Objekt per GDI+ als JPG
' Übergabewert : Pic = StdPicture Objekt
'                FileName = Pfad\Dateiname.jpg
'                Quality = JPG Kompression
' Rückgabewert : True = speichern erfolgreich
'                False = speichern fehlgeschlagen
'------------------------------------------------------
Public Function SavePictureAsJPG(ByVal pic As StdPicture, _
    ByVal FileName As String, Optional ByVal Quality As Long = 85) _
    As Boolean
    
    Dim retStatus As eGdiStatus
    Dim retVal As Boolean
    Dim lBitmap As Long
    Debug.Print pic.Handle
    StartUpGDIPlus
    ' Erzeugt eine GDI+ Bitmap vom StdPicture Handle -> lBitmap
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(pic.Handle, 0, _
        lBitmap))
    
    If retStatus = gdiOk Then
        
        Dim PicEncoder As Guid
        Dim tParams As EncoderParameters
        
        '// Ermitteln der CLSID vom mimeType Encoder
        retVal = GetEncoderClsid(MIME_JPG, PicEncoder)
        If retVal = True Then
            
            If Quality > 100 Then Quality = 100
            If Quality < 0 Then Quality = 0
            
            ' Initialisieren der Encoderparameter
            tParams.Count = 1
            With tParams.Parameter(0) ' Quality
                ' Setzen der Quality GUID
                CLSIDFromString StrPtr(ENCODER_QUALITY), .Guid
                .NumberOfValues = 1
                .Type = ENCODER_PARAMETER_VALUE_TYPE_LONG
                .Value = VarPtr(Quality)
            End With
            
            ' Speichert lBitmap als JPG
            retStatus = Execute(GdipSaveImageToFile(lBitmap, _
                StrPtr(FileName), PicEncoder, tParams))
            
            If retStatus = gdiOk Then
                SavePictureAsJPG = True
            Else
                SavePictureAsJPG = False
            End If
        Else
            SavePictureAsJPG = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", _
            vbOKOnly, "Encoder Error"
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

'------------------------------------------------------
' Funktion     : SavePictureAsBMP
' Beschreibung : Speichert ein StdPicture Objekt per GDI+ als BMP
' Übergabewert : Pic = StdPicture Objekt
'                FileName = Pfad\Dateiname.bmp
' Rückgabewert : True = speichern erfolgreich
'                False = speichern fehlgeschlagen
'------------------------------------------------------
Public Function SavePictureAsBMP(ByVal pic As StdPicture, _
    ByVal sFileName As String) As Boolean
    
    Dim lBitmap As Long
    Dim tPicEncoder As Guid
    
    StartUpGDIPlus
    ' Erzeugt eine GDI+ Bitmap vom StdPicture Handle -> lBitmap
    If Execute(GdipCreateBitmapFromHBITMAP( _
    pic.Handle, 0, lBitmap)) = gdiOk Then
        
        ' Ermitteln der CLSID vom mimeType Encoder
        If GetEncoderClsid(MIME_BMP, tPicEncoder) = True Then
            
            ' Speichert lBitmap als GIF
            If Execute(GdipSaveImageToFile(lBitmap, _
                StrPtr(sFileName), tPicEncoder, ByVal 0)) = gdiOk Then
                
                ' speichern erfolgreich
                SavePictureAsBMP = True
            Else
                ' speichern nicht erfolgreich
                SavePictureAsBMP = False
            End If
        Else
            ' speichern nicht erfolgreich
            SavePictureAsBMP = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", vbOKOnly, "Encoder Error"
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function
'------------------------------------------------------
' Funktion     : FlipRotatePicture
' Beschreibung : Drehen von Bildern per GDI+
' Übergabewert : Pic = StdPicture
'                FlipRotate = RotateFlipType
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function FlipRotatePicture(ByVal pic As StdPicture, _
    Optional ByVal FlipRotate As RotateFlipType = _
    RotateNoneFlipNone) As StdPicture
    
    Dim retStatus As eGdiStatus
    Dim lBitmap As Long
    Dim hBitmap As Long
    
    StartUpGDIPlus
    ' Erzeuge ein GDI+ Bitmap vom Image Handle
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(pic.Handle, 0, lBitmap))
    
    If retStatus = gdiOk Then
        
        ' FlipRotate
        retStatus = Execute(GdipImageRotateFlip(lBitmap, FlipRotate))
        
        If retStatus = gdiOk Then
            
            ' Erzeugen der GDI bitmap
            retStatus = Execute(GdipCreateHBITMAPFromBitmap(lBitmap, _
                hBitmap, 0))
            
            If retStatus = gdiOk Then
                ' Erzeugen des StdPicture Objekts
                Set FlipRotatePicture = HandleToPicture(hBitmap, _
                    vbPicTypeBitmap)
            End If
            
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

Public Function SetBrightnessContrast(ByVal oPicBox As PictureBox, _
        Optional ByVal sBrightness As Single = 0, _
        Optional ByVal sContrast As Single = 0) As Boolean
    
    Dim lBitmap As Long
    Dim lGraphics As Long
    Dim lAttribute As Long
    Dim sWidth As Single
    Dim sHeight As Single
    Dim lOldScaleMode As Long
    Dim bOldAutoRedraw As Boolean
    Dim tMatrixColor As COLORMATRIX
    Dim tMatrixGray As COLORMATRIX
    Dim sDiff As Single
    Dim retStatus As eGdiStatus
    
    Dim bRet As Boolean
    
    StartUpGDIPlus
    
    lBitmap = oPicBox.Picture.Handle
    
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(oPicBox.Picture.Handle, 0, lBitmap))
    If retStatus <> gdiOk Then Exit Function
    
    ' Parameter zwischenspeichern und setzen
    With oPicBox
        lOldScaleMode = .ScaleMode
        bOldAutoRedraw = .AutoRedraw
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .Cls
    End With
    
    ' Min/Max
    If sBrightness < -1 Then sBrightness = -1
    If sBrightness > 1 Then sBrightness = 1
    If sContrast < -1 Then sContrast = -1
    If sContrast > 1 Then sContrast = 1
    
    ' Differenz berechnen zur korrekten Darstellung
    ' beim verändern des Kontrastwertes
    sDiff = (sBrightness / 2) - (sContrast / 2)
    
    ' ColorMatrix Parameter setzen
    With tMatrixColor
        .m(0, 0) = 1 + sContrast: .m(0, 4) = sBrightness + sDiff
        .m(1, 1) = 1 + sContrast: .m(1, 4) = sBrightness + sDiff
        .m(2, 2) = 1 + sContrast: .m(2, 4) = sBrightness + sDiff
        .m(3, 3) = 1
        .m(4, 4) = 1
    End With
    
    ' Dimensionen von lBitmap ermitteln
    If Execute(GdipGetImageDimension(lBitmap, _
    sWidth, sHeight)) = gdiOk Then
        
        ' Graphicsobjekt vom HDC erstellen -> lGraphics
        If Execute(GdipCreateFromHDC(oPicBox.hDC, _
        lGraphics)) = gdiOk Then
            
            ' ImageAttributeobjekt erstellen -> lAttribute
            If Execute(GdipCreateImageAttributes(lAttribute)) _
            = gdiOk Then
                
                ' ColorMatrix an ImageAttributeobjekt übergeben
                If Execute(GdipSetImageAttributesColorMatrix( _
                lAttribute, ColorAdjustTypeDefault, True, _
                tMatrixColor, tMatrixGray, _
                ColorMatrixFlagsDefault)) = gdiOk Then
                    
                    ' zeichnet lBitmap in das Graphicsobjekt
                    ' lGraphics mit dem entsprechenden ImageAttribute
                    ' und Dimensionen
                    If Execute(GdipDrawImageRectRect(lGraphics, _
                    lBitmap, 0, 0, sWidth, sHeight, _
                    0, 0, sWidth, sHeight, UnitPixel, _
                    lAttribute, 0, 0)) = gdiOk Then
                        
                        bRet = True
                        
                    End If
                End If
                
                ' lAttribute löschen
                Call Execute(GdipDisposeImageAttributes(lAttribute))
            End If
            
            ' lGraphics löschen
            Call Execute(GdipDeleteGraphics(lGraphics))
        End If
    End If
    
    ' zwichengespeicherte Werte zurücksetzen
    With oPicBox
        .ScaleMode = lOldScaleMode
        .AutoRedraw = bOldAutoRedraw
        .Refresh
    End With
    
    ' Rückgabewert übergeben
    SetBrightnessContrast = bRet
    ' Lösche lBitmap
    Call Execute(GdipDisposeImage(lBitmap))
End Function


'------------------------------------------------------
' Funktion     : HandleToPicture
' Beschreibung : Umwandeln einer GDI+ Bitmap Handle in ein StdPicture Objekt
' Übergabewert : hGDIHandle = GDI+ Bitmap Handle
'                ObjectType = Bitmaptyp
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Private Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hPal As Long = 0) As StdPicture
    
    Dim tPictDesc As PictDesc
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
        .cbSizeofStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hPal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
    
    ' Rückgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    
End Function

'------------------------------------------------------
' Funktion     : GetEncoderClsid
' Beschreibung : Ermittelt die Clsid des Encoders
' Übergabewert : mimeType = mimeType des Encoders
'                pClsid = CLSID des Encoders (in/out)
' Rückgabewert : True = Ermitteln erfolgreich
'                False = Ermitteln fehlgeschlagen
'------------------------------------------------------
Private Function GetEncoderClsid(MimeType As String, pClsid As Guid) As Boolean
Dim Size As Long
Dim pImageCodecInfo() As ImageCodecInfo
Dim j As Long, n As Long
Dim Buffer As String
    
    Call GdipGetImageEncodersSize(n, Size)
    If (Size = 0) Then
        GetEncoderClsid = False  '// fehlgeschlagen
        Exit Function
    End If
    
    ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)
    Call GdipGetImageEncoders(n, Size, pImageCodecInfo(0))
    n = n - 1
    For j = 0 To n
        Buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
        
        Call lstrcpyW(ByVal StrPtr(Buffer), ByVal _
            pImageCodecInfo(j).MimeTypePtr)
            
        If (StrComp(Buffer, MimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo
            GetEncoderClsid = True  '// erfolgreich
            Exit Function
        End If
    Next j
    
    Erase pImageCodecInfo
    GetEncoderClsid = False  '// fehlgeschlagen
End Function

Public Function CreateThumbnail(ByVal pic As StdPicture, ByVal IW As Long, IH As Long) As StdPicture
    
    Dim retStatus As eGdiStatus
    Dim lBitmap As Long
    Dim lThumb As Long
    Dim hBitmap As Long
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    
    StartUpGDIPlus
    ' Erzeuge ein GDI+ Bitmap vom Image Handle
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(pic.Handle, 0, lBitmap))
    
    If retStatus = gdiOk Then
        
        ' Ermitteln der ImageDimensionen
        Call Execute(GdipGetImageDimension(lBitmap, ImageWidth, ImageHeight))
        
        ' Thumbnail erzeugen
        retStatus = Execute(GdipGetImageThumbnail(lBitmap, IW, IH, _
            lThumb, 0, 0))
        
        If retStatus = eGdiStatus.gdiOk Then
            
            ' Erzeugen der GDI Bitmap von der Thumbnail Bitmap
            retStatus = Execute(GdipCreateHBITMAPFromBitmap(lThumb, hBitmap, 0))
            If retStatus = eGdiStatus.gdiOk Then
                ' Erzeugen des StdPicture Objekts von hBitmap
                Set CreateThumbnail = HandleToPicture(hBitmap, vbPicTypeBitmap)
            End If
            
            ' Lösche lThumb
            Call Execute(GdipDisposeImage(lThumb))
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
    End If
End Function

Public Function CopyStdPicture(ByVal pic As StdPicture) As StdPicture
    Dim retStatus As eGdiStatus
Dim lBitmap As Long
Dim hBitmap As Long
    
    StartUpGDIPlus
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(pic.Handle, 0, lBitmap))
    If retStatus = gdiOk Then
            ' Erzeugen der GDI bitmap
            retStatus = Execute(GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0))
            If retStatus = gdiOk Then
                ' Erzeugen des StdPicture Objekts
                Set CopyStdPicture = HandleToPicture(hBitmap, vbPicTypeBitmap)
            End If
    End If
    Call Execute(GdipDisposeImage(lBitmap))

End Function




