Attribute VB_Name = "GdiPlus"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, Id As GUID) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Boolean, lplpvObj As Object)
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef BITMAP As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, ByRef BITMAP As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal mBrush As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal mPen As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal x2 As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal Count As Long) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal brush As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal brush As Long, ByRef points As Any, ByVal Count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, ByRef Encoders As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numEncoders As Long, ByRef Size As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Graphics As Long, ByVal rfType As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal smoothingmode As Long) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type EncoderParameter
    EncoderGUID As GUID
    NumberOfValues As Long
    Type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(15) As EncoderParameter
End Type

Private Type ImageCodecInfo
    Clsid As GUID
    FormatID As GUID
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

Private Const SmoothingModeNone As Long = &H3
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const UnitPixel = 2

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Public Enum SEShapeConstants
    seShapeRectangle = vbShapeRectangle ' 0
    seShapeSquare = vbShapeSquare ' 1
    seShapeOval = vbShapeOval ' 2
    seShapeCircle = vbShapeCircle ' 3
    seShapeRoundedRectangle = vbShapeRoundedRectangle ' 4
    seShapeRoundedSquare = vbShapeRoundedSquare ' 5
    seShapeLine
    seShapePolygon
End Enum

Private mGdipToken As Long

Public Function CopyStdPicture(ByVal Pic As StdPicture, Optional RotateFlip As Long) As StdPicture
Dim lBitmap As Long
Dim hBitmap As Long
    
    If mGdipToken = 0 Then InitGDI
    If GdipCreateBitmapFromHBITMAP(Pic.Handle, 0, lBitmap) = 0 Then
        If RotateFlip Then GdipImageRotateFlip lBitmap, RotateFlip
        If GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0) = 0 Then
            Set CopyStdPicture = HandleToPicture(hBitmap, vbPicTypeBitmap)
        End If
        If lBitmap <> 0 Then GdipDisposeImage lBitmap
    End If

End Function

Public Function OpenPicture(ByVal FileName As String) As StdPicture
Dim ret As Long
Dim lBitmap As Long
Dim hBitmap As Long
    
    If mGdipToken = 0 Then InitGDI
    ret = GdipCreateBitmapFromFile(StrPtr(FileName), lBitmap)
    If ret = 0 Then
        ret = GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0)
        If ret = 0 Then Set OpenPicture = HandleToPicture(hBitmap, vbPicTypeBitmap)
        Call GdipDisposeImage(lBitmap)
    End If
End Function

Public Sub PaintPolygon(Canvas As Object, points() As POINTAPI, _
                 Optional BorderStyle As BorderStyleConstants = vbBSSolid, Optional BorderColor As OLE_COLOR = vbBlack, Optional BorderWidth As Integer = 1, _
                 Optional Filled As Boolean = False, Optional FillColor As OLE_COLOR = vbWhite, Optional Opacity As Single = 100)
Dim iExpandOutsideForAngle As Long, iExpandOutsideForFigure As Long
Dim iGraphics As Long, iExpandForPen As Long
Dim hRgn As Long
Dim rgnRect As RECT
Dim hRgnExpand As Long
Dim prevAutoRedraw As Boolean
Dim prevScaleMode As ScaleModeConstants

    prevAutoRedraw = Canvas.AutoRedraw
    prevScaleMode = Canvas.ScaleMode
    Canvas.ScaleMode = vbPixels
    
    iExpandForPen = BorderWidth / 2
    hRgn = CreateRectRgn(0, 0, 0, 0)
    If GetClipRgn(Canvas.hDC, hRgn) = 0& Then
        DeleteObject hRgn: hRgn = 0
    Else
        GetRgnBox hRgn, rgnRect
        If (iExpandForPen <> 0) Or (iExpandOutsideForAngle <> 0) Or (iExpandOutsideForFigure <> 0) Then
            hRgnExpand = CreateRectRgn(rgnRect.Left - iExpandForPen - iExpandOutsideForAngle - iExpandOutsideForFigure, rgnRect.Top - iExpandForPen - iExpandOutsideForAngle - iExpandOutsideForFigure, rgnRect.Right + iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure, rgnRect.Bottom + iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure)
            SelectClipRgn Canvas.hDC, hRgnExpand
            DeleteObject hRgnExpand
        End If
    End If
    If mGdipToken = 0 Then InitGDI
    If GdipCreateFromHDC(Canvas.hDC, iGraphics) = 0 Then
        Canvas.AutoRedraw = False
        DrawPolygon iGraphics, points, BorderStyle, BorderColor, BorderWidth, Filled, FillColor, Opacity
    End If
    Call GdipDeleteGraphics(iGraphics)
    If hRgnExpand <> 0 Then SelectClipRgn Canvas.hDC, hRgn
    If hRgn <> 0 Then DeleteObject hRgn
    If prevAutoRedraw Then Canvas.Refresh
    Canvas.AutoRedraw = prevAutoRedraw
    Canvas.ScaleMode = prevScaleMode

End Sub

Public Sub PaintShape(Canvas As Object, Shape As SEShapeConstants, ByVal X As Single, ByVal Y As Single, ByVal Width As Long, ByVal Height As Long, _
                 Optional BorderStyle As BorderStyleConstants = vbBSSolid, Optional BorderColor As OLE_COLOR = vbBlack, Optional BorderWidth As Integer = 1, _
                 Optional Filled As Boolean = False, Optional FillColor As OLE_COLOR = vbWhite, Optional Opacity As Single = 100)
Dim iExpandForPen As Long
Dim iExpandOutsideForAngle As Long, iExpandOutsideForFigure As Long
Dim iGraphics As Long
Dim hRgn As Long
Dim rgnRect As RECT
Dim hRgnExpand As Long
Dim prevAutoRedraw As Boolean
Dim prevScaleMode As ScaleModeConstants

    prevAutoRedraw = Canvas.AutoRedraw
    prevScaleMode = Canvas.ScaleMode
    Canvas.ScaleMode = vbPixels
    
    If prevScaleMode <> vbPixels Then
        X = Canvas.ScaleX(X, prevScaleMode, vbPixels)
        Y = Canvas.ScaleY(Y, prevScaleMode, vbPixels)
        Width = Canvas.ScaleX(Width, prevScaleMode, vbPixels)
        Height = Canvas.ScaleX(Height, prevScaleMode, vbPixels)
    End If
    
    iExpandForPen = BorderWidth / 2
    hRgn = CreateRectRgn(0, 0, 0, 0)
    If GetClipRgn(Canvas.hDC, hRgn) = 0& Then
        DeleteObject hRgn: hRgn = 0
    Else
        GetRgnBox hRgn, rgnRect
        If (iExpandForPen <> 0) Or (iExpandOutsideForAngle <> 0) Or (iExpandOutsideForFigure <> 0) Then
            hRgnExpand = CreateRectRgn(rgnRect.Left - iExpandForPen - iExpandOutsideForAngle - iExpandOutsideForFigure, rgnRect.Top - iExpandForPen - iExpandOutsideForAngle - iExpandOutsideForFigure, rgnRect.Right + iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure, rgnRect.Bottom + iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure)
            SelectClipRgn Canvas.hDC, hRgnExpand
            DeleteObject hRgnExpand
        End If
    End If

    If mGdipToken = 0 Then InitGDI
    If GdipCreateFromHDC(Canvas.hDC, iGraphics) = 0 Then
        Canvas.AutoRedraw = False
        Select Case Shape
            Case seShapeCircle
                DrawCircle iGraphics, CLng(X - (Width \ 2)), CLng(Y - (Width \ 2)), Width, BorderStyle, BorderColor, BorderWidth, Filled, FillColor, Opacity
            Case seShapeRectangle
                DrawRectangle iGraphics, CLng(X), CLng(Y), Width, Height, BorderStyle, BorderColor, BorderWidth, Filled, FillColor, Opacity
            Case seShapeLine
                DrawLine iGraphics, CLng(X), CLng(Y), Width, Height, BorderStyle, BorderColor, BorderWidth, Opacity
        End Select
    End If
    Call GdipDeleteGraphics(iGraphics)
    If hRgnExpand <> 0 Then SelectClipRgn Canvas.hDC, hRgn
    If hRgn <> 0 Then DeleteObject hRgn
    If prevAutoRedraw Then Canvas.Refresh
    Canvas.AutoRedraw = prevAutoRedraw
    Canvas.ScaleMode = prevScaleMode
    
End Sub

Public Function ResizePicture(Canvas As PictureBox, Width As Long, Height As Long) As StdPicture
    Dim lBitmap As Long, hBitmap   As Long
    Const InterpolationMode As Long = 7&
    
    If GdipCreateBitmapFromHBITMAP(Canvas.Picture.Handle, 0, lBitmap) = 0 Then
        GdipCreateFromHDC Canvas.hDC, hBitmap
        GdipSetInterpolationMode hBitmap, InterpolationMode
        GdipDrawImageRectI hBitmap, lBitmap, 0, 0, Width, Height
        If hBitmap <> 0 Then GdipDeleteGraphics hBitmap
        If lBitmap <> 0 Then GdipDisposeImage lBitmap
    End If
    Set ResizePicture = Canvas.Image

End Function

Public Function SavePicture(ByVal Pic As StdPicture, ByRef FileName As String, Optional ByVal Quality As Long = 85) As Boolean
Dim lBitmap As Long
Dim tPicEncoder As GUID
Dim MimeType As String
Dim tParams As EncoderParameters
Const ENCODER_PARAMETER_VALUE_TYPE_LONG As Long = 4
Const ENCODER_QUALITY As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

    If mGdipToken = 0 Then InitGDI
    Select Case LCase$(GetFileExtension(FileName))
        Case "png": MimeType = "image/png"
        Case "bmp": MimeType = "image/bmp"
        Case "gif": MimeType = "image/gif"
        Case "jpg": MimeType = "image/jpeg"
        Case Else
            MimeType = "image/jpeg"
            FileName = FileName & ".jpg"
    End Select
    If MimeType = "image/jpeg" Then
        tParams.Count = 1
        If Quality > 100 Then Quality = 100
        If Quality < 0 Then Quality = 0
        With tParams.Parameter(0)
            CLSIDFromString StrPtr(ENCODER_QUALITY), .EncoderGUID
            .NumberOfValues = 1
            .Type = ENCODER_PARAMETER_VALUE_TYPE_LONG
            .Value = VarPtr(Quality)
        End With
    End If
    
    If GetEncoderClsid(MimeType, tPicEncoder) = False Then
        MsgBox "Es konnte kein passenden Encoder für " & MimeType & "-Dateien gefunden werden.", vbOKOnly, "Encoder Error"
        Exit Function
    End If
    
    If GdipCreateBitmapFromHBITMAP(Pic.Handle, 0, lBitmap) = 0 Then
        If MimeType = "image/jpeg" Then
            SavePicture = CBool(GdipSaveImageToFile(lBitmap, StrPtr(FileName), tPicEncoder, tParams) = 0)
        Else
            SavePicture = CBool(GdipSaveImageToFile(lBitmap, StrPtr(FileName), tPicEncoder, ByVal 0) = 0)
        End If
        Call GdipDisposeImage(lBitmap)
    End If
End Function

Public Sub TerminateGDI()
    If mGdipToken <> 0 Then
        Call GdiplusShutdown(mGdipToken)
        mGdipToken = 0
    End If
End Sub

Private Function ConvertColor(nColor As Long, nOpacity As Single) As Long
Dim bgra(0 To 3) As Byte
Dim iColor As Long
    TranslateColor nColor, 0&, iColor
    bgra(3) = CByte((nOpacity / 100) * 255)
    bgra(0) = ((iColor \ &H10000) And &HFF)
    bgra(1) = ((iColor \ &H100) And &HFF)
    bgra(2) = (iColor And &HFF)
    CopyMemory iColor, bgra(0), 4&
    ConvertColor = iColor
End Function

Private Sub DrawCircle(iGraphics As Long, X As Long, Y As Long, Width As Long, _
                      BorderStyle As BorderStyleConstants, BorderColor As OLE_COLOR, BorderWidth As Integer, _
                      iFilled As Boolean, iFillColor As Long, Opacity As Single)
    If iFilled Then FillEllipse iGraphics, iFillColor, Opacity, X, Y, Width, Width
    If BorderStyle <> vbTransparent Then DrawEllipse iGraphics, BorderStyle, BorderColor, BorderWidth, Opacity, X, Y, Width, Width
    
End Sub

Private Sub DrawEllipse(ByVal nGraphics As Long, ByVal BorderStyle As BorderStyleConstants, ByVal nColor As Long, ByVal nDrawnWidth As Long, Opacity As Single, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, Opacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        If ((BorderStyle > vbBSSolid) And (BorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, BorderStyle - 1)
        End If
        If BorderStyle = vbBSInsideSolid Then
            X = X + nDrawnWidth / 2
            Y = Y + nDrawnWidth / 2
            nWidth = nWidth - nDrawnWidth
            nHeight = nHeight - nDrawnWidth
        End If
        GdipDrawEllipseI nGraphics, hPen, X, Y, nWidth, nHeight
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub DrawLine(iGraphics As Long, x0 As Long, y0 As Long, x1 As Long, y1 As Long, _
                      BorderStyle As BorderStyleConstants, BorderColor As OLE_COLOR, BorderWidth As Integer, Opacity As Single)
Dim hPen As Long
    
    If BorderStyle <> vbTransparent Then
        If GdipCreatePen1(ConvertColor(BorderColor, Opacity), BorderWidth, UnitPixel, hPen) = 0 Then
            If BorderWidth = 1 And Opacity = 100 Then
                Call GdipSetSmoothingMode(iGraphics, SmoothingModeNone)
            Else
                Call GdipSetSmoothingMode(iGraphics, SmoothingModeAntiAlias)
            End If
            GdipDrawLine iGraphics, hPen, x0, y0, x1, y1
        End If
    End If
    
End Sub

Private Sub DrawPolygon(iGraphics As Long, points() As POINTAPI, _
                      BorderStyle As BorderStyleConstants, BorderColor As OLE_COLOR, BorderWidth As Integer, _
                      iFilled As Boolean, iFillColor As Long, Opacity As Single)
Dim hPen As Long
    If iFilled Then FillPolygon iGraphics, iFillColor, Opacity, points
    If BorderStyle <> vbTransparent Then
        If GdipCreatePen1(ConvertColor(BorderColor, Opacity), BorderWidth, UnitPixel, hPen) = 0 Then
            If BorderWidth = 1 And Opacity = 100 Then
                Call GdipSetSmoothingMode(iGraphics, SmoothingModeNone)
            Else
                Call GdipSetSmoothingMode(iGraphics, SmoothingModeAntiAlias)
            End If
            GdipDrawPolygonI iGraphics, hPen, points(0), UBound(points) + 1
        End If
    End If
    
End Sub

Private Sub DrawRectangle(iGraphics As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                      BorderStyle As BorderStyleConstants, BorderColor As OLE_COLOR, BorderWidth As Integer, _
                      iFilled As Boolean, iFillColor As Long, Opacity As Single)
Dim hPen As Long
Dim iPts(3) As POINTAPI
    
    iPts(0).X = X:          iPts(0).Y = Y
    iPts(1).X = X + Width:  iPts(1).Y = Y
    iPts(2).X = X + Width:  iPts(2).Y = Y + Height
    iPts(3).X = X:          iPts(3).Y = Y + Height
    
    If iFilled Then FillPolygon iGraphics, iFillColor, Opacity, iPts
    If BorderStyle <> vbTransparent Then
        If GdipCreatePen1(ConvertColor(BorderColor, Opacity), BorderWidth, UnitPixel, hPen) = 0 Then
            If BorderWidth <= 1 And Opacity = 100 Then
                Call GdipSetSmoothingMode(iGraphics, SmoothingModeNone)
            Else
                Call GdipSetSmoothingMode(iGraphics, SmoothingModeAntiAlias)
            End If
            GdipDrawPolygonI iGraphics, hPen, iPts(0), UBound(iPts) + 1
        End If
    End If
    
End Sub

Private Sub FillEllipse(ByVal nGraphics As Long, ByVal nColor As Long, Opacity As Single, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
Dim hBrush As Long

    If GdipCreateSolidFill(ConvertColor(nColor, Opacity), hBrush) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        GdipFillEllipseI nGraphics, hBrush, X, Y, nWidth, nHeight
        Call GdipDeleteBrush(hBrush)
    End If
End Sub

Private Sub FillPolygon(ByVal nGraphics As Long, ByVal nColor As Long, Opacity As Single, points() As POINTAPI)
Dim hBrush As Long
Const Fill_Mode_Alternate = &H0
Const Fill_Mode_Winding = &H1

    If GdipCreateSolidFill(ConvertColor(nColor, Opacity), hBrush) = 0 Then
        GdipFillPolygonI nGraphics, hBrush, points(0), UBound(points) + 1, Fill_Mode_Alternate
        Call GdipDeleteBrush(hBrush)
    End If
End Sub

Private Function GetEncoderClsid(MimeType As String, pClsid As GUID) As Boolean
Dim n As Long, s As Long, j As Long
Dim pImageCodecInfo() As ImageCodecInfo
Dim Buffer As String
    Call GdipGetImageEncodersSize(n, s)
    If s = 0 Then
        GetEncoderClsid = False
        Exit Function
    End If
    
    ReDim pImageCodecInfo(0 To s \ Len(pImageCodecInfo(0)) - 1)
    Call GdipGetImageEncoders(n, s, pImageCodecInfo(0))
    For j = 0 To n - 1
        Buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
        Call lstrcpyW(ByVal StrPtr(Buffer), ByVal pImageCodecInfo(j).MimeTypePtr)
            
        If (StrComp(Buffer, MimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo
            GetEncoderClsid = True
            Exit Function
        End If
    Next j
    Erase pImageCodecInfo
    GetEncoderClsid = False
End Function

Private Function HandleToPicture(ByVal hGDIHandle As Long, ByVal ObjectType As PictureTypeConstants, Optional ByVal hPal As Long = 0) As StdPicture
Dim tPictDesc As PictDesc
Dim GUID_IPicture As GUID
Dim oPicture As IPicture
    
    If mGdipToken = 0 Then InitGDI

    With tPictDesc
        .cbSizeofStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hPal
    End With
    
    With GUID_IPicture
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
    
    OleCreatePictureIndirect tPictDesc, GUID_IPicture, True, oPicture
    Set HandleToPicture = oPicture
    
End Function


Private Sub InitGDI()
Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(mGdipToken, GdipStartupInput, ByVal 0)
End Sub





