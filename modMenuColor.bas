Attribute VB_Name = "modMenuColor"
Option Explicit
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden wird nicht gehaftet.
'Geschrieben von Wolfgang Ehrhardt woeh@gmx.de
   
'Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCountA Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, Mi As MENUINFO) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" ( _
        ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
        ByVal hBitmapChecked As Long) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" _
    (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long


Public Type MENUINFO
    cbSize          As Long
    fMask           As Long
    dwStyle         As Long
    cyMax           As Long
    hbrBack         As Long
    dwContextHelpID As Long
    dwMenuData      As Long
End Type

Public Enum MenuNFO
    nfoMenuBarColor = 1
    nfoMenuColor = 2
    nfoSysMenuColor = 3
End Enum

Private Const MIM_BACKGROUND As Long = &H2&
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Public Function Convert_OLEtoRBG(ByVal OLEcolor As Long) As Long
Const CLR_INVALID = -1
    If OleTranslateColor(OLEcolor, 0, Convert_OLEtoRBG) Then Convert_OLEtoRBG = CLR_INVALID
End Function


Public Function Set_MenuColor(SetWhat As MenuNFO, _
    ByVal hwnd As Long, ByVal Color As Long, _
    Optional MenuIndex As Integer = -1, _
    Optional IncludeSubmenus As Boolean = False) As Boolean
    
Dim Mi As MENUINFO
Dim clrref As Long, hSysMenu As Long, mHwnd As Long
         
    On Error GoTo Set_MenuColor_Error
    clrref = Convert_OLEtoRBG(Color)
    Mi.cbSize = Len(Mi)
    Mi.hbrBack = CreateSolidBrush(clrref)
    
    Select Case SetWhat
        Case nfoMenuBarColor
            Mi.fMask = MIM_BACKGROUND
            Call SetMenuInfo(GetMenu(hwnd), Mi)
            
        Case nfoMenuColor
            If MenuIndex = -1 Then
                Set_MenuColor = Set_MenuColor(nfoMenuBarColor, hwnd, Color)
                Exit Function
            End If
            
            If Get_MenuItemCount(hwnd) < MenuIndex Then Exit Function
            Mi.fMask = IIf(IncludeSubmenus, MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS, MIM_BACKGROUND)
            mHwnd = GetMenu(hwnd)
            mHwnd = GetSubMenu(mHwnd, MenuIndex)
            Call SetMenuInfo(mHwnd, Mi)
            hwnd = mHwnd
            
        Case nfoSysMenuColor
            hSysMenu = GetSystemMenu(hwnd, False)
            Mi.fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
            Call SetMenuInfo(hSysMenu, Mi)
            hwnd = hSysMenu
    End Select
    
    Call DrawMenuBar(hwnd)
    Set_MenuColor = True
Exit Function

Set_MenuColor_Error:

End Function

Private Function Get_MenuHwnd(ByVal hwnd As Long) As Long
    Get_MenuHwnd = GetMenu(hwnd)
End Function

Private Function Get_MenuItemCount(ByVal hwnd As Long) As Long
    Get_MenuItemCount = GetMenuItemCountA(Get_MenuHwnd(hwnd))
End Function

