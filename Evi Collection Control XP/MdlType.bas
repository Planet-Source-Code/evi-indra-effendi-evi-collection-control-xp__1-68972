Attribute VB_Name = "MdlType"
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left      As Long
    Top       As Long
    Right     As Long
    Bottom    As Long
End Type

Public Type RGB
    Red As Double
    Green As Double
    Blue As Double
End Type

Public Type textparametreleri
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Type ICCEX
    dwSize As Long
    dwICC As Long
End Type

Public Type ControlType
    cntrlObjectForm As Object
    cntrlHwnd As Long
    cntrlToolTipsText As String
    cntrlToolTipsTitle As String
    cntrlToolTipsIcon As Integer
End Type

Public Type TOOLINFO
    cbSize As Long
    dwFlags As Long
    hWnd As Long
    dwID As Long
    rtRect As RECT
    hInst As Long
    lpszText As Long
    lParam  As Long
End Type

Public Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    ttiIcon As Long
End Type
