VERSION 5.00
Begin VB.UserControl EviAnimationForm 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   Picture         =   "EviAnimationForm.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   450
   ToolboxBitmap   =   "EviAnimationForm.ctx":084E
End
Attribute VB_Name = "EviAnimationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'by evi indra effendi
'email:effendi24@gmail.com
Option Explicit

Dim m_ControlType() As ControlType
Dim m_Counter As Long

Public GraphicForm As New GraphicForms

Enum EditTipIcon
    etiNone = 0
    etiInfo = 1
    etiWarning = 2
    etiError = 3
End Enum

Const HWND_TOPMOST As Long = -1
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Long = &H1

Const ICC_WIN95_CLASSES As Long = &HFF

Const CCM_FIRST As Long = &H2000
Const CCM_SETWINDOWTHEME As Long = (CCM_FIRST + &HB)
Const WM_USER As Long = &H400
Const CW_USEDEFAULT As Long = &H80000000
Const ECM_FIRST As Long = &H1500

Const EM_SHOWBALLOONTIP = ECM_FIRST + 3

Const WS_POPUP As Long = &H80000000
Const WS_EX_TOPMOST As Long = &H8&

Const TOOLTIPS_CLASSA As String = "tooltips_class32"

Const TTF_ABSOLUTE As Long = &H80
Const TTF_CENTERTIP As Long = &H2
Const TTF_DI_SETITEM As Long = &H8000
Const TTF_IDISHWND As Long = &H1
Const TTF_RTLREADING As Long = &H4
Const TTF_SUBCLASS As Long = &H10
Const TTF_TRACK As Long = &H20
Const TTF_TRANSPARENT As Long = &H100

Const TTI_ERROR As Long = 3
Const TTI_INFO As Long = 1
Const TTI_NONE As Long = 0
Const TTI_WARNING As Long = 2

Const TTM_ACTIVATE As Long = (WM_USER + 1)
Const TTM_ADDTOOL As Long = (WM_USER + 4)
Const TTM_ADJUSTRECT As Long = (WM_USER + 31)
Const TTM_DELTOOL As Long = (WM_USER + 5)
Const TTM_ENUMTOOLS As Long = (WM_USER + 14)
Const TTM_GETBUBBLESIZE As Long = (WM_USER + 30)
Const TTM_GETCURRENTTOOL As Long = (WM_USER + 15)
Const TTM_GETDELAYTIME As Long = (WM_USER + 21)
Const TTM_GETMARGIN As Long = (WM_USER + 27)
Const TTM_GETMAXTIPWIDTH As Long = (WM_USER + 25)
Const TTM_GETTEXT As Long = (WM_USER + 11)
Const TTM_GETTIPBKCOLOR As Long = (WM_USER + 22)
Const TTM_GETTIPTEXTCOLOR As Long = (WM_USER + 23)
Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
Const TTM_GETTOOLINFO As Long = (WM_USER + 8)
Const TTM_HITTEST As Long = (WM_USER + 10)
Const TTM_NEWTOOLRECT As Long = (WM_USER + 6)
Const TTM_POP As Long = (WM_USER + 28)
Const TTM_POPUP As Long = (WM_USER + 34)
Const TTM_RELAYEVENT As Long = (WM_USER + 7)
Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Const TTM_SETMARGIN As Long = (WM_USER + 26)
Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Const TTM_SETTITLE As Long = (WM_USER + 32)
Const TTM_SETTOOLINFO As Long = (WM_USER + 9)
Const TTM_SETWINDOWTHEME As Long = CCM_SETWINDOWTHEME
Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Const TTM_TRACKPOSITION As Long = (WM_USER + 18)
Const TTM_UPDATE As Long = (WM_USER + 29)
Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Const TTM_WINDOWFROMPOINT As Long = (WM_USER + 16)

Const TTN_FIRST As Long = (-520)
Const TTN_GETDISPINFO As Long = (TTN_FIRST - 0)
Const TTN_LAST As Long = (-549)
Const TTN_LINKCLICK As Long = (TTN_FIRST - 3)
Const TTN_NEEDTEXT As Long = TTN_GETDISPINFO
Const TTN_POP As Long = (TTN_FIRST - 2)
Const TTN_SHOW As Long = (TTN_FIRST - 1)

Const TTS_ALWAYSTIP As Long = &H1
Const TTS_BALLOON As Long = &H40
Const TTS_NOANIMATE As Long = &H10
Const TTS_NOFADE As Long = &H20
Const TTS_NOPREFIX As Long = &H2

Private ghWndTip As Long, ghWndParent As Long

Dim m_Object As Object

Enum ttIconType
  [No Icon] = 0
  [Icon Info] = 1
  [Icon Warning] = 2
  [Icon Error] = 3
End Enum

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set m_Object = UserControl.Parent
End Sub

Private Sub UserControl_Resize()
If UserControl.Width <> 450 Then
    UserControl.Width = 450
End If
If UserControl.Height <> 495 Then
    UserControl.Height = 495
End If
End Sub

Private Sub UserControl_Show()
Set m_Object = UserControl.Parent
End Sub

Public Sub Show()
Dim m_m As Long
If m_Counter <= 0 Then Exit Sub
For m_m = 1 To m_Counter
    ShowToolTipsBalloon m_ControlType(m_m).cntrlObjectForm, m_ControlType(m_m).cntrlHwnd, m_ControlType(m_m).cntrlToolTipsText, m_ControlType(m_m).cntrlToolTipsTitle, m_ControlType(m_m).cntrlToolTipsIcon
Next m_m
End Sub

Public Function Add(Optional ObjectFormOwner As Object = Nothing, Optional AddObjectToShowToolTips As Object = Nothing, Optional ToolTipText As String _
= "", Optional ToolTipTitle As String = "", Optional _
ToolTipIcon As ttIconType = 1)
m_Counter = m_Counter + 1
ReDim Preserve m_ControlType(m_Counter)
Set m_ControlType(m_Counter).cntrlObjectForm = ObjectFormOwner
m_ControlType(m_Counter).cntrlHwnd = AddObjectToShowToolTips.hWnd
m_ControlType(m_Counter).cntrlToolTipsText = ToolTipText
m_ControlType(m_Counter).cntrlToolTipsTitle = ToolTipTitle
m_ControlType(m_Counter).cntrlToolTipsIcon = ToolTipIcon
End Function

Private Sub ShowToolTipsBalloon(Optional ObjectForm As Object, Optional OwnHwnd As Long, Optional _
ToolTipsText As String, Optional ToolTipTitle As String, Optional _
ToolTipIcon As Integer)
    Dim tiInfo As TOOLINFO
    Dim MyHwnD As Long
    Dim hWndTip As Long, dwFlags As Long, ICEx As ICCEX
    
    dwFlags = TTS_NOPREFIX Or TTS_ALWAYSTIP Or TTS_BALLOON
    
    With ICEx
        .dwSize = Len(ICEx)
        .dwICC = ICC_WIN95_CLASSES
    End With
    
    InitCommonControlsEx ICEx
    
    hWndTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, "", WS_POPUP Or dwFlags, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, OwnHwnd, 0, App.hInstance, ByVal 0&)
    
    If hWndTip = 0 Then Exit Sub
    
    SetWindowPos hWndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ghWndTip = hWndTip
    ghWndParent = ObjectForm.hWnd
    
    With tiInfo
        .dwFlags = TTF_SUBCLASS Or TTF_TRANSPARENT
        .hWnd = OwnHwnd
        .lpszText = StrPtr(StrConv(ToolTipsText, vbFromUnicode))
        .hInst = App.hInstance
        GetClientRect OwnHwnd, .rtRect
        
        .cbSize = Len(tiInfo)

    End With
    
    SendMessage ghWndTip, TTM_ADDTOOL, 0&, tiInfo
    If ToolTipTitle <> vbNullString Or ToolTipIcon <> 0 Then
        SendMessage ghWndTip, TTM_SETTITLE, CLng(ToolTipIcon), ByVal ToolTipTitle
    End If
End Sub

Private Sub UserControl_Terminate()
m_Counter = 0
ReDim Preserve m_ControlType(m_Counter)
End Sub
