VERSION 5.00
Begin VB.UserControl EviTextBox 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2115
   LockControls    =   -1  'True
   PropertyPages   =   "EviTextBox.ctx":0000
   ScaleHeight     =   465
   ScaleWidth      =   2115
   ToolboxBitmap   =   "EviTextBox.ctx":004F
   Begin VB.TextBox txtBoth 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtVertical 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtHorizontal 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   330
      Left            =   1980
      ScaleHeight     =   270
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   330
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   20
      TabIndex        =   0
      Top             =   20
      Width           =   1935
   End
End
Attribute VB_Name = "EviTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Evi Text Box Style XP"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'by evi indra effendi
'email:effendi24@gmail.com
Option Explicit

Private m_Color       As OLE_COLOR
Private m_hDC         As Long
Private m_hWnd        As Long

Private m_MemDC    As Boolean
Private m_ThDC     As Long
Private m_hBmp     As Long
Private m_hBmpOld  As Long
Private m_lWidth   As Long
Private m_lHeight  As Long
Private TR         As RECT
Private TBR        As RECT

Const ColorXPRec = 16777215

Enum ScrollTextEnum
    [None] = 0
    [Horizontal] = 1
    [Vertical] = 2
    [Both] = 3
End Enum

Enum AlignmentEnumTextBox
    [Left Justify] = 0
    [Right Justify] = 1
    [Center] = 2
End Enum

Dim m_Align As AlignmentEnumTextBox
Dim m_Passchar As String
Dim m_MaxLength As Long
Dim m_Locked As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_Focus As Boolean

Dim ScrollSetting As ScrollTextEnum

Dim sText As String

'Default Property Values:
'Const m_def_MousePointer = 0
Const m_def_ForeColor = vbBlack
Const m_def_HideSelection = 0
Const m_def_Enabled = 1
'Property Variables:
'Dim m_MouseIcon As Picture
'Dim m_MousePointer As Integer
Dim m_Font As Font
Dim m_ForeColor As OLE_COLOR
Dim m_HideSelection As Boolean
Dim m_Enabled As Boolean

Dim m_Point As Long
Dim m_selText As String
Dim m_SelStart As Long
Dim m_SelLength As Long

Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()

Public Sub DoTextBoxStyler()
    GetClientRect m_hWnd, TR
    DrawFillRectangle TR, ColorXPRec, m_hDC
    DrawingObjectTextBox
    If m_MemDC Then
        With UserControl
            pDraw .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
        End With
    End If
End Sub

Private Sub DrawingObjectTextBox()
    DrawRectangle TR, ShiftColorXP(m_Color, 100), m_hDC
    With TBR
        .Left = 1
        .Top = 1
        .Bottom = TR.Bottom - 1
        .Right = TR.Left + (TR.Right - TR.Left) * 0
    End With
    DrawFillRectangle TBR, ShiftColorXP(m_Color, 180), m_hDC
End Sub

Private Function GetLngColor(Color As Long) As Long
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Private Sub DrawRectangle(ByRef bRect As RECT, ByVal Color As Long, ByVal hdc As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(Color)
    FrameRect hdc, bRect, hBrush
    DeleteObject hBrush
End Sub

Public Sub DrawLine( _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)
    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim POS     As POINTAPI
    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    MoveToEx cHdc, X, Y, POS
    LineTo cHdc, Width, Height
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1
End Sub

Private Function ShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long
    Dim R As Long, G As Long, b As Long, Delta As Long
    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    b = ((MyColor \ &H10000) Mod &H100)
    Delta = &HFF - Base
    b = Base + b * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If b > 255 Then b = 255
    ShiftColorXP = R + 256& * G + 65536 * b
End Function

Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(GetLngColor(Color))
    FillRect MyHdc, hRect, hBrush
    DeleteObject hBrush
End Sub

Private Function ThDC(Width As Long, Height As Long) As Long
    If m_ThDC = 0 Then
        If (Width + Height) > 0 Then pCreate Width, Height
    Else
        If Width > m_lWidth Or Height > m_lHeight Then pCreate Width, Height
    End If
    ThDC = m_ThDC
End Function

Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
    Dim lhDCC As Long
    pDestroy
    lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If lhDCC Then
        m_ThDC = CreateCompatibleDC(lhDCC)
        If m_ThDC Then
            m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
            If m_hBmp Then
                m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
                If m_hBmpOld Then
                    m_lWidth = Width
                    m_lHeight = Height
                    DeleteDC lhDCC
                    Exit Sub
                End If
            End If
        End If
        DeleteDC lhDCC
        pDestroy
    End If
End Sub

Public Sub pDraw( _
      ByVal hdc As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
    If WidthSrc <= 0 Then WidthSrc = m_lWidth
    If HeightSrc <= 0 Then HeightSrc = m_lHeight
    BitBlt hdc, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy
End Sub

Private Sub pDestroy()
    If m_hBmpOld Then
        SelectObject m_ThDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    If m_hBmp Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If m_ThDC Then
        DeleteDC m_ThDC
        m_ThDC = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
End Sub

Private Sub Text1_Change()
RaiseEvent Change
sText = Text1.Text
End Sub

Private Sub Text1_Click()
RaiseEvent Click
End Sub

Private Sub Text1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Text1_GotFocus()
If m_Focus = True Then
    FocusText Text1
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub txtBoth_Change()
RaiseEvent Change
sText = txtBoth.Text
End Sub

Private Sub txtBoth_Click()
RaiseEvent Click
End Sub

Private Sub txtBoth_DblClick()
RaiseEvent DblClick
End Sub

Private Sub txtBoth_GotFocus()
If m_Focus = True Then
    FocusText txtBoth
End If
End Sub

Private Sub txtBoth_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtBoth_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtBoth_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtBoth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtBoth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtBoth_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub txtHorizontal_Change()
RaiseEvent Change
sText = txtHorizontal.Text
End Sub

Private Sub txtHorizontal_Click()
RaiseEvent Click
End Sub

Private Sub txtHorizontal_DblClick()
RaiseEvent DblClick
End Sub

Private Sub txtHorizontal_GotFocus()
If m_Focus = True Then
    FocusText txtHorizontal
End If
End Sub

Private Sub txtHorizontal_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtHorizontal_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtHorizontal_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtHorizontal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtHorizontal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtHorizontal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub txtVertical_Change()
RaiseEvent Change
sText = txtVertical.Text
End Sub

Private Sub txtVertical_Click()
RaiseEvent Click
End Sub

Private Sub txtVertical_DblClick()
RaiseEvent DblClick
End Sub

Private Sub txtVertical_GotFocus()
If m_Focus = True Then
    FocusText txtVertical
End If
End Sub

Private Sub txtVertical_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtVertical_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtVertical_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtVertical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    With UserControl
        .BackColor = vbWhite
        .ScaleMode = vbPixels
    End With
    hdc = UserControl.hdc
    hWnd = UserControl.hWnd
    m_Color = GetLngColor(vbHighlight)
    DoTextBoxStyler
End Sub

Private Sub UserControl_InitProperties()
sText = ""
ScrollSetting = None
m_Passchar = ""
m_MaxLength = 0
m_Locked = False
m_Align = 0
m_BackColor = vbWhite
Set m_Font = Ambient.Font
m_ForeColor = m_def_ForeColor
m_HideSelection = m_def_HideSelection
m_Enabled = m_def_Enabled
m_Point = 0
m_Focus = False
RefreshStyler
End Sub

Private Function RefreshStyler()
Text1.Text = sText
txtVertical.Text = sText
txtHorizontal.Text = sText
txtBoth.Text = sText
Text1.PasswordChar = m_Passchar
txtVertical.PasswordChar = m_Passchar
txtHorizontal.PasswordChar = m_Passchar
txtBoth.PasswordChar = m_Passchar
Text1.MaxLength = m_MaxLength
txtVertical.MaxLength = m_MaxLength
txtHorizontal.MaxLength = m_MaxLength
txtBoth.MaxLength = m_MaxLength
Text1.Locked = m_Locked
txtVertical.Locked = m_Locked
txtHorizontal.Locked = m_Locked
txtBoth.Locked = m_Locked
Text1.Alignment = m_Align
txtVertical.Alignment = m_Align
txtHorizontal.Alignment = m_Align
txtBoth.Alignment = m_Align
Text1.BackColor = m_BackColor
txtVertical.BackColor = m_BackColor
txtHorizontal.BackColor = m_BackColor
txtBoth.BackColor = m_BackColor
Text1.Enabled = m_Enabled
txtVertical.Enabled = m_Enabled
txtHorizontal.Enabled = m_Enabled
txtBoth.Enabled = m_Enabled
Set Text1.Font = m_Font
Set txtVertical.Font = m_Font
Set txtHorizontal.Font = m_Font
Set txtBoth.Font = m_Font
Text1.ForeColor = m_ForeColor
txtVertical.ForeColor = m_ForeColor
txtHorizontal.ForeColor = m_ForeColor
txtBoth.ForeColor = m_ForeColor
Text1.MousePointer = m_Point
txtVertical.MousePointer = m_Point
txtHorizontal.MousePointer = m_Point
txtBoth.MousePointer = m_Point
Text1.SelText = m_selText
txtVertical.SelText = m_selText
txtHorizontal.SelText = m_selText
txtBoth.SelText = m_selText
Text1.SelStart = m_SelStart
txtVertical.SelStart = m_SelStart
txtHorizontal.SelStart = m_SelStart
txtBoth.SelStart = m_SelStart
Text1.SelLength = m_SelLength
txtVertical.SelLength = m_SelLength
txtHorizontal.SelLength = m_SelLength
txtBoth.SelLength = m_SelLength
Select Case ScrollSetting
    Case 0:
            txtVertical.Visible = False
            txtHorizontal.Visible = False
            txtBoth.Visible = False
            Text1.Visible = True
    Case 1:
            Text1.Visible = False
            txtHorizontal.Visible = True
            txtBoth.Visible = False
            txtVertical.Visible = False
            txtHorizontal.Top = Text1.Top
            txtHorizontal.Left = Text1.Left
    Case 2:
            Text1.Visible = False
            txtHorizontal.Visible = False
            txtBoth.Visible = False
            txtVertical.Visible = True
            txtVertical.Top = Text1.Top
            txtVertical.Left = Text1.Left
    Case 3:
            Text1.Visible = False
            txtHorizontal.Visible = False
            txtBoth.Visible = True
            txtVertical.Visible = False
            txtBoth.Top = Text1.Top
            txtBoth.Left = Text1.Left
End Select
End Function

Private Sub UserControl_Paint()
    DoTextBoxStyler
End Sub

Private Sub UserControl_Resize()
RaiseEvent Resize
    hdc = UserControl.hdc
    Text1.Width = Picture1.Width - 2
    Text1.Height = Picture2.Height - 2
    txtHorizontal.Width = Text1.Width
    txtVertical.Width = Text1.Width
    txtBoth.Width = Text1.Width
    txtHorizontal.Height = Text1.Height
    txtVertical.Height = Text1.Height
    txtBoth.Height = Text1.Height
    If UserControl.Height < 135 Then
        UserControl.Height = 255
    End If
End Sub

Private Sub UserControl_Show()
RefreshStyler
End Sub

Private Sub UserControl_Terminate()
    pDestroy
End Sub

Private Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Private Property Let hWnd(ByVal chWnd As Long)
    m_hWnd = chWnd
End Property

Private Property Get hdc() As Long
    hdc = m_hDC
End Property

Public Property Get Text() As String
Text = sText
End Property

Public Property Let Text(ByVal New_Text As String)
sText = New_Text
PropertyChanged "Text"
RefreshStyler
End Property

Public Property Get Scrolling() As ScrollTextEnum
Scrolling = ScrollSetting
End Property

Public Property Let Scrolling(ByVal New_Scroll As ScrollTextEnum)
ScrollSetting = New_Scroll
PropertyChanged "Scrolling"
RefreshStyler
End Property

Private Property Let hdc(ByVal cHdc As Long)
    m_hDC = ThDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
    If m_hDC = 0 Then
        m_hDC = UserControl.hdc
    Else
        m_MemDC = True
    End If
End Property

Public Property Get PasswordChar() As String
PasswordChar = m_Passchar
End Property

Public Property Let PasswordChar(ByVal New_Pass As String)
m_Passchar = New_Pass
PropertyChanged "PasswordChar"
RefreshStyler
End Property

Public Property Get MaxLength() As Long
MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal New_Max As Long)
m_MaxLength = New_Max
PropertyChanged "MaxLength"
RefreshStyler
End Property

Public Property Get Locked() As Boolean
Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
m_Locked = New_Locked
PropertyChanged "Locked"
RefreshStyler
End Property

Public Property Get Alignment() As AlignmentEnumTextBox
Alignment = m_Align
End Property

Public Property Let Alignment(ByVal New_Align As AlignmentEnumTextBox)
m_Align = New_Align
PropertyChanged "Alignment"
RefreshStyler
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
m_BackColor = New_BackColor
PropertyChanged "BackColor"
RefreshStyler
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
sText = PropBag.ReadProperty("Text", "")
ScrollSetting = PropBag.ReadProperty("Scrolling", 0)
m_Passchar = PropBag.ReadProperty("PasswordChar", "")
m_MaxLength = PropBag.ReadProperty("MaxLength", 0)
m_Locked = PropBag.ReadProperty("Locked", False)
m_Align = PropBag.ReadProperty("Alignment", 0)
m_BackColor = PropBag.ReadProperty("BackColor", vbWhite)
Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
m_HideSelection = PropBag.ReadProperty("HideSelection", m_def_HideSelection)
m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
m_Point = PropBag.ReadProperty("MousePointer", 0)
Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
m_SelLength = PropBag.ReadProperty("SelLength", 0)
m_SelStart = PropBag.ReadProperty("SelStart", 0)
m_selText = PropBag.ReadProperty("SelText", "")
m_Focus = PropBag.ReadProperty("Focus", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Text", sText, "")
Call PropBag.WriteProperty("Scrolling", ScrollSetting, 0)
Call PropBag.WriteProperty("PasswordChar", m_Passchar, "")
Call PropBag.WriteProperty("MaxLength", m_MaxLength, 0)
Call PropBag.WriteProperty("Locked", m_Locked, False)
Call PropBag.WriteProperty("Alignment", m_Align, 0)
Call PropBag.WriteProperty("BackColor", m_BackColor, vbWhite)
Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
Call PropBag.WriteProperty("HideSelection", m_HideSelection, m_def_HideSelection)
Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
Call PropBag.WriteProperty("MousePointer", m_Point, 0)
Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
Call PropBag.WriteProperty("SelLength", m_SelLength, 0)
Call PropBag.WriteProperty("SelStart", m_SelStart, 0)
Call PropBag.WriteProperty("SelText", m_selText, "")
Call PropBag.WriteProperty("Focus", m_Focus, False)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    RefreshStyler
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    RefreshStyler
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    m_HideSelection = New_HideSelection
    PropertyChanged "HideSelection"
    RefreshStyler
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    RefreshStyler
End Property

Public Property Get DataFormat() As IStdDataFormatDisp
    Set DataFormat = Text1.DataFormat
End Property

Public Property Set DataFormat(ByVal New_DataFormat As IStdDataFormatDisp)
    Set Text1.DataFormat = New_DataFormat
    Set txtVertical.DataFormat = New_DataFormat
    Set txtHorizontal.DataFormat = New_DataFormat
    Set txtBoth.DataFormat = New_DataFormat
    PropertyChanged "DataFormat"
End Property

Public Property Get MousePointer() As Long
    MousePointer = m_Point
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Long)
    m_Point = New_MousePointer
    PropertyChanged "MousePointer"
    RefreshStyler
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Text1.MouseIcon = New_MouseIcon
    Set txtVertical.MouseIcon = New_MouseIcon
    Set txtHorizontal.MouseIcon = New_MouseIcon
    Set txtBoth.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = m_SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    m_SelLength = New_SelLength
    PropertyChanged "SelLength"
    RefreshStyler
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = m_SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    m_SelStart = New_SelStart
    PropertyChanged "SelStart"
    RefreshStyler
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelText
Public Property Get SelText() As String
    SelText = m_selText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    m_selText = New_SelText
    PropertyChanged "SelText"
    RefreshStyler
End Property

Private Function FocusText(ByVal Text As TextBox)
With Text
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Function

Public Property Get Focus() As Boolean
Focus = m_Focus
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)
m_Focus = New_Focus
PropertyChanged "Focus"
End Property
