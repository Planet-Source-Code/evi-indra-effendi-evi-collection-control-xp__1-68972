VERSION 5.00
Begin VB.UserControl EviFrame 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ControlContainer=   -1  'True
   PropertyPages   =   "EviFrame.ctx":0000
   ScaleHeight     =   2460
   ScaleWidth      =   3630
   ToolboxBitmap   =   "EviFrame.ctx":003D
   Begin VB.PictureBox mTitle 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   242
      TabIndex        =   0
      Top             =   0
      Width           =   3630
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   75
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   60
         Picture         =   "EviFrame.ctx":034F
         Stretch         =   -1  'True
         Top             =   75
         Width           =   240
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F7DFD6&
      Height          =   2040
      Left            =   0
      Top             =   375
      Width           =   3615
   End
End
Attribute VB_Name = "EviFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'by evi indra effendi
'email:effendi24@gmail.com
Option Explicit

Enum OrientationFrameEnum
    [Top Orientation] = &H0
    [Bottom Orientation] = &H1
    [Left Orientation] = &H2
    [Right Orientation] = &H3
End Enum

Dim m_Orientation As OrientationFrameEnum
Dim m_BackColor As OLE_COLOR
Dim m_ColorGradient1 As OLE_COLOR
Dim m_ColorGradient2 As OLE_COLOR
Dim m_BorderColor As OLE_COLOR

Dim m_IconShow As Boolean
Dim m_Caption As String

Enum ThemeFrameEnum
    [Blue Theme] = 0
    [Olive Theme] = 1
    [Silver Theme] = 2
    [Royal Theme] = 3
    [Black Theme] = 4
    [Red Theme] = 5
End Enum

Enum BackStyleFrameEnum
    [Transparent] = 0
    [Opaque] = 1
End Enum

Dim m_Theme As ThemeFrameEnum

Dim GetHeight As Long
Dim GetWidth As Long

Dim m_MouseIcon As Picture
Dim m_MousePointer As MousePointerConstants

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()

Public Property Get Orientation() As OrientationFrameEnum
Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationFrameEnum)
m_Orientation = New_Orientation
PropertyChanged "Orientation"
DrawGradientTitleFrame
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
m_BackColor = New_BackColor
PropertyChanged "BackColor"
RefreshEviFrame
End Property

Public Property Get ColorGradient1() As OLE_COLOR
ColorGradient1 = m_ColorGradient1
End Property

Public Property Let ColorGradient1(ByVal New_ColorGradient1 As OLE_COLOR)
m_ColorGradient1 = New_ColorGradient1
PropertyChanged "ColorGradient1"
RefreshEviFrame
End Property

Public Property Get ColorGradient2() As OLE_COLOR
ColorGradient2 = m_ColorGradient2
End Property

Public Property Let ColorGradient2(ByVal New_ColorGradient2 As OLE_COLOR)
m_ColorGradient2 = New_ColorGradient2
PropertyChanged "ColorGradient2"
RefreshEviFrame
End Property

Public Property Get BorderColor() As OLE_COLOR
BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
m_BorderColor = New_BorderColor
PropertyChanged "BorderColor"
RefreshEviFrame
End Property

Public Property Get ShowIcon() As Boolean
ShowIcon = m_IconShow
End Property

Public Property Let ShowIcon(ByVal New_Show As Boolean)
m_IconShow = New_Show
PropertyChanged "ShowIcon"
RefreshEviFrame
End Property

Public Property Get Caption() As String
Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
m_Caption = New_Caption
PropertyChanged "Caption"
RefreshEviFrame
End Property

Private Sub RefreshEviFrame()
DrawGradientTitleFrame
UserControl.BackColor = m_BackColor
Shape1.BorderColor = m_BorderColor
Image1.Height = 16
Image1.Width = 16
Image1.Visible = m_IconShow
If m_IconShow = True Then
    Label1.Left = 24
Else
    Label1.Left = 8
End If
Label1.Caption = m_Caption
Set Label2.MouseIcon = m_MouseIcon
Set UserControl.MouseIcon = m_MouseIcon
Label2.MousePointer = m_MousePointer
UserControl.MousePointer = m_MousePointer
End Sub

Private Sub DrawGradientTitleFrame()
DrawGradient mTitle, m_Orientation, mTitle.ScaleHeight, mTitle.ScaleWidth
End Sub

Private Sub DrawGradient(Optional ObjDraw As Object, Optional _
NewOrientation As OrientationFrameEnum, Optional SH As Long, Optional _
SW As Long)
Dim VR, VG, VB As Single
Dim Color1, Color2 As Long
Dim R, G, b, R2, G2, X, Y, B2 As Integer
Dim temp As Long
Dim m_Position, m_Right, m_Left As Long

m_Position = 0
m_Right = 0
m_Left = 0

Color1 = m_ColorGradient1
Color2 = m_ColorGradient2

temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
b = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
B2 = temp And 255

If NewOrientation = [Top Orientation] Then

VR = Abs(R - R2) / SH
VG = Abs(G - G2) / SH
VB = Abs(b - B2) / SH

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For Y = 0 To SH
R2 = R + VR * Y
G2 = G + VG * Y
B2 = b + VB * Y

ObjDraw.Line (0, Y)-(SW, Y), RGB(R2, G2, B2)
Next Y

ElseIf NewOrientation = [Bottom Orientation] Then
m_Position = SH / 30
m_Left = SH - m_Position
m_Right = m_Left + m_Position

VR = Abs(R - R2) / SH
VG = Abs(G - G2) / SH
VB = Abs(b - B2) / SH

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For Y = 0 To SH
R2 = R + VR * Y
G2 = G + VG * Y
B2 = b + VB * Y

ObjDraw.Line (0, m_Left)-(SW, m_Right), RGB(R2, G2, B2), BF
m_Left = m_Left - m_Position
m_Right = m_Left + m_Position
Next Y

ElseIf NewOrientation = [Left Orientation] Then

VR = Abs(R - R2) / SW
VG = Abs(G - G2) / SW
VB = Abs(b - B2) / SW

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To SW
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDraw.Line (X, 0)-(X, SH), RGB(R2, G2, B2)
Next X

ElseIf NewOrientation = [Right Orientation] Then

m_Position = SW / 200
m_Left = SW - m_Position
m_Right = m_Left + m_Position

VR = Abs(R - R2) / SW
VG = Abs(G - G2) / SW
VB = Abs(b - B2) / SW

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To SW
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDraw.Line (m_Left, 0)-(m_Right, SH), RGB(R2, G2, B2)

m_Left = m_Left - m_Position
m_Right = m_Left + m_Position

Next X
End If
End Sub

Private Sub Label2_Click()
RaiseEvent Click
End Sub

Private Sub Label2_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
m_Orientation = [Left Orientation]
m_BackColor = &HF7DFD6
m_ColorGradient1 = vbWhite
m_ColorGradient2 = &HF7D3C6
m_BorderColor = &HF7DFD6
m_IconShow = True
m_Caption = Ambient.DisplayName
m_Theme = [Blue Theme]
Set m_MouseIcon = Nothing
m_MousePointer = 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
RefreshEviFrame
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_Orientation = PropBag.ReadProperty("Orientation", 1)
m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
m_ColorGradient1 = PropBag.ReadProperty("ColorGradient1", vbWhite)
m_ColorGradient2 = PropBag.ReadProperty("ColorGradient2", &HF7DFD6)
m_BorderColor = PropBag.ReadProperty("BorderColor", &HF7DFD6)
m_IconShow = PropBag.ReadProperty("ShowIcon", True)
m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
Set Icon = PropBag.ReadProperty("Icon", Nothing)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
m_MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Resize()
GetHeight = UserControl.Height
GetWidth = UserControl.Width

Shape1.Width = GetWidth - 20
Shape1.Height = GetHeight - 380
Label1.Width = GetWidth
RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
RefreshEviFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Orientation", m_Orientation, 1)
Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
Call PropBag.WriteProperty("ColorGradient1", m_ColorGradient1, vbWhite)
Call PropBag.WriteProperty("ColorGradient2", m_ColorGradient2, &HF7DFD6)
Call PropBag.WriteProperty("BorderColor", m_BorderColor, &HF7DFD6)
Call PropBag.WriteProperty("ShowIcon", m_IconShow, True)
Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
Call PropBag.WriteProperty("Icon", Icon, Nothing)
Call PropBag.WriteProperty("ForeColor", ForeColor, vbBlack)
Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
Call PropBag.WriteProperty("MousePointer", m_MousePointer, 0)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,MouseIcon
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Sets a custom mouse icon."
    Set Icon = Image1.Picture
End Property

Public Property Set Icon(ByVal New_MouseIcon As Picture)
    Set Image1.Picture = New_MouseIcon
    PropertyChanged "Icon"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Label1.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"
End Property

Public Property Get Theme() As ThemeFrameEnum
Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As ThemeFrameEnum)
m_Theme = New_Theme
PropertyChanged "Theme"
Select Case Theme
    Case 0:
            m_ColorGradient1 = vbWhite
            m_ColorGradient2 = &HF7D3C6
            m_BorderColor = &HF7DFD6
            m_BackColor = &HF7DFD6
            m_Orientation = [Left Orientation]
            ForeColor = vbBlack
    Case 1:
            m_ColorGradient1 = vbWhite
            m_ColorGradient2 = &HB8E7E0
            m_BorderColor = &HECF6F6
            m_BackColor = &HECF6F6
            m_Orientation = [Left Orientation]
            ForeColor = vbBlack
    Case 2:
            m_ColorGradient1 = vbWhite
            m_ColorGradient2 = &HE0D7D6
            m_BorderColor = &HF5F1F0
            m_BackColor = &HF5F1F0
            m_Orientation = [Left Orientation]
            ForeColor = vbBlack
    Case 3:
            m_ColorGradient1 = vbBlue
            m_ColorGradient2 = &H808000
            m_BorderColor = &HB75F31
            m_BackColor = &HFFFFFF
            m_Orientation = [Top Orientation]
            ForeColor = &H8000000F
    Case 4:
            m_ColorGradient1 = &HC0C0C0
            m_ColorGradient2 = vbBlack
            m_BorderColor = vbBlack
            m_BackColor = &HE0E0E0
            m_Orientation = [Top Orientation]
            ForeColor = vbWhite
    Case 5:
            m_ColorGradient1 = &H8080FF
            m_ColorGradient2 = &H80&
            m_BorderColor = &H80&
            m_BackColor = &HC0C0FF
            m_Orientation = [Top Orientation]
            ForeColor = vbWhite
End Select
RefreshEviFrame
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackStyleFrameEnum
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleFrameEnum)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set m_MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
    RefreshEviFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
    RefreshEviFrame
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

