VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'by evi indra effendi
'email:effendi24@gmail.com
Option Explicit

Enum GradientPositionEnum
    [Top Gradient] = &H0
    [Bottom Gradient] = &H1
    [Left Gradient] = &H2
    [Right Gradient] = &H3
    [Spin Gradient] = &H4
    [Box Gradient] = &H5
    [Diagonal Gradient] = &H6
    [Rectangular Gradient] = &H7
    [Circle Gradient] = &H8
End Enum

Public Sub DrawGradient(Optional ObjDrawGradient As Object, Optional _
NewGradient As GradientPositionEnum, Optional Color1 As OLE_COLOR, Optional _
Color2 As OLE_COLOR)
Dim VR, VG, VB As Single
Dim R, G, b, R2, G2, X, Y, B2 As Integer
Dim temp As Long
Dim m_Position, m_Right, m_Left As Long

ObjDrawGradient.Cls
ObjDrawGradient.AutoRedraw = True
ObjDrawGradient.DrawWidth = 5
ObjDrawGradient.ScaleMode = vbPixels

m_Position = 0
m_Right = 0
m_Left = 0

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

If NewGradient = [Top Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleHeight
VG = Abs(G - G2) / ObjDrawGradient.ScaleHeight
VB = Abs(b - B2) / ObjDrawGradient.ScaleHeight

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For Y = 0 To ObjDrawGradient.ScaleHeight
R2 = R + VR * Y
G2 = G + VG * Y
B2 = b + VB * Y

ObjDrawGradient.Line (0, Y)-(ObjDrawGradient.ScaleWidth, Y), RGB(R2, G2, B2)
Next Y

ElseIf NewGradient = [Bottom Gradient] Then
m_Position = ObjDrawGradient.ScaleHeight / 30
m_Left = ObjDrawGradient.ScaleHeight - m_Position
m_Right = m_Left + m_Position

VR = Abs(R - R2) / ObjDrawGradient.ScaleHeight
VG = Abs(G - G2) / ObjDrawGradient.ScaleHeight
VB = Abs(b - B2) / ObjDrawGradient.ScaleHeight

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For Y = 0 To ObjDrawGradient.ScaleHeight
R2 = R + VR * Y
G2 = G + VG * Y
B2 = b + VB * Y

ObjDrawGradient.Line (0, m_Left)-(ObjDrawGradient.ScaleWidth, m_Right), RGB(R2, G2, B2), BF
m_Left = m_Left - m_Position
m_Right = m_Left + m_Position
Next Y

ElseIf NewGradient = [Left Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To ObjDrawGradient.ScaleWidth
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Line (X, 0)-(X, ObjDrawGradient.ScaleHeight), RGB(R2, G2, B2)
Next X

ElseIf NewGradient = [Right Gradient] Then

m_Position = ObjDrawGradient.ScaleWidth / 200
m_Left = ObjDrawGradient.ScaleWidth - m_Position
m_Right = m_Left + m_Position

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To ObjDrawGradient.ScaleWidth
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Line (m_Left, 0)-(m_Right, ObjDrawGradient.ScaleHeight), RGB(R2, G2, B2)

m_Left = m_Left - m_Position
m_Right = m_Left + m_Position

Next X

ElseIf NewGradient = [Spin Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth / 2
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth / 2
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth / 2

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To ObjDrawGradient.ScaleWidth
R2 = R2 - VR '* X
G2 = G2 - VG '* X
B2 = B2 - VB '* X
ObjDrawGradient.Line (X, 0)-(ObjDrawGradient.ScaleWidth - X, ObjDrawGradient.ScaleHeight), RGB(R2, G2, B2)
Next X

For X = 0 To ObjDrawGradient.ScaleHeight
R2 = R2 - VR '* X
G2 = G2 - VG '* X
B2 = B2 - VB '* X
ObjDrawGradient.Line (ObjDrawGradient.ScaleWidth, X)-(0, ObjDrawGradient.ScaleHeight - X), RGB(R2, G2, B2)
Next X

ElseIf NewGradient = [Box Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = ObjDrawGradient.ScaleWidth To 0 Step -1
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Line ((X / 5), (X / 5))-(ObjDrawGradient.ScaleWidth - (X / 5), ObjDrawGradient.ScaleHeight - (X / 5)), RGB(R2, G2, B2), B

Next X

ElseIf NewGradient = [Diagonal Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To ObjDrawGradient.ScaleWidth
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Line (0, X)-(X, 0), RGB(R2, G2, B2)
Next X

For X = ObjDrawGradient.ScaleWidth To 0 Step -1
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Line (ObjDrawGradient.ScaleWidth - X, ObjDrawGradient.ScaleWidth)-(ObjDrawGradient.ScaleWidth, ObjDrawGradient.ScaleWidth - X), RGB(R2, G2, B2)
Next X

ElseIf NewGradient = [Rectangular Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = ObjDrawGradient.ScaleWidth To 0 Step -1
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Line ((X / 2), (X / 2))-(ObjDrawGradient.ScaleWidth - (X / 2), ObjDrawGradient.ScaleHeight - (X / 2)), RGB(R2, G2, B2), B
Next X

ElseIf NewGradient = [Circle Gradient] Then

VR = Abs(R - R2) / ObjDrawGradient.ScaleWidth
VG = Abs(G - G2) / ObjDrawGradient.ScaleWidth
VB = Abs(b - B2) / ObjDrawGradient.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < b Then VB = -VB

For X = 0 To ObjDrawGradient.ScaleWidth
R2 = R + VR * X
G2 = G + VG * X
B2 = b + VB * X

ObjDrawGradient.Circle (ObjDrawGradient.ScaleWidth / 2, ObjDrawGradient.ScaleHeight / 2), X, RGB(R2, G2, B2)
Next X
End If
End Sub

Public Sub Animation(Optional ObjForAnimation As Object, Optional _
NewBackColor As OLE_COLOR, Optional CounterLine As Long)
    Dim X As Long
    Dim X1, X2, Y1, Y2, m_colo As Long
    ObjForAnimation.Cls
    ObjForAnimation.BackColor = NewBackColor
    ObjForAnimation.ScaleHeight = 100
    ObjForAnimation.ScaleWidth = 100
    For X = 0 To CounterLine
        DoEvents
        X1 = Int(Rnd * 101)
        X2 = Int(Rnd * 101)
        Y1 = Int(Rnd * 101)
        Y2 = Int(Rnd * 101)
        m_colo = Int(Rnd * 15)
        ObjForAnimation.Line (X1, Y1)-(X2, Y2), QBColor(m_colo)
        ObjForAnimation.Line (X1, Y2)-(X2, Y1), QBColor(m_colo)
        ObjForAnimation.Line (X2, Y1)-(X1, Y2), QBColor(m_colo)
        ObjForAnimation.Line (Y1, Y2)-(X1, X2), QBColor(m_colo)
    Next X
End Sub
