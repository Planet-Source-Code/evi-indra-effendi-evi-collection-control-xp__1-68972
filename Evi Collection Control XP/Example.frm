VERSION 5.00
Object = "*\AprjEviCollectionControlXP.vbp"
Begin VB.Form Example 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "Example.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CollectionControlXP.EviButtons EviButtons10 
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Top             =   5880
      Width           =   1455
      _extentx        =   2566
      _extenty        =   661
      caption         =   "Close"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "Example.frx":000C
      font            =   "Example.frx":0038
      mousepointer    =   99
   End
   Begin CollectionControlXP.EviButtons EviButtons9 
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Top             =   5880
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      caption         =   "Test Gradient"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "Example.frx":0064
      font            =   "Example.frx":0090
      mousepointer    =   99
   End
   Begin CollectionControlXP.EviButtons EviButtons8 
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      caption         =   "Test Animation"
      buttonstyle     =   3
      originalpicsizew=   0
      originalpicsizeh=   0
      font            =   "Example.frx":00BC
      font            =   "Example.frx":00E8
      mousepointer    =   99
   End
   Begin CollectionControlXP.EviAnimationForm EviAnimationForm1 
      Left            =   120
      Top             =   4560
      _extentx        =   794
      _extenty        =   873
   End
   Begin CollectionControlXP.EviFrame EviFrame7 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   7095
      _extentx        =   12515
      _extenty        =   2778
      orientation     =   0
      backcolor       =   14737632
      colorgradient1  =   12632256
      colorgradient2  =   0
      bordercolor     =   0
      caption         =   "Frame With Black Theme Using BackStyle Transparent"
      icon            =   "Example.frx":0114
      forecolor       =   16777215
      font            =   "Example.frx":04B0
      backstyle       =   0
   End
   Begin CollectionControlXP.EviFrame EviFrame6 
      Height          =   1815
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3201
      orientation     =   0
      backcolor       =   12632319
      colorgradient1  =   8421631
      colorgradient2  =   128
      bordercolor     =   128
      caption         =   "Frame With Red Theme"
      icon            =   "Example.frx":04DC
      forecolor       =   16777215
      font            =   "Example.frx":0878
      Begin CollectionControlXP.EviButtons EviButtons7 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Test Progress Bar"
         buttonstyle     =   3
         originalpicsizew=   0
         originalpicsizeh=   0
         font            =   "Example.frx":08A4
         font            =   "Example.frx":08D0
         mousepointer    =   99
      End
      Begin CollectionControlXP.EviProgressBar EviProgressBar3 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         font            =   "Example.frx":08FC
         brushstyle      =   0
         color           =   12937777
         color2          =   12937777
         style           =   4
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pastel Style"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
   End
   Begin CollectionControlXP.EviFrame EviFrame5 
      Height          =   1815
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3201
      orientation     =   0
      backcolor       =   14737632
      colorgradient1  =   12632256
      colorgradient2  =   0
      bordercolor     =   0
      caption         =   "Frame With Black Theme"
      icon            =   "Example.frx":0928
      forecolor       =   16777215
      font            =   "Example.frx":0CC4
      Begin CollectionControlXP.EviButtons EviButtons6 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Test Progress Bar"
         buttonstyle     =   3
         icon            =   "Example.frx":0CF0
         iconwidth       =   16
         iconheight      =   16
         iconsize        =   0
         originalpicsizew=   16
         originalpicsizeh=   16
         font            =   "Example.frx":108C
         font            =   "Example.frx":10B8
         mousepointer    =   99
      End
      Begin CollectionControlXP.EviProgressBar EviProgressBar2 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         font            =   "Example.frx":10E4
         brushstyle      =   0
         color           =   12937777
         color2          =   12937777
         style           =   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Style"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2055
      End
   End
   Begin CollectionControlXP.EviFrame EviFrame4 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3201
      orientation     =   0
      colorgradient1  =   16711680
      colorgradient2  =   8421376
      bordercolor     =   12017457
      caption         =   "Frame With Royal Theme"
      icon            =   "Example.frx":1110
      forecolor       =   -2147483633
      font            =   "Example.frx":14AC
      Begin CollectionControlXP.EviProgressBar EviProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         font            =   "Example.frx":14D8
         brushstyle      =   0
         color           =   12937777
         color2          =   12937777
      End
      Begin CollectionControlXP.EviButtons EviButtons5 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Test Progress Bar"
         buttonstyle     =   3
         icon            =   "Example.frx":1504
         iconwidth       =   16
         iconheight      =   16
         iconsize        =   0
         originalpicsizew=   17
         originalpicsizeh=   17
         font            =   "Example.frx":1918
         font            =   "Example.frx":1944
         mousepointer    =   99
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard XP Style"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
   End
   Begin CollectionControlXP.EviFrame EviFrame3 
      Height          =   2055
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3625
      orientation     =   2
      backcolor       =   16118256
      colorgradient2  =   14735318
      bordercolor     =   16118256
      caption         =   "Frame With Silver Theme"
      icon            =   "Example.frx":1970
      font            =   "Example.frx":1D0C
      Begin CollectionControlXP.EviButtons EviButtons4 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Button No Border"
         buttonstyle     =   4
         originalpicsizew=   0
         originalpicsizeh=   0
         font            =   "Example.frx":1D38
         font            =   "Example.frx":1D64
         mousepointer    =   99
      End
      Begin CollectionControlXP.EviTextBox EviTextBox3 
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2055
         _extentx        =   3625
         _extenty        =   1296
         text            =   "TextBox With Scrolling Vertical"
         scrolling       =   2
         font            =   "Example.frx":1D90
         dataformat      =   "Example.frx":1DBC
      End
   End
   Begin CollectionControlXP.EviFrame EviFrame2 
      Height          =   2055
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3625
      orientation     =   2
      backcolor       =   15529718
      colorgradient2  =   12117984
      bordercolor     =   15529718
      caption         =   "Frame With Olive Theme"
      icon            =   "Example.frx":1E00
      font            =   "Example.frx":219C
      Begin CollectionControlXP.EviButtons EviButtons3 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Button Office XP"
         originalpicsizew=   0
         originalpicsizeh=   0
         font            =   "Example.frx":21C8
         font            =   "Example.frx":21F4
         mousepointer    =   99
      End
      Begin CollectionControlXP.EviTextBox EviTextBox2 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2055
         _extentx        =   3625
         _extenty        =   1296
         text            =   "TextBox With Scrolling Horizontal"
         scrolling       =   1
         font            =   "Example.frx":2220
         dataformat      =   "Example.frx":224C
      End
   End
   Begin CollectionControlXP.EviFrame EviFrame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _extentx        =   4048
      _extenty        =   3625
      orientation     =   2
      backcolor       =   16244694
      colorgradient2  =   16241606
      showicon        =   0   'False
      caption         =   "Frame with Blue Theme"
      icon            =   "Example.frx":2290
      forecolor       =   -2147483630
      font            =   "Example.frx":262C
      Begin CollectionControlXP.EviButtons EviButtons2 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Standard Button"
         buttonstyle     =   0
         originalpicsizew=   0
         originalpicsizeh=   0
         font            =   "Example.frx":2658
         font            =   "Example.frx":2684
         mousepointer    =   99
      End
      Begin CollectionControlXP.EviButtons EviButtons1 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         caption         =   "Button XP Style"
         buttonstyle     =   3
         originalpicsizew=   0
         originalpicsizeh=   0
         font            =   "Example.frx":26B0
         font            =   "Example.frx":26DC
         mousepointer    =   99
      End
      Begin CollectionControlXP.EviTextBox EviTextBox1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "d"
         Top             =   480
         Width           =   2055
         _extentx        =   3625
         _extenty        =   661
         text            =   "This TextBox XP Style"
         alignment       =   2
         font            =   "Example.frx":2708
         dataformat      =   "Example.frx":2734
         showtooltiptext =   -1  'True
         tooltiptitle    =   "dd"
         tooltipicon     =   2
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   2760
      ScaleHeight     =   975
      ScaleWidth      =   2175
      TabIndex        =   24
      Top             =   4680
      Width           =   2175
   End
End
Attribute VB_Name = "Example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Counter As Long

Private Sub EviButtons10_Click()
Unload Me
End Sub

Private Sub EviButtons5_Click()
Dim m_Value As Integer
Do While 100
    m_Value = m_Value + 1
    EviProgressBar1.Value = m_Value
    If m_Value = 100 Then Exit Do
    Pause 0.001
Loop
EviProgressBar1.Value = 0
End Sub

Sub Pause(interval)
Dim Current
On Error GoTo Error
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
Error:
End Sub

Private Sub EviButtons6_Click()
Dim m_Value As Integer
Do While 100
    m_Value = m_Value + 1
    EviProgressBar2.Value = m_Value
    If m_Value = 100 Then Exit Do
    Pause 0.001
Loop
EviProgressBar2.Value = 0
End Sub

Private Sub EviButtons7_Click()
Dim m_Value As Integer
Do While 100
    m_Value = m_Value + 1
    EviProgressBar3.Value = m_Value
    If m_Value = 100 Then Exit Do
    Pause 0.001
Loop
EviProgressBar3.Value = 0
End Sub

Private Sub EviButtons8_Click()
m_Counter = 0
With EviAnimationForm1
Do While 300
    m_Counter = m_Counter + 1
    .GraphicForm.Animation Picture1, vbBlack, m_Counter
    If m_Counter = 300 Then Exit Do
    Pause 0.2
Loop
Picture1.Cls
End With
End Sub

Private Sub EviButtons9_Click()
Dim m_a As Long
With EviAnimationForm1
    For m_a = 0 To 8
        .GraphicForm.DrawGradient Picture1, m_a, vbRed, vbBlue
        Pause 0.2
    Next
    Picture1.Cls
End With
End Sub

Private Sub Form_Load()
With EviAnimationForm1
    .Add Me, EviButtons10, "Do you see this ToolTipText", Me.Caption, [Icon Info]
    .Show
End With
End Sub
