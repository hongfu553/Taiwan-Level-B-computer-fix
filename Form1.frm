VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "硬體裝修乙級第一站第9題"
   ClientHeight    =   10740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16365
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   15.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   16365
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "Connect Bluetooth"
      Height          =   855
      Left            =   6840
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   1455
      Index           =   3
      Left            =   6720
      TabIndex        =   3
      Top             =   8040
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Red LED"
      Height          =   1455
      Index           =   2
      Left            =   10200
      TabIndex        =   2
      Top             =   5880
      Width           =   2775
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   15120
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   8880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Green LED"
      Height          =   1455
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   2520
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   3102
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   3684
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   4266
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   4848
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   5430
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   6012
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   0
      Left            =   6600
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   8160
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   9000
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   9840
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   10680
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   11520
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   12360
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   13200
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   0
      Left            =   14040
      Shape           =   3  '圓形
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   1935
      Left            =   5520
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b(99), c As Integer

Private Sub Command1_Click(Index As Integer)
a = Index
c = 0
End Sub
Private Sub display(no)
For i = 0 To 7
    If no Mod 2 = 1 And a = 1 Then G(i).FillColor = RGB(0, 255, 0)
    If no Mod 2 = 1 And a = 2 Then R(i).FillColor = RGB(255, 0, 0)
    no = no \ 2
Next i
End Sub
Private Sub Command2_Click()
If MSComm1.PortOpen Then
    MSComm1.Output = "R0"
    MSComm1.Output = "G0"
    MSComm1.PortOpen = False
    Command2.Caption = "connect Bluetooth"
Else
    MSComm1.PortOpen = True
    Command2.Caption = "Disconnect Bluetooth"
    MSComm1.Output = "R0"
    MSComm1.Output = "G0"
End If
End Sub

Private Sub Timer1_Timer()
b(0) = 1 + 2
b(1) = 2 + 4
b(2) = 4 + 8
b(3) = 8 + 16
b(4) = 16 + 32
b(5) = 32 + 64
b(6) = 64 + 128

Label1.Caption = "Current Time:" & Time$

For i = 0 To 7
    G(i).FillColor = vbWhite
    R(i).FillColor = vbWhite
Next i

If MSComm1.PortOpen Then
    For i = 0 To 7
        G(i).FillColor = RGB(0, 128, 0)
        R(i).FillColor = RGB(128, 0, 0)
    Next i
    
    If a = 1 Then MSComm1.Output = "G" & b(c): display (b(c))
    If a = 2 And c <= 8 Then MSComm1.Output = "R" & 2 ^ c: display (2 ^ c)
End If

If a = 3 Then End
If c > 15 Then c = 15 Else c = c + 1


End Sub
