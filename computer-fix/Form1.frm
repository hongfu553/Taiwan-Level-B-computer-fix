VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "電腦硬體裝修乙級第一站第9題"
   ClientHeight    =   11685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18810
   LinkTopic       =   "Form1"
   ScaleHeight     =   11685
   ScaleWidth      =   18810
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "Connect Bluetooth"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   4
      Top             =   2760
      Width           =   3735
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   17280
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   9120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   7560
      TabIndex        =   3
      Top             =   8040
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   2
      Left            =   12120
      TabIndex        =   2
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Green LED"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   1920
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   2554
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   3188
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   3840
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   4456
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   5090
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   5724
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape G 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   0
      Left            =   6360
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   7
      Left            =   10440
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   6
      Left            =   11145
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   5
      Left            =   11850
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   4
      Left            =   12540
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   3
      Left            =   13245
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   2
      Left            =   13950
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   1
      Left            =   14655
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape R 
      FillStyle       =   0  '實心
      Height          =   375
      Index           =   0
      Left            =   15360
      Shape           =   3  '圓形
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   4215
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
    Command2.Caption = "Connect Bluetooth"
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
    G(i).FillColor = vbWhile
    R(i).FillColor = vbWhile
Next

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
