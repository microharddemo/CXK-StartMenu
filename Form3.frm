VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN体育"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "第114届CXK篮球联赛"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label Label4 
         Caption         =   "比赛结果：小黑子队获胜"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "目前比分：114 : 514"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "比赛状态：已结束，点击查看回放"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "对阵双方：iKUN VS 小黑子"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox "只因UWP应用实在是太美" & vbCrLf & "再多看一眼电脑就会爆炸" & vbCrLf & "所以我们为你准备了桌面应用", , "MSN体育"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
End Sub

Private Sub Label2_Click()
Shell ("explorer.exe https://www.bilibili.com/video/BV1J4411v7g6")
End Sub
