VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN美食"
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
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1755
      ItemData        =   "Form2.frx":424A
      Left            =   120
      List            =   "Form2.frx":4266
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "金额不足，无法开团"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "尽情挑选你喜欢的美食"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox "只因UWP应用实在是太美" & vbCrLf & "再多看一眼电脑就会爆炸" & vbCrLf & "所以我们为你准备了桌面应用", , "MSN美食"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
End Sub
