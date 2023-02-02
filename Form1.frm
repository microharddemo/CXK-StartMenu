VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "开团"
   ClientHeight    =   10770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14670
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2295
      ScaleWidth      =   2295
      TabIndex        =   12
      Top             =   7800
      Width           =   2295
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "桌面"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   72
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2295
      ScaleWidth      =   4815
      TabIndex        =   7
      Top             =   5280
      Width           =   4815
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "点击此处可前往网站检查更新"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Microhard Demo制作"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "开始蔡单 1.14.51.4 Beta"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "关于此应用"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   3720
      ScaleHeight     =   2295
      ScaleWidth      =   2295
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "MSN体育"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   42
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2295
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   42
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "MSN美食"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "开团"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   36
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
End
End Sub

Private Sub Label10_Click()
End
End Sub

Private Sub Label11_Click()
End
End Sub

Private Sub Label2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Label3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Label4_Click()
Form3.Show
Unload Me
End Sub

Private Sub Label5_Click()
Form3.Show
Unload Me
End Sub

Private Sub Label6_Click()
Shell ("explorer.exe https://microharddemo.github.io/2023/01/25/cxk.html")
End
End Sub

Private Sub Label7_Click()
Shell ("explorer.exe https://microharddemo.github.io/2023/01/25/cxk.html")
End
End Sub

Private Sub Label8_Click()
Shell ("explorer.exe https://microharddemo.github.io/2023/01/25/cxk.html")
End
End Sub

Private Sub Label9_Click()
Shell ("explorer.exe https://microharddemo.github.io/2023/01/25/cxk.html")
End
End Sub

Private Sub Picture1_Click()
Form2.Show
Unload Me
End Sub

Private Sub Picture2_Click()
Form3.Show
Unload Me
End Sub

Private Sub Picture3_Click()
Shell ("explorer.exe https://microharddemo.github.io/2023/01/25/cxk.html")
End
End Sub

Private Sub Picture4_Click()
End
End Sub
