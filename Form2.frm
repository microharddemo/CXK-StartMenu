VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN��ʳ"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "���㣬�޷�����"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "������ѡ��ϲ������ʳ"
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
MsgBox "ֻ��UWPӦ��ʵ����̫��" & vbCrLf & "�ٶ࿴һ�۵��Ծͻᱬը" & vbCrLf & "��������Ϊ��׼��������Ӧ��", , "MSN��ʳ"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
End Sub
