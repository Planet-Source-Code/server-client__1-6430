VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send a message"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   ControlBox      =   0   'False
   Icon            =   "msg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Below type the message you want to sent :"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3030
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "msg.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Text3.Text = "Msg " + Text1.Text
Form1.Winsock1.SendData Form1.Text3.Text
Form2.Hide
Form1.Enabled = True
Form1.SetFocus
End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Enabled = True
Form1.SetFocus
Text1.Text = ""
Form1.Text3.Text = ""
End Sub

Private Sub Command8_Click()


End Sub
