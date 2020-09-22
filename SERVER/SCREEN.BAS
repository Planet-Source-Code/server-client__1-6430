Attribute VB_Name = "Module2"
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    '
    '2. The Function
    ' You can add this to your form's code
    ' or you can put it in a module if the declaration is in a module
    '
Public Function fSaveGuiToFile(ByVal theFile As String) As Boolean


    ' Name: fSaveGuiToFile
    ' Author: Dalin Nie
    ' Written: 4/2/99
    ' Purpose:
    ' This procedure will Capture the Screen or the active window of
    '     your Computer and Save it as
    ' a .bmp file
    ' Input:
    ' theFile file Name with path, where you want the .bmp to be save
    '     d
    '
    ' Output:
    ' True if successful
    '
    Dim lString As String
    On Error GoTo Trap
    'Check if the File Exist
    If Dir(theFile) <> "" Then Exit Function
    'To get the Entire Screen
    Call keybd_event(vbKeySnapshot, 1, 0, 0)

    SavePicture Clipboard.GetData(vbCFBitmap), theFile
    fSaveGuiToFile = True
    Exit Function
Trap:
    'Error handling

    MsgBox "Error Occured in fSaveGuiToFile. Error #: " & Err.Number & ", " & Err.Description
End Function
