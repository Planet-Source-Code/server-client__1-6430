Attribute VB_Name = "Module1"
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInstertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Enum Desktop_Constants
        ison = True
        isoff = False
End Enum

Public Enum StartBar_Constants
        isontaskbar = 1
        innotontaskbar = 0
End Enum

Public Function StartButton(State As StartBar_Constants)
        Dim SendValue As Long
        Dim SetOption As Long
        SetOption = FindWindow("Shell_TrayWnd", "")
        SendValue = FindWindowEx(SetOption, 0, "Button", vbNullString)
        ShowWindow SendValue, State
End Function

Public Function TaskbarIcons(State As StartBar_Constants)
        Dim SendValue As Long
        Dim SetOption As Long
        SetOption = FindWindow("Shell_TrayWnd", "")
        SendValue = FindWindowEx(SetOption, 0, "TrayNotifyWnd", vbNullString)
        ShowWindow SendValue, State
End Function

Public Function Desktop(State As Desktop_Constants)
        Dim DesktopHwnd As Long
        Dim SetOption As Long
        DesktopHwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
        SetOption = IIf(State, SW_SHOW, SW_HIDE)
        ShowWindow DesktopHwnd, SetOption
End Function

Public Function NewLine()
        NewLine = vbCrLf
End Function



