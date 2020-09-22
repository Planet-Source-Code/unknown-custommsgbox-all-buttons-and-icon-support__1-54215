Attribute VB_Name = "Module1"
'This code is free to use in your program.
'Coded in VB 6
'Version 2 of my code.
'Age of programmer: 15
'Contact: Natalichwolf1n on Yahoo! Messenger
'Declares of Windows API
Type MB
 Ok As String ' Custom okbutton holder
 Hook As Long  ' Holds the hook
 Retry As String
 Cancel As String
 Ignore As String
 No As String
 Yes As String
End Type
Public Declare Function GetCurrentThreadId& Lib "kernel32" ()
Public Declare Function UnhookWindowsHookEx& Lib "user32" (ByVal hHook&)
Public Declare Function SetWindowsHookEx& Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&)
Public Declare Function SetDlgItemText& Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg&, ByVal nIDDlgItem&, ByVal lpString As String)
Public Declare Function MessageBox& Lib "user32" Alias "MessageBoxA" (ByVal hwnd&, ByVal lpText As String, ByVal lpCaption As String, ByVal wType&)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd&, ByVal nIndex&)
'Variables
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_OKCANCEL = &H1&
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_OK = &H0&
Dim MT As MB 'MT = MessageBoxType variable
'All the Messagebox buttons
Public Const IDOK = 1
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDYES = 6
Public Const IDCANCEL = 2
Public Const IDIGNORE = 5
Public Const IDNO = 7

Public Const MB_TASKMODAL = &H2000&

'Functions
Function MsgBoxA(strMsg As String, strCap As String, okButton As String, Optional Flags As String, Optional RetryB As String, Optional CancelB As String, Optional YesB As String, Optional NoB As String, Optional IgnoreB As String) As Long
'Variable list:
'strMsg - The message of the messagebox
'strCap - The caption
'okButton - What to label the ok button
On Error Resume Next ' error protection
If Flags = vbNullString Then
Flags = MB_OK
End If
MT.Hook = SetWindowsHookEx(5, AddressOf MsgBoxP, GetWindowLong(0, (-6)), GetCurrentThreadId) ' Ok this is dangerous because of hooking, but it works good to edit stuff you normally cant, like the Ok button
MT.Ok = okButton$ ' mov okButton,MT.msg - asm version. This moves the text to the holder stack
MT.Cancel = CancelB
MT.Ignore = IgnoreB
MT.No = NoB
MT.Retry = RetryB
MT.Yes = YesB
MsgBoxA = MessageBox(0, strMsg$, strCap$, Flags Or MB_TASKMODAL) ' Calls the MessageBoxA function in the user32.dll causing the code to work:)
End Function

Function MsgBoxP(ByVal UINT As Long, ByVal wParam As Long, lParam As Long) As Long
If UINT = 5 Then ' 5= WH_CBT(WindowsHook_ComputerBasedTraining)
SetDlgItemText wParam, IDOK, MT.Ok
SetDlgItemText wParam, IDCANCEL, MT.Cancel
SetDlgItemText wParam, IDIGNORE, MT.Ignore
SetDlgItemText wParam, IDNO, MT.No
SetDlgItemText wParam, IDYES, MT.Yes
SetDlgItemText wParam, IDRETRY, MT.Retry
'Edits the text on the ok button
UnhookWindowsHookEx MT.Hook
'application would actually crash if you dont do this
MsgBoxP = 0
 'Makes MsgBoxP return 0 so it wont error the function/app
End If

End Function
