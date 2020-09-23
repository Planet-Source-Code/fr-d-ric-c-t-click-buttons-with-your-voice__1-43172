Attribute VB_Name = "Global"
Option Explicit

Private Const BN_CLICKED = 0
Private Const WM_COMMAND = &H111
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lparam As Long) As Long

Public mlngCurWindow As Long
Public mstrPhrase As String

'Function called each time EnumChildWindows finds a child
Public Function EnumAChild(ByVal hwnd As Long, ByVal lparam As Long) As Boolean

Dim lngButtonID As Long
Dim lngLength As Long
Dim strTitle As String
Dim lngResult As Long
    
    EnumAChild = True 'Let EnumChildWindows find another child
    lngLength = GetWindowTextLength(hwnd) 'get length of the caption
    If lngLength > 0 Then
        strTitle = String$(100, Chr(0))
        'get the caption
        lngResult = GetWindowText(hwnd, strTitle, lngLength + 1) 'length + \0
        strTitle = Left$(strTitle, lngLength)
        'if the caption of the control is the same as the word you said
        If UCase(strTitle) = UCase(mstrPhrase) Or (UCase(strTitle) = "OK" And UCase(mstrPhrase) = "OKAY") Then
            lngButtonID = GetDlgCtrlID(hwnd)
            lngResult = PostMessage(mlngCurWindow, WM_COMMAND, lngButtonID, BN_CLICKED * &H10000 + hwnd)
            If lngResult <> 0 Then
                'You found what you wanted, stop the EnumChildWindows function
                EnumAChild = False
            End If
        End If
    End If

End Function
