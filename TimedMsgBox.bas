Attribute VB_Name = "TimedMsgBox"
Option Explicit

' IMPORTANT NOTE:
' Demo project showing how to use the Timed MessageBox
' by Anirudha Vengurlekar anirudhav@yahoo.com(http://domaindlx.com/anirudha)
' this demo is released into the public domain "as is" without
' warranty or guaranty of any kind.  In other words, use at your own risk.
' Please send me you comments or suggestions at anirudhav@yahoo.com
' Thanks in advance.

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const WM_CLOSE = &H10
Private Const BN_CLICKED = 0
Private Const WM_COMMAND = &H111

' Used for storing information
Private m_lMsgHandle As Long
Private m_lNoHandle As Long
Private m_lhHook As Long
Private bTimedOut As Boolean
Private sMsgText As String
Private lCount As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Function EnumChildWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim sClassName As String
    
    sClassName = Space(100)
    lRet = GetClassName(hWnd, sClassName, 100)
    sClassName = Left$(sClassName, lRet)
    
    Debug.Print sClassName
    If UCase$(sClassName) = UCase$("Button") Then
        m_lNoHandle = hWnd
        EnumChildWindowsProc = 0
    Else
        EnumChildWindowsProc = 1
    End If
    
End Function

' *********************************************************************************************************
' THIS IS CALLBACK procedure. Will called by Hook procedure
Private Function GetMessageBoxHandle(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' AV get the message box handle
    If lMsg = HCBT_ACTIVATE Then
        'Release the CBT hook
        m_lMsgHandle = wParam  ' Msg Box Window Handle
        UnhookWindowsHookEx m_lhHook
        m_lhHook = 0
        ' enumerate all the children so we can send a number
        ' button message to the No button if our box has one
        ' this avoids the Microsoft error in the message box
        ' Added by Daniels, Michael A (KPMG Group)
        EnumChildWindows m_lMsgHandle, AddressOf EnumChildWindowsProc, 0
    End If
    GetMessageBoxHandle = False
End Function

' *********************************************************************************************************
' THIS IS CALLBACK procedure. Will called by timer procedure
' This function is called when time out occurs by the timer
Private Sub MessageBoxTimerUpdateEvent(hWnd As Long, uiMsg As Long, idEvent As Long, dwTime As Long)
    Dim lRet As Long
    Dim sStr As String
    If m_lMsgHandle = 0 Then Exit Sub
    
    lCount = lCount + 1
    If sMsgText = "" Then
        sStr = Space(255)
        lRet = GetWindowText(m_lMsgHandle, sStr, 255)
        sStr = Left$(sStr, lRet)
        sMsgText = sStr
    End If
    sStr = sMsgText & " " & "(Time elapsed:" & lCount & ")"
    SetWindowText m_lMsgHandle, sStr
End Sub


' *********************************************************************************************************
' THIS IS CALLBACK procedure. Will called by timer procedure
' This function is called when time out occurs by the timer
Private Sub MessageBoxTimerEvent(hWnd As Long, uiMsg As Long, idEvent As Long, dwTime As Long)
    ' Close the message box
    
    'Debug.Print "Sending close message"
    
    If m_lNoHandle = 0 Then
        SendMessage m_lMsgHandle, WM_CLOSE, 0, 0
    Else
        Dim lButtonCommand
        
        lButtonCommand = (BN_CLICKED * (2 ^ 16)) And &HFFFF
        lButtonCommand = lButtonCommand Or GetDlgCtrlID(m_lNoHandle)
        
        SendMessage m_lMsgHandle, WM_COMMAND, lButtonCommand, m_lNoHandle
    End If
    
    m_lMsgHandle = 0  ' Set handle to ZERO
    m_lNoHandle = 0   ' Set handle to ZERO
    bTimedOut = True  ' Set flag to True
End Sub


' *********************************************************************************************************
Public Function MsgBoxEx(sMsgText As String, dwWait As Long, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional sTitle As String = "Timed MessageBox Demo") As VbMsgBoxResult
    Dim lTimer As Long
    Dim lTimerUpdate As Long
    
    ' SET CBT hook
    m_lhHook = SetWindowsHookEx(WH_CBT, AddressOf GetMessageBoxHandle, App.hInstance, GetCurrentThreadId())
    ' set the timer
    lTimer = SetTimer(0, 0, dwWait * 1000, AddressOf MessageBoxTimerEvent) ' Set timer
    lTimerUpdate = SetTimer(0, 0, 1 * 1000, AddressOf MessageBoxTimerUpdateEvent)  ' Set timer
    ' Set the flag to false
    bTimedOut = False
    ' Display the message Box
    MsgBoxEx = MsgBox(sMsgText, Buttons, sTitle)
    ' Kill the timer
    Call KillTimer(0, lTimer)
    Call KillTimer(0, lTimerUpdate)
    ' Return ZERO so that caller routine will decide what to do
    sMsgText = ""
    lCount = 0
    If bTimedOut = True Then MsgBoxEx = 0
End Function
' *********************************************************************************************************
