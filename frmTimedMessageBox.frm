VERSION 5.00
Begin VB.Form frmTimedMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timed MessageBox Demo"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmTimedMessageBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3315
      TabIndex        =   3
      Top             =   1185
      Width           =   1215
   End
   Begin VB.CommandButton cmdShowMessage 
      Caption         =   "Show Timed  Message Box."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      TabIndex        =   2
      Top             =   555
      Width           =   4530
   End
   Begin VB.TextBox txtWait 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Text            =   "5"
      Top             =   60
      Width           =   3075
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   510
      Left            =   45
      TabIndex        =   4
      Top             =   1170
      Width           =   3225
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   4755
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Wait in Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   15
      TabIndex        =   0
      Top             =   75
      Width           =   1410
   End
End
Attribute VB_Name = "frmTimedMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' IMPORTANT NOTE:
' Demo project showing how to use the Timed MessageBox
' by Anirudha Vengurlekar anirudhav@yahoo.com(http://domaindlx.com/anirudha)
' this demo is released into the public domain "as is" without
' warranty or guaranty of any kind.  In other words, use at
' your own risk.
' Please send me you comments or suggestions at anirudhav@yahoo.com
' Thanks in advance.
Option Explicit

Private Sub cmdExit_Click()
    ' End the application
    End
End Sub

Private Sub cmdShowMessage_Click()
    Dim sStr As String
    Dim lRet As VbMsgBoxResult
    
    If Val(txtWait) >= 1 Then
        sStr = "This Message is displayed only for " & txtWait & " Sec."
        sStr = sStr & vbCrLf & " After Time out Cancel will be selectecd automatically "
        lRet = MsgBoxEx(sStr, Val(txtWait), vbAbortRetryIgnore)
        If lRet = 0 Then
            lblInfo = "Time out occurs. Cancel Selected."
        Else
            lblInfo = lRet & " Selected"
        End If
    Else
        sStr = "For demonstration please enter value grater than zero."
        MsgBox sStr, vbOKOnly + vbInformation
    End If
End Sub
