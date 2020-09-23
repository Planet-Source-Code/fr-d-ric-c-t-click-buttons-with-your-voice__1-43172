VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Testing Voice Recognition"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   540
      TabIndex        =   4
      Top             =   1665
      Width           =   1410
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   540
      TabIndex        =   3
      Top             =   1170
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   675
      Width           =   1410
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   540
      TabIndex        =   1
      Top             =   180
      Width           =   1410
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR dsrReco 
      Height          =   330
      Left            =   45
      OleObjectBlob   =   "frmTest.frx":0000
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    MsgBox "You chose the Cancel button", vbOKOnly, "Test"

End Sub

Private Sub cmdNo_Click()

    MsgBox "You chose the No button", vbOKOnly, "Test"

End Sub

Private Sub cmdOk_Click()

    MsgBox "You chose the Ok button", vbOKOnly, "Test"

End Sub

Private Sub cmdYes_Click()

    MsgBox "You chose the Yes button", vbOKOnly, "Test"

End Sub

'When the Voice Recognition engine 'thinks' you finished talking
Private Sub dsrReco_PhraseFinish(ByVal Flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)

Dim lngResult As Long

    'phrase will contain the recognized words, only those in the grammar
    If Phrase <> "" Then
        mstrPhrase = Phrase
        'get the handle of the active window
        mlngCurWindow = GetForegroundWindow
        'get the childs of the active window and pass them to the EnumAChild function
        lngResult = EnumChildWindows(mlngCurWindow, AddressOf EnumAChild, 0)
        Debug.Print Phrase
    End If

End Sub

Private Sub Form_Load()

    dsrReco.GrammarFromFile "test.txt" 'load the grammar
    dsrReco.Activate

End Sub
