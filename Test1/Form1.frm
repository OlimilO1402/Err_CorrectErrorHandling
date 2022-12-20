VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function MyFunction() As Boolean
Try: On Error GoTo Catch
    Dim F As File: 'Set F = New File
    'F.OpenFile
    MyFunction = True
    GoTo Finally
Catch:
    MsgBox "MyFunction1 Catch: " & Err.Number & " " & Err.Description
Finally: On Error GoTo 0
    'if the error occurs only in the Finally-block,
    'the app will jump back to the Catch-block, shows the error
    'then runs into the Finally-block again, and then does *not* crash
    'only if the Caller has got a error-handling mechanism too,
    'otherwise it will crash in any circumstance
    'if the finally-block starts with "On Error GoTo 0" the only thing what it does
    'it never runs into the Catch-block and you get no chance to get a decent Error-information.
    F.CloseFile
End Function

Private Sub Command1_Click()
    'Without any errorhandling in the caller, the app will crash!!!
Try: 'On Error Resume Next
    Dim i As Long
    i = MyFunction
End Sub
