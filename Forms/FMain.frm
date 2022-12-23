VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Error-Handling"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCompleteGuard2 
      Caption         =   "Error only in Try"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton BtnCompleteGuard1 
      Caption         =   "No Error at all"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton BtnCompleteGuard3 
      Caption         =   "Error in Try and Finally"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton BtnCompleteGuard4 
      Caption         =   "Error only in Finally"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton BtnNesting2 
      Caption         =   "Nesting 2"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton BtnNesting1 
      Caption         =   "Nesting 1"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton BtnProvokeWinApiError 
      Caption         =   "Provoke WinApi Error"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton BtnStartExe 
      Caption         =   "Start Exe"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileClose2 
      Caption         =   "File Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileOpen2 
      Caption         =   "File Open"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileClose1 
      Caption         =   "File Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileOpen1 
      Caption         =   "File Open"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As String
Private m_FNr1 As Integer
Private m_FNr2 As Integer

Private m_File As PathFileName

Private Declare Function RegOpenKeyExA Lib "advapi32" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Sub Form_Load()
        
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    'for being able to show correct error handling we have to provoke an error at first.
    m_PFN = App.Path & "\testfile.txt"
    Set m_File = New PathFileName: m_File.PFN = m_PFN
    Me.BtnFileClose1.Enabled = False
    Me.BtnFileClose2.Enabled = False
    
End Sub


Private Sub BtnInfo_Click()
    
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription, vbInformation
    
End Sub

Private Sub BtnProvokeWinApiError_Click()
    
Try: On Error GoTo Catch
    
    Dim hr As Long: hr = RegOpenKeyExA(0, 0, 0, 0, 0)
    
    If hr <> 0 Then GoTo Catch
    
    GoTo Finally
Catch:
    ErrHandler "BtnProvokeWinApiError_Click", "Trying to access registry", hr
Finally:
End Sub

Private Sub BtnFileOpen1_Click()
    
Try: On Error GoTo Catch
    
    m_FNr1 = OOpen(m_PFN)
    
    Text1.Text = ReadContent(m_FNr1)
    
    If m_FNr1 Then ToggleBtn1
    
    GoTo Finally
Catch:
    
    If ErrHandler("BtnFileOpen1_Click", , , , , , True) = vbRetry Then Resume Try
    
Finally:
    
End Sub

Private Sub BtnFileClose1_Click()
    Close m_FNr1
    m_FNr1 = 0
    ToggleBtn1
End Sub

Public Sub ToggleBtn1()
    Me.BtnFileOpen1.Enabled = Not Me.BtnFileOpen1.Enabled
    Me.BtnFileClose1.Enabled = Not Me.BtnFileClose1.Enabled
End Sub

Public Sub ToggleBtn2()
    Me.BtnFileOpen2.Enabled = Not Me.BtnFileOpen2.Enabled
    Me.BtnFileClose2.Enabled = Not Me.BtnFileClose2.Enabled
End Sub

Private Sub BtnFileOpen2_Click()
    m_FNr2 = m_File.OOpen()
    Text2.Text = m_File.ReadContent
    If m_FNr2 Then ToggleBtn2
End Sub

Private Sub BtnFileClose2_Click()
    Close m_FNr2
    m_FNr2 = 0
    ToggleBtn2
End Sub

Private Sub BtnStartExe_Click()
    Shell App.Path & "\" & "ErrorHandling.exe", vbNormalFocus
End Sub


'in VBC we often see some code similar to the following
'    On Error GoTo ErrHandler
'    '. . . some error prone code . . .
'    Exit Sub/Function/Property
'ErrHandler:
'    MsgBox Err.Description

'and most of the time they end up having plenty of MsgBoxes, doing similar things, spread all over the code.

'During an error the user often is in a kind of shock-situation
'so don't be rude and give informations what is to do now!

'In Error-Messages the following Informations are _always_ needed:
' * the name of the class where the error occurs
' * the name of the function where the error occurs
' * some additional information about the specific object the filename etc.
' * what to do next
' * how to avoid this error
'not only for the user but essentially for you the developer

'we could easily solve the task by using a globally available standard error message
'so lets do a module for our error messages (see module "MErr")

'in VB.net we have the Try..Catch..Finally-syntax
'this is very useful because we have a standard syntax always for the same thing

'But don't hesitate we can do it in VBC very similiarly like this:
'just add "GoTo Finally" before "Catch:"

Private Function OOpen(PFN As String) As Integer
    
Try: On Error GoTo Catch
    
    Dim FNr As Integer: FNr = FreeFile
    
    Open PFN For Binary Access Read Lock Read Write As FNr
    
    OOpen = FNr
    
    GoTo Finally
    'here you could also use "Exit Sub", "Exit Function" or "Exit Property"
    'but using Goto Finally is more generic, because you even do not have to
    'distinguish between Sub, Function or Property, so code copying is easy
Catch:
    'call the ErrHandler function, which can be private in every class, form or module
    'add the information: "name of the function", the name of the class or form is known
    'you even have the chance to call the function more times
    If ErrHandler("Open", "Trying to open the file: " & PFN, , , , , True) = vbRetry Then Resume Try

Finally:
End Function

Private Function ReadContent(ByVal FNr As Integer) As String

Try: On Error GoTo Catch
    
    Dim s As String: s = Space(LOF(FNr))
    
    Get FNr, , s
    
    ReadContent = s
    
    GoTo Finally
Catch:
    If ErrHandler("ReadContent", , , , , , True) = vbRetry Then Resume Try
Finally:
End Function


' v ############################## v '    Code Nesting    ' v ############################## v '
Private Sub BtnNesting1_Click()
    Dim i As Long: i = 1
    Dim j As Long: j = 0
Try1: On Error GoTo Catch1
    'i = i / j
Try2:   On Error GoTo Catch2
        i = i / j
        GoTo Finally2
Catch2:
        MsgBox "Catch2"
Finally2:
        MsgBox "Finally2"
End_Try2:
    GoTo Finally1
Catch1:
    MsgBox "Catch1"
Finally1:
    MsgBox "Finally1"
End_Try1:
End Sub
'Result:
'Catch2
'Finally2
'Finally1

Private Sub BtnNesting2_Click()
    Dim i As Long: i = 1
    Dim j As Long: j = 0
Try1: On Error GoTo Catch1
    i = i / j
Try2:   On Error GoTo Catch2
        i = i / j
        GoTo Finally2
Catch2:
        MsgBox "Catch2"
Finally2:
        MsgBox "Finally2"
End_Try2:
    GoTo Finally1
Catch1:
    MsgBox "Catch1"
Finally1:
    MsgBox "Finally1"
End_Try1:
End Sub
'Result:
'Catch1
'Finally1
' ^ ############################## ^ '    Code Nesting    ' ^ ############################## ^ '

' v ############################## v '  Error in Try and/or Finally  ' v ############################## v '
Private Sub BtnCompleteGuard1_Click()
Try: On Error GoTo Catch
    Dim file As New PathFileName: file.PFN = m_PFN
    file.OOpen
    GoTo Finally
Catch:
    ErrHandler "BtnCompleteGuard1_Click", "Catch"
    Resume Finally
Finally: On Error GoTo Catch2
    file.CClose
    GoTo End_Try
Catch2:
    ErrHandler "BtnCompleteGuard1_Click", "Catch2"
End_Try:
End Sub

Private Sub BtnCompleteGuard2_Click()
    'Error will occur only in the Try-block
Try: On Error GoTo Catch
    Dim file As PathFileName ': file.PFN = m_PFN
    file.OOpen
    GoTo Finally
Catch:
    ErrHandler "BtnCompleteGuard2_Click", "Catch"
    Resume Finally
Finally: On Error GoTo Catch2
    'file.CClose
    GoTo End_Try
Catch2:
    ErrHandler "BtnCompleteGuard2_Click", "Catch2"
End_Try:
End Sub

Private Sub BtnCompleteGuard3_Click()
    'Error will occur twice, in Try- and in Finally-block
Try: On Error GoTo Catch
    Dim file As PathFileName ': file.PFN = m_PFN
    file.OOpen
    GoTo Finally
Catch:
    ErrHandler "BtnCompleteGuard3_Click", "Catch"
    Resume Finally
Finally: On Error GoTo Catch2
    file.CClose
    GoTo End_Try
Catch2:
    ErrHandler "BtnCompleteGuard3_Click", "Catch2"
End_Try:
End Sub

Private Sub BtnCompleteGuard4_Click()
    'Error will occur only in the Finally-block
Try: On Error GoTo Catch
    Dim file As PathFileName ': file.PFN = m_PFN
    'file.OOpen
    GoTo Finally
Catch:
    ErrHandler "BtnCompleteGuard4_Click", "Catch"
    Resume Finally
Finally: On Error GoTo Catch2
    file.CClose
    GoTo End_Try
Catch2:
    ErrHandler "BtnCompleteGuard4_Click", "Catch2"
End_Try:
End Sub
' ^ ############################## ^ '  Error in Try and/or Finally  ' ^ ############################## ^ '


'copy this same function to every class or form
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    
    If bRetry Then
        
        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)
        
    Else
        
        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)
        
    End If
    
End Function

