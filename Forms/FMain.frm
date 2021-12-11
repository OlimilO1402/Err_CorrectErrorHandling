VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton BtnStartExe 
      Caption         =   "Start Exe"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileClose2 
      Caption         =   "File Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileOpen2 
      Caption         =   "File Open"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileClose1 
      Caption         =   "File Close"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton BtnFileOpen1 
      Caption         =   "File Open"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
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

'und wenn man statt Datei öffnen einfach umbenennen nimmt?

Private Sub Form_Load()
    
    'for showing correct error handling we have to produce/provoke an error
    m_PFN = App.Path & "\testfile.txt"
    Set m_File = New PathFileName: m_File.PFN = m_PFN
    Me.BtnFileClose1.Enabled = False
    Me.BtnFileClose2.Enabled = False
    
End Sub

Private Sub BtnFileOpen1_Click()
Try:
    On Error GoTo Catch
    m_FNr1 = OOpen(m_PFN)
    Text1.Text = ReadContent(m_FNr1)
    If m_FNr1 Then ToggleBtn1
    Exit Sub
Catch:
    If ErrHandler("BtnFileOpen1_Click", , , , , True) = vbRetry Then
        On Error GoTo -1
        GoTo Try
    End If
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
    'Shell App.Path & "\" & App.EXEName & ".exe"
    Shell App.Path & "\" & "ErrorHandling.exe", vbNormalFocus
End Sub


'in VBC we often see some code simliar to the following
'    On Error GoTo ErrHandler
'    '. . . some error prone code . . .
'    Exit Sub/Function/Property
'ErrHandler:
'    MsgBox Err.Description

'and they end up having plenty of MsgBoxes, doing similar things, spreaded all over the code.

'In Error-Messages almost _always_ the following Informations are needed:
' * the name of the class where the error occurs
' * the name of the function where the error occurs
' * some additional information about the specific object the filename etc.
' * what to do next
' * how to avoid this error
'not only for the user but essentially for you the developer

'we could easily solve the task by using a globally availabel standard error message
'so lets do a module for our error messages (see module "MErr")


'in VB.net we have the Try..Catch..Finally-syntax
'this is very useful because we have a standard syntax always for the same thing

'But don't hesitate we can do it in VBC very similiarly like this:
'just add "GoTo Finally" before "Catch:"
Private Function OOpen(PFN As String) As Integer
'    On Error GoTo Catch
'Try:
    
    Dim FNr As Integer: If FNr = 0 Then FNr = FreeFile
    
    Open PFN For Binary Access Read Lock Read Write As FNr
    
    OOpen = FNr
    
    'GoTo Finally
'Catch:
    'On Error GoTo 0
    'call the Errhandler function, which is private in every class, form or module
    'add the information: "name of the function", the name of the class or form is known
    'you even have to chance to call the funvction twice
    'Dim mr As VbMsgBoxResult
    'mr = ErrHandler("Open", "Trying to open the file: " & PFN, , , , True)
    'If ErrHandler("Open", "Trying to open the file: " & PFN, , , , True) = vbRetry Then
    'Close FNr
    'FNr = 0
    'Err.Raise 70
    'If mr = vbRetry Then GoTo Try
'Finally:
    'On Error GoTo 0
    'Err.Clear
End Function

Private Function ReadContent(ByVal FNr As Integer) As String
Try: On Error GoTo Catch
    Dim s As String: s = Space(LOF(FNr))
    Get FNr, , s
    ReadContent = s
    Exit Function
Catch:
    ErrHandler "ReadContent"
End Function

'copy this same function to every class or form
'the name of the class for form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional BolLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    'ErrHandler = GlobalErrHandler(Me, PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
    If bRetry Then
        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, bErrLog)
    Else
        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, BolLoud, bErrLog, vbDecor)
    End If
End Function

