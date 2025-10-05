VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Error-Handling"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestFncAssert 
      Caption         =   "Test Function Assert"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton BtnMonadic 
      Caption         =   "Monadic Error Handling"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton BtnCompleteGuard5 
      Caption         =   "Error only in Finally (b)"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton BtnCompleteGuard4 
      Caption         =   "Error only in Finally (a)"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton BtnCompleteGuard3 
      Caption         =   "Error in Try and Finally"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton BtnCompleteGuard2 
      Caption         =   "Error only in Try"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton BtnCompleteGuard1 
      Caption         =   "No Error at all"
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton BtnNesting2 
      Caption         =   "Nesting 2"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton BtnNesting1 
      Caption         =   "Nesting 1"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton BtnProvokeWinApiError 
      Caption         =   "Provoke WinApi Error"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton BtnStartExe 
      Caption         =   "Start Exe"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton BtnFileClose2 
      Caption         =   "File Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton BtnFileOpen2 
      Caption         =   "File Open"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton BtnFileClose1 
      Caption         =   "File Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton BtnFileOpen1 
      Caption         =   "File Open"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
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

Private Sub BtnMonadic_Click()

    'Joshua Johanan
    'Dont throw exceptions in CSharp use Monads
    'https://ejosh.co/de/2022/07/dont-throw-exceptions-in-csharp-use-monads/
    
    'Computerphile
    'What is a Monad? - Computerphile
    'https://www.youtube.com/watch?v=t1e8gqXLbsU
    '
    'data Expr = Val Int | Div Expr Expr
    '
    'Math          |  Haskell
    '1             |  Val 1
    '6 / 2         |  Div (Val 6) (Val 2)
    '6 / (3 / 1)   |  Div (Val 6) (Div (Val 3) (Val 1))
    '
    'eval :: Expr -> Int               'actually this should be a Double or a Decimal
    'eval (Val n)   = n
    'eval (Div x y) = eval x / eval y
    
    Dim one As Expr:  Set one = MNew.ExprVal(1)
    Dim res As Maybe: Set res = one.Eval
    Dim six As Expr:  Set six = MNew.ExprVal(6)
    'if y evals to a zero the program will crash
    'what do we do to fix this problem?
    '
    'safediv :: Int -> Int -> Maybe Int
    '
    '"Maybe" is the way that we deal with things that can possibly fail in Haskell
    '
    'safediv n m = if m == 0 then
    '                  Nothing
    '              else
    '                  Just (n / m)
    '
    'Nothing is one of the constructors in the Maybe-Type
    'Just    is another    constructor  in the Maybe-type
    '
    'Better version:
    '
    'eval :: Expr -> Maybe Int
    'eval (Val n)   = Just n
    'eval (Div x y) = case eval x of
    '                     Nothing -> Nothing
    '                     Just    -> case eval y of
    '                                    Nothing -> Nothing
    '                                    Just m  -> safediv n m
    'Pattern: 2 case analyses
    'abstact them out and have them as a definition
    'Picture:
    '
    '       m                      m:=Maybe
    'case [///] of
    '     Nothing -> Nothing
    '     Just x  -> [\\\] x
    '                  f           f:=function
    '
    'm >>= f = case m of
    '              Nothing -> Nothing
    '              Just x  -> f x
    '
    'eval :: Expr  -> Maybe Int
    'eval (Val n)   = return n
    'eval (Div x y) = eval x >>= (lam n ->
    '                 eval y >>= (lam m ->
    '                 safediv n m))
    '13:23
    'Do-notation what has it to do with monads
    '
    'eval :: Expr -> Maybe Int
    'eval (Val n) = return n
    'eval (Div x y) = do n <- eval x
    '                    m <- eval y
    '                    safediv n m
    '
    'The Maybe-Monad
    '---------------
    '   return ::  a -> Maybe a
    '     >>=  ::
    '
    'it gives you a bridge between the pure world of values here
    'and the impure world of things tha could go wrong
    'so its a bridge from pure to impure if you like
    '
    'What's the point?
    '
    '1. Same idea works for _other_effects_ as well
    '2. Supports _pure_programming_ with effects
    '3. Use of effects explicit in types
    '4. Functions that work for _any_effect_
    '
    'Graham Hutton "Programming in Haskell" 2. Edition
    '
    'Visual Basic:
    'We have the Type Variant which itself can have either a Value or can be Empty or Nothing
    'and so are the Nullable-Types in .net
    'so in other words a Monad is a function that returns a Variant

End Sub

Private Sub BtnTestFncAssert_Click()
    
    MsgBox "Testing Function Assert():"
    MsgBox "We define a variable 'taf As TestingAssertFoo', but we forgot to assigen a value to the member 'Bar As Variant'"
    Dim taf As TestingAssertFoo
    
    MsgBox "now we hand the variable 'taf' over to a function 'morphfoo'."
    morphFoo taf
    
End Sub

Private Function morphFoo(aFoo As TestingAssertFoo) As Long
    MsgBox "In the function morphfoo we try to get value from the 'foo'"
    MsgBox "The following message you will only get to see in debug-mode during designtime"
    'note: every debug-line will be removed in kompiledexe-Release, so will only be executed during debug
    Debug.Assert Assert(aFoo.Bar <> Empty, TypeName(Me), "morphFoo", "foo.Bar must not be empty")
    morphFoo = aFoo.Bar
End Function

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
    'No error at all
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
    Dim file As PathFileName ' The object never got created
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
    Dim file As PathFileName ' The object never got created
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
    Dim file As PathFileName ' The object never got created
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

Private Sub BtnCompleteGuard5_Click()
    'swap things out to another function
Try: On Error GoTo Catch
    TryToOpenFile
Catch:
    ErrHandler "BtnCompleteGuard5_Click", "Catch"
End Sub

Private Sub TryToOpenFile()
    'Error will occur only in the Finally-block
Try: On Error GoTo Catch
    Dim file As PathFileName ' The object never got created
    'file.OOpen
    GoTo Finally
Catch:
    ErrHandler "TryToOpenFile", "Catch"
Finally:
    file.CClose
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

