VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Just a dummy PathFileName class for showing the principle of cumulating
'necessary informations for the user and the developer during an Error.
'During an error the user often is in a kind of shock-situation
'so don't be rude and give informations what is to do now!

Public PFN As String
Private FNr As Integer

Public Function OOpen() As Integer

Try: On Error GoTo Catch
    
    If FNr = 0 Then FNr = FreeFile
    
    Open PFN For Binary Access Read Lock Read Write As FNr
    
    OOpen = FNr
    
    GoTo Finally
Catch:
    
    If ErrHandler("Open", "Trying to open the file: " & PFN, , , , , True) = vbRetry Then Resume Try

Finally:
    
End Function

Public Function ReadContent() As String
    
Try: On Error GoTo Catch
    
    If FNr = 0 Then FNr = OOpen
    
    If FNr = 0 Then Exit Function
    
    Dim s As String: s = Space(LOF(FNr))
    
    Get FNr, , s
    
    ReadContent = s
    
    GoTo Finally
Catch:
    
    If ErrHandler("ReadContent", , , , , , True) = vbRetry Then Resume Try
    
Finally:
    
End Function

Public Sub CClose()
    Close FNr
    FNr = 0
End Sub


'copy this same function to every class or form
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional WinApiErr, _
                            Optional BolLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    
    If bRetry Then
        
        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiErr, bErrLog)
        
    Else
        
        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiErr, BolLoud, bErrLog, vbDecor)
        
    End If
    
End Function




