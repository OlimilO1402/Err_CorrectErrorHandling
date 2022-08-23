Attribute VB_Name = "MErr"
Option Explicit ' Zeilen: 91
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF&
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_FROM_STRING     As Long = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE    As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
#If VBA7 Then
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Private Declare PtrSafe Function FormatMessageW Lib "kernel32.dll" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByRef Arguments As Long) As Long
#Else
    'Public Enum LongPtr
    '    [_]
    'End Enum
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Function FormatMessageW Lib "kernel32.dll" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Any, ByVal nSize As Long, ByRef Arguments As Long) As Long
#End If
Public ErrLog As String

'here 4 different ways to get the error-code and 2 different
'ways to translate the error-code to a human readable string
' * VBC-Runtime:
'     a) Err.Number       -> Err.Description
'     b) Err.LastDllError -> Err.Description
' * Windows-API:
'     c) GetLastError (~=Err.LastDllError) -> FormatMessage
'     d) HResult or any other WinaPI-Error -> FormatMessage

Public Function MessError(ClsName As String, FncName As String, _
                          Optional AddInfo As String = "", _
                          Optional WinApiErr, _
                          Optional bLoud As Boolean = True, _
                          Optional bErrLog As Boolean = True, _
                          Optional vbDecor As VbMsgBoxStyle = vbOKCancel) As VbMsgBoxResult ' vbOKOnly Or vbCritical
    If bLoud Then

        Dim sErr As String:  sErr = ClsName & "::" & FncName
        If Len(AddInfo) Then sErr = sErr & vbCrLf & "Info:   " & AddInfo
        If Err.Number Then sErr = sErr & vbCrLf & "ErrNr " & Err.Number & ": " & Err.Description
        If Err.LastDllError Then sErr = sErr & vbCrLf & "DllErrNr: " & Err.LastDllError & ": " & WinApiError_ToStr(Err.LastDllError) '& Err.Description
        Dim LastError As Long: LastError = GetLastError
        If LastError Then sErr = sErr & vbCrLf & "LastError " & LastError & ": " & WinApiError_ToStr(LastError)
        If Not IsMissing(WinApiErr) Then sErr = sErr & vbCrLf & "WinApiErr " & WinApiErr & ": " & WinApiError_ToStr(WinApiErr)
        
        MessError = MsgBox(sErr, vbDecor)
    End If
    If bErrLog Then
        ErrLog = ErrLog & vbCrLf & Now & " " & sErr
    End If
End Function

Public Function MessErrorRetry(ClsName As String, FncName As String, _
                               Optional AddInfo As String = "", _
                               Optional WinApiErr, _
                               Optional bErrLog As Boolean = True) As VbMsgBoxResult
    MessErrorRetry = MessError(ClsName, FncName, AddInfo, WinApiErr, True, bErrLog, vbRetryCancel)
End Function

Public Function WinApiError_ToStr(ByVal MessageID As Long) As String
    'MessageID e.g. hResult
    Dim L As Long:   L = 512
    Dim s As String: s = Space(L)
    L = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, MessageID, 0&, StrPtr(s), L, ByVal 0&)
    If L Then WinApiError_ToStr = Left$(s, L)
End Function

''copy this same function to every class, form or module
''the name of the class or form will be added automatically
''in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
'' v ############################## v '   Local ErrHandler   ' v ############################## v '
'Private Function ErrHandler(ByVal FuncName As String, _
'                            Optional ByVal AddInfo As String, _
'                            Optional WinApiError, _
'                            Optional bLoud As Boolean = True, _
'                            Optional bErrLog As Boolean = True, _
'                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
'                            Optional bRetry As Boolean) As VbMsgBoxResult
'
'    If bRetry Then
'
'        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)
'
'    Else
'
'        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)
'
'    End If
'
'End Function
