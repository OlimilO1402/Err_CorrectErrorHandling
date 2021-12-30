# Err_CorrectErrorHandling  
## How to handle errors in VBC the easy and correct way  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Err_CorrectErrorHandling?style=plastic)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Err_CorrectErrorHandling?style=plastic)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Err_CorrectErrorHandling/total.svg)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/releases/download/v1.0.15/ErrorHandling_v1.0.15.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/watchers)

Project started around may 2005.  

### General

In VBC we often see code similar to the following
```vba
    On Error GoTo ErrHandler
    '. . . some error-prone code here . . .
    Exit Sub/Function/Property
ErrHandler:
    MsgBox Err.Description
```

and most of the time they end up having plenty of MsgBoxes, doing similar things, spreaded all 
over the code. During an error the user often is in a kind of shock-situation so don't be rude 
and give informations what is to do now!

### Informations Needed

In Error-Messages the following Informations are badly needed:
 * the name of the class where the error occurs
 * the name of the function where the error occurs
 * some additional information about the specific object, the filename etc.
 * what the user could do next
 * how to avoid the error in the future
not only for the user but also for you the developer.

We could easily solve the task by using a globally available standard error message.
So let's use a module for our error messages (like module "MErr")

### Syntax

In VB.net there is the Try..Catch..Finally-syntax.
This is very useful because we have a standard syntax always for the same thing

But don't hesitate we can do it in VBC very similiar like this:
```vba
Sub DoIt()
Try: On Error GoTo Catch
    
	'here some error-prone code 
	
	GoTo Finally
Catch:
'. . .
Finally:
End Sub
```

Instead of "GoTo Finally" you could also use "Exit Sub", "Exit Function" or "Exit Property",
but using "Goto Finally" instead is more generic, because you even do not have to distinguish 
between Sub, Function or Property, so reusing the code is made more easily.

Now call the ErrHandler function, which can be private in every class, form or module.
Add the information: "name of the function", VB already knows the name of the class or form.
You even have the chance to call the function plenty of times, by using "Resume Try"
```vba
    If ErrHandler("Open", "Trying to open the file: " & PFN, , , , True) = vbRetry Then
        Resume Try
    End If
Finally:
End Sub
```

### Handling Errors Inline

If you, the developer, have fundamental knowledge about the errors that can occur in certain 
situations, you should handle the error inline in your code. In such a situation there is no 
need for "Try: On Error Goto" at all. This could be the case if for instance some API-functions or
even your own functions of course, return a Boolean whether a function succeeded or not. 
Do not use Err.Raise in the codes only meant to be used by yourself. 
Just use Err.Raise if you develop some API-functions for other developers, like for instance 
when developing controls, or dlls.

### Handling Errors Explicitely

In every other case for example if you develop with functions of the Windows-API use 
"Try: On Error GoTo" if there are explicit errors to occur.
In this case you get Error-codes and you have to translate them to a human readable language.
Just handle the error-code by using the "WinApiErr"-Variable to the ErrHandler function, then 
the error-code will be translated by using FormatMessageW.

This is how the function "ErrHandler" looks like. Just use it in every module, class, form or 
control, the name of it will be added automatically. 
In standard-modules the function "TypeName(Me)" will not work, so simply replace it then with 
the name of the module.

```vba
'copy this same function to every class or form
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    
    If bRetry Then
        
        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)
        
    Else
        
        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)
        
    End If
    
End Function
```

And the globally available Function MessError in the module "MErr" that finally shows the error-message, could look like this:
That's it, simple as that.

```vba
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
    Public Enum LongPtr
        [_]
    End Enum
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Function FormatMessageW Lib "kernel32.dll" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByRef Arguments As Long) As Long
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
                          Optional vbDecor As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult ' vbOKOnly Or vbCritical
    If bLoud Then

        Dim sErr As String:  sErr = ClsName & "::" & FncName
        If Len(AddInfo) Then sErr = sErr & vbCrLf & "Info:   " & AddInfo
        If Err.Number Then sErr = sErr & vbCrLf & "ErrNr " & Err.Number & ": " & Err.Description
        If Err.LastDllError Then sErr = sErr & vbCrLf & "DllErrNr: " & Err.LastDllError & " " & Err.Description
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
    Dim l As Long:   l = 512
    Dim s As String: s = Space(l)
    l = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, MessageID, 0&, StrPtr(s), l, ByVal 0&)
    If l Then WinApiError_ToStr = Left$(s, l)
End Function
```

![ErrorHandling Image](Resources/ErrorHandling.png "ErrorHandling Image")
