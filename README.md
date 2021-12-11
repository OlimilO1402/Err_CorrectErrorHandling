# Err_CorrectErrorHandling  
## How to handle errors in VBC the easy and correct way  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Err_CorrectErrorHandling?style=plastic)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Err_CorrectErrorHandling?style=plastic)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Err_CorrectErrorHandling/total.svg)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/releases/download/v1.0.12/ErrorHandling_v1.0.12.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/Err_CorrectErrorHandling/watchers)

Project started around may 2005.  

In VBC we often see code similar to the following
```vba
    On Error GoTo ErrHandler
    '. . . some error-prone code here . . .
    Exit Sub/Function/Property
ErrHandler:
    MsgBox Err.Description
```

and most of the time they end up having plenty of MsgBoxes, doing similar things, spreaded all over the code.

In Error-Messages the following Informations are badly needed:
 * the name of the class where the error occurs
 * the name of the function where the error occurs
 * some additional information about the specific object, the filename etc.
 * what the user could do next
 * how to avoid the error in the future
not only for the user but also for you the developer.

We could easily solve the task by using a globally available standard error message.
So let's use a module for our error messages (see module "MErr")

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

This is how the function "ErrHandler" looks like. Just use it in every class or form, the name 
of the class or form will be added automatically. In standard-modules the function "TypeName(Me)" 
will not work, so simply replace it with the name of the module.
```vba
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional BolLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    If bRetry Then
        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, bErrLog)
    Else
        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, BolLoud, bErrLog, vbDecor)
    End If
End Function
```

And the globally available Function MessError in the module "MErr" that finally shows the error-message, could look like this:
```vba
Public ErrLog As String

Public Function MessError(ClsName As String, FncName As String, _
                            Optional AddInfo As String = "", _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical) As VbMsgBoxResult
    If bLoud Then
        Dim sErr As String
        sErr = "Fehler: " & Err.Number & vbCrLf & _
               "in:     " & ClsName & "::" & FncName & vbCrLf & _
               "Info:   " & Err.Description & vbCrLf & AddInfo
        MessError = MsgBox(sErr, vbDecor)
        'On Error GoTo 0
        'Err.Clear
    End If
    If bErrLog Then
        ErrLog = ErrLog & vbCrLf & Now & " " & sErr
    End If
End Function

Public Function MessErrorRetry(ClsName As String, FncName As String, _
                               Optional AddInfo As String = "", _
                               Optional bErrLog As Boolean = True) As VbMsgBoxResult
    MessErrorRetry = MessError(ClsName, FncName, AddInfo, True, bErrLog, vbRetryCancel)
End Function
```
That's it, simple as that.

![ErrorHandling Image](Resources/ErrorHandling.png "ErrorHandling Image")
