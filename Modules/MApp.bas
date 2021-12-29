Attribute VB_Name = "MApp"
Option Explicit
Public Declare Function RegOpenKeyExA Lib "advapi32" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Sub Main()
    
    FMain.Show

End Sub
