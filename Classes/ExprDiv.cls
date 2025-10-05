VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expr
'dividend / divisor
Private m_dividend As Expr
Private m_divisor  As Expr

Friend Sub New_(dividend As Expr, divisor As Expr)
    Set m_dividend = dividend
    Set m_divisor = divisor
End Sub

'copy to MNew:
'Public Function ExprDiv(dividend As Expr, divisor As Expr) As ExprDiv
'    Set ExprDiv = New ExprDiv: ExprDiv.New_ dividend, divisor
'End Function

Private Function Expr_Eval() As Variant
    Expr_Eval = m_dividend.Eval / m_divisor.Eval
End Function
