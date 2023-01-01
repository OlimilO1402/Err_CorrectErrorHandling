Attribute VB_Name = "MNew"
Option Explicit

Public Function ExprVal(ByVal number) As ExprVal
    Set ExprVal = New ExprVal: ExprVal.New_ number
End Function

Public Function ExprDiv(dividend As Expr, divisor As Expr) As ExprDiv
    Set ExprDiv = New ExprDiv: ExprDiv.New_ dividend, divisor
End Function

Public Function MaybeNothing() As Maybe
    '
End Function

Public Function MaybeJust() As Maybe
    '
End Function
