Attribute VB_Name = "Calculator"
Declare Function FFIDivide Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long
Declare Function FFIMultiply Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long
Declare Function FFISubtract Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long
Declare Function FFISum Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long

Declare Function FFIString Lib "vbrs.dll" () As String
Declare Function my_hasher Lib "vbrs.dll" (ByVal myStr As String) As String
Declare Function how_many_characters Lib "vbrs.dll" (ByVal myStr As String) As Long

Function EnsureOperand(ByVal Expression As String) As String
    EnsureOperand = Replace(Replace(Replace(Replace(Expression, "+", ""), "-", ""), "*", ""), "/", "")
    Exit Function
End Function

Function Divide(ByVal x As Integer, ByVal y As Integer) As Integer
    Divide = FFIDivide(x, y)
    Exit Function
End Function

Function Multiply(ByVal x As Integer, ByVal y As Integer) As Integer
    Multiply = FFIMultiply(x, y)
    Exit Function
End Function

Function Subtract(ByVal x As Integer, ByVal y As Integer) As Integer
    Subtract = FFISubtract(x, y)
    Exit Function
End Function

Function Sum(ByVal x As Integer, ByVal y As Integer) As Integer
    Sum = FFISum(x, y)
    Exit Function
End Function

Function CountString(ByVal myStr As String) As Integer
  CountString = how_many_characters(myStr)
End Function

Function Hash(ByVal myStr As String) As String
  Hash = my_hasher(myStr)
End Function

