Attribute VB_Name = "Airwar_QT"
Public Type 二元解
    a As Single
    b As Single
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function 线性求解(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal Sd As Long) As 二元解
Dim m As Double
If x2 > x1 Then
    m = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    线性求解.b = (y2 - y1) * (Sd / m)
    线性求解.a = Sqr(Sd ^ 2 - 线性求解.b ^ 2)
Else
    m = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    线性求解.b = -(y1 - y2) * (Sd / m)
    线性求解.a = -Sqr(Sd ^ 2 - 线性求解.b ^ 2)
End If
End Function
Public Function 两点距离(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
两点距离 = Sqr((y2 - y1) ^ 2 + (x2 - x1) ^ 2)
End Function
Public Function 距离判断(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal Limit As Long) As Boolean
If (y2 - y1) ^ 2 + (x2 - x1) ^ 2 <= Limit ^ 2 Then 距离判断 = True
End Function


