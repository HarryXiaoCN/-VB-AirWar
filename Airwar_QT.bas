Attribute VB_Name = "Airwar_QT"
Public Type ��Ԫ��
    a As Single
    b As Single
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function �������(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal Sd As Long) As ��Ԫ��
Dim m As Double
If x2 > x1 Then
    m = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    �������.b = (y2 - y1) * (Sd / m)
    �������.a = Sqr(Sd ^ 2 - �������.b ^ 2)
Else
    m = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    �������.b = -(y1 - y2) * (Sd / m)
    �������.a = -Sqr(Sd ^ 2 - �������.b ^ 2)
End If
End Function
Public Function �������(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
������� = Sqr((y2 - y1) ^ 2 + (x2 - x1) ^ 2)
End Function
Public Function �����ж�(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal Limit As Long) As Boolean
If (y2 - y1) ^ 2 + (x2 - x1) ^ 2 <= Limit ^ 2 Then �����ж� = True
End Function


