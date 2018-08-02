Attribute VB_Name = "Airwar_QT"
Public Type 二元解
    a As Single
    b As Single
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function 线性求解(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Sd As Long) As 二元解
Dim m As Double
If X2 > X1 Then
    m = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
    线性求解.b = (Y2 - Y1) * (Sd / m)
    线性求解.a = Sqr(Sd ^ 2 - 线性求解.b ^ 2)
Else
    m = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
    线性求解.b = -(Y1 - Y2) * (Sd / m)
    线性求解.a = -Sqr(Sd ^ 2 - 线性求解.b ^ 2)
End If
End Function
Public Function 两点距离(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
两点距离 = Sqr((Y2 - Y1) ^ 2 + (X2 - X1) ^ 2)
End Function
Public Function 距离判断(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Limit As Long) As Boolean
If (Y2 - Y1) ^ 2 + (X2 - X1) ^ 2 <= Limit ^ 2 Then 距离判断 = True
End Function
Public Function KeyCodeToStr(KeyCode As Integer, Optional Shift As Integer) As String
    If KeyCode > 47 And KeyCode < 91 Then
        KeyCodeToStr = Chr(KeyCode)
        Exit Function
    ElseIf KeyCode > 111 And KeyCode < 124 Then
        KeyCodeToStr = "F" & KeyCode - 111
        Exit Function
    ElseIf KeyCode > 95 And KeyCode < 106 Then
        KeyCodeToStr = "Keypad " & Str(KeyCode - 96)
        Exit Function
    End If
    Select Case KeyCode
        Case 8: KeyCodeToStr = "Back"
        Case 9: KeyCodeToStr = "Tab"
        Case 12: KeyCodeToStr = "Clear"
        Case 13: KeyCodeToStr = "Enter"
        Case 16: KeyCodeToStr = "Shift"
        Case 17: KeyCodeToStr = "Ctrl"
        Case 18: KeyCodeToStr = "Alt"
        Case 19
            KeyCodeToStr = "Pause"
        Case 20
            KeyCodeToStr = "Caps Lock"
        Case 27
            KeyCodeToStr = "Esc"
        Case 32
            KeyCodeToStr = "Space"
        Case 33
            KeyCodeToStr = "Page Up"
        Case 34
            KeyCodeToStr = "Page Down"
        Case 35
            KeyCodeToStr = "End"
        Case 36
            KeyCodeToStr = "Home"
        Case 41
            KeyCodeToStr = "Select"
        Case 42
            KeyCodeToStr = "Print Screen"
        Case 43
            KeyCodeToStr = "Execute"
        Case 44
            KeyCodeToStr = "SnapShot"
        Case 45
            KeyCodeToStr = "Insert"
        Case 46
            KeyCodeToStr = "Delete"
        Case 47
            KeyCodeToStr = "Help"
        Case 106
            KeyCodeToStr = "Keypad *"
        Case 107
            KeyCodeToStr = "Keypad +"
        Case 109
            KeyCodeToStr = "Keypad -"
        Case 110
            KeyCodeToStr = "Keypad Delete"
        Case 111
            KeyCodeToStr = "Keypad /"
        Case 144
            KeyCodeToStr = "Num Lock"
        Case 189
            If Shift = 0 Then KeyCodeToStr = "-" Else If Shift = 1 Then KeyCodeToStr = "_"
        Case 187
            If Shift = 0 Then KeyCodeToStr = "=" Else If Shift = 1 Then KeyCodeToStr = "+"
        Case 255
            KeyCodeToStr = "Unknown"
        Case 192
            If Shift = 0 Then KeyCodeToStr = "`" Else If Shift = 1 Then KeyCodeToStr = "~"
        Case 37
            KeyCodeToStr = "Left Arrow"
        Case 38
            KeyCodeToStr = "Up Arrow"
        Case 39
            KeyCodeToStr = "Right Arrow"
        Case 40
            KeyCodeToStr = "Dowm Arrow"
        Case 219
            If Shift = 0 Then KeyCodeToStr = "[" Else If Shift = 1 Then KeyCodeToStr = "{"
        Case 221
            If Shift = 0 Then KeyCodeToStr = "]" Else If Shift = 1 Then KeyCodeToStr = "}"
        Case 186
            If Shift = 0 Then KeyCodeToStr = ";" Else If Shift = 1 Then KeyCodeToStr = ":"
        Case 222
            If Shift = 0 Then KeyCodeToStr = "'" Else If Shift = 1 Then KeyCodeToStr = """"
        Case 220
            If Shift = 0 Then KeyCodeToStr = "\" Else If Shift = 1 Then KeyCodeToStr = "|"
        Case 188
            If Shift = 0 Then KeyCodeToStr = "," Else If Shift = 1 Then KeyCodeToStr = "<"
        Case 190
            If Shift = 0 Then KeyCodeToStr = "." Else If Shift = 1 Then KeyCodeToStr = ">"
        Case 191
            If Shift = 0 Then KeyCodeToStr = "/" Else If Shift = 1 Then KeyCodeToStr = "?"
        Case 193
            KeyCodeToStr = "\"
        Case Else
            KeyCodeToStr = "Unknown"
    End Select
End Function

