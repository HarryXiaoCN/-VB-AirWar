Attribute VB_Name = "Airwar_QT"
Public Type 二元解
    A As Single
    B As Single
End Type
Public Type 三原色
    R As Long
    G As Long
    B As Long
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Declare Function joySetCapture Lib "winmm.dll" (ByVal hwnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long

  
' think they are all necessary though.
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1
Public Const JOY_POVCENTERED = -1
Public Const JOY_POVFORWARD = 0
Public Const JOY_POVRIGHT = 9000
Public Const JOY_POVLEFT = 27000
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNU = &H10
Public Const JOY_RETURNV = &H20
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNRAWDATA = &H100&
Public Const JOY_RETURNPOVCTS = &H200&
Public Const JOY_RETURNCENTERED = &H400&
Public Const JOY_USEDEADZONE = &H800&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Public Const JOY_CAL_READALWAYS = &H10000
Public Const JOY_CAL_READRONLY = &H2000000
Public Const JOY_CAL_READ3 = &H40000
Public Const JOY_CAL_READ4 = &H80000
Public Const JOY_CAL_READXONLY = &H100000
Public Const JOY_CAL_READYONLY = &H200000
Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000
  
Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type
Public Function 彩虹函数() As 三原色
Randomize
彩虹函数.R = Int(Rnd * (256))
彩虹函数.G = Int(Rnd * (256))
彩虹函数.B = Int(Rnd * (256))
End Function
Public Function 线性求解(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Sd As Long) As 二元解
Dim m As Double
If X2 > X1 Then
    m = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
    线性求解.B = (Y2 - Y1) * (Sd / m)
    线性求解.A = Sqr(Sd ^ 2 - 线性求解.B ^ 2)
Else
    m = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
    线性求解.B = -(Y1 - Y2) * (Sd / m)
    线性求解.A = -Sqr(Sd ^ 2 - 线性求解.B ^ 2)
End If
End Function
Public Function 两点距离(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
两点距离 = Sqr((Y2 - Y1) ^ 2 + (X2 - X1) ^ 2)
End Function
Public Function 距离判断(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Limit As Long) As Boolean
If (Y2 - Y1) ^ 2 + (X2 - X1) ^ 2 <= Limit ^ 2 Then 距离判断 = True
End Function
Public Function 对角检查(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Range As Long) As Boolean
If X2 < X1 And Y2 < Y1 Then
    If Range >= (X2 - (X1 - (Y1 - Y2))) / Sqr(2) _
    And -Range <= (X2 - (X1 - (Y1 - Y2))) / Sqr(2) Then 对角检查 = True: Exit Function
End If
If X2 > X1 And Y2 < Y1 Then
    If Range >= ((X1 + (Y1 - Y2)) - X2) / Sqr(2) And _
    -Range <= ((X1 + (Y1 - Y2)) - X2) / Sqr(2) Then 对角检查 = True: Exit Function
End If
If X2 < X1 And Y2 > Y1 Then
    If Range >= ((X1 - X2) + Y1 - Y2) / Sqr(2) And _
    -Range <= ((X1 - X2) + Y1 - Y2) / Sqr(2) Then 对角检查 = True: Exit Function
End If
If X2 > X1 And Y2 > Y1 Then
    If Range >= (Y1 + (X2 - X1) - Y2) / Sqr(2) And _
    -Range <= (Y1 + (X2 - X1) - Y2) / Sqr(2) Then 对角检查 = True: Exit Function
End If
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

