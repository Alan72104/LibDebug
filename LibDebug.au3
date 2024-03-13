; Update history
; 3/13/2024 - Add `Const` to ca() array parameter
;             Add ArrayAdd(), ArrayFind(), Min() and Max()
; 8/30/2023 - Add Consoleout Mock cmock(),
;             useful for function local debugging,
;             `Local $c = $debugging ? c : cmock`
; 4/19/2023 - Make c() return the original type instead of string
; 4/14/2023 - Add CheckedHotKeySet()
; 1/20/2023 - Add ConsoleGet cget()
; 3/2/2022 - Add ConsoleoutTimerDiff ct()
; 1/27/2022 - Remove Eval() usage in Throw() and cv()
; 1/12/2022 - Refactor a few lines of comment, and happy new year!
;             TODO: Remove the stupid $_LD_Debug things,
;             throw("obsolete") on invocation instead
; 11/7/2021 - Add comment about newline interpolation
; 11/7/2021 - Add newline interpolation ("\n") to string functions
; 10/17/2021 - Fix functions interpolating wrong variable
; 8/4/2021 - Rename throw() to Throw(),
;            because it's not just for debugging
;            Remove last modified date
; 8/1/2021 - Add update history
;            change Consoleout() to always return string written

#include-once
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <WinAPIError.au3>

Global $_LD_Debug = True
Global $_Profile_Map[0][3]

Func _DebugOff()
    $_LD_Debug = False
EndFunc

Func _DebugOn()
    $_LD_Debug = True
EndFunc

; Consoleout
; Automatically replaces $ to variables given
; Escape $ using $$
; Use \n for \r\n
Func c($v = "", $nl = True, $v1 = 0x0, $v2 = 0x0, $v3 = 0x0, _
                            $v4 = 0x0, $v5 = 0x0, $v6 = 0x0, _
                            $v7 = 0x0, $v8 = 0x0, $v9 = 0x0, $v10 = 0x0)
    If Not $_LD_Debug Then
        Return
    EndIf
    ; Preserve the original type
    If IsString($v) Then
        $v = StringReplace($v, "\n", @CRLF)
        If @NumParams > 2 Then
            $v = StringReplace($v, "$$", "@PH@")
            $v = StringReplace($v, "$", "@PH2@")
            For $i = 1 To @NumParams - 2
                ; Don't use Eval() to prevent breaking when compiled using stripper param /rm "rename variables"
                Switch ($i)
                    Case 1
                        $v = StringReplace($v, "@PH2@", $v1, 1)
                    Case 2
                        $v = StringReplace($v, "@PH2@", $v2, 1)
                    Case 3
                        $v = StringReplace($v, "@PH2@", $v3, 1)
                    Case 4
                        $v = StringReplace($v, "@PH2@", $v4, 1)
                    Case 5
                        $v = StringReplace($v, "@PH2@", $v5, 1)
                    Case 6
                        $v = StringReplace($v, "@PH2@", $v6, 1)
                    Case 7
                        $v = StringReplace($v, "@PH2@", $v7, 1)
                    Case 8
                        $v = StringReplace($v, "@PH2@", $v8, 1)
                    Case 9
                        $v = StringReplace($v, "@PH2@", $v9, 1)
                    Case 10
                        $v = StringReplace($v, "@PH2@", $v10, 1)
                EndSwitch
                If @extended = 0 Then ExitLoop
            Next
            $v = StringReplace($v, "@PH@", "$")
            $v = StringReplace($v, "@PH2@", "$")
        EndIf
    EndIf
    If $nl Then
        ConsoleWrite($v & @CRLF)
    Else
        ConsoleWrite($v)
    EndIf
    Return $v
EndFunc

; Consoleout Mock
Func cmock($v = "", $nl = True, $v1 = 0x0, $v2 = 0x0, $v3 = 0x0, _
                                 $v4 = 0x0, $v5 = 0x0, $v6 = 0x0, _
                                 $v7 = 0x0, $v8 = 0x0, $v9 = 0x0, $v10 = 0x0)
    Return $v
EndFunc

; Insert Variable
; Returns a string with all the given variables inserted into
; Use \n for newline char
Func iv($s = "", $v1 = 0x0, $v2 = 0x0, $v3 = 0x0, _
                 $v4 = 0x0, $v5 = 0x0, $v6 = 0x0, _
                 $v7 = 0x0, $v8 = 0x0, $v9 = 0x0, $v10 = 0x0)
    $s = StringReplace($s, "\n", @CRLF)
    If @NumParams > 1 Then
        $s = StringReplace($s, "$$", "@PH@")
        $s = StringReplace($s, "$", "@PH2@")
        For $i = 1 To @NumParams - 1
            ; Don't use Eval() to prevent breaking when compiled using stripper param /rm "rename variables"
            Switch ($i)
                Case 1
                    $s = StringReplace($s, "@PH2@", $v1, 1)
                Case 2
                    $s = StringReplace($s, "@PH2@", $v2, 1)
                Case 3
                    $s = StringReplace($s, "@PH2@", $v3, 1)
                Case 4
                    $s = StringReplace($s, "@PH2@", $v4, 1)
                Case 5
                    $s = StringReplace($s, "@PH2@", $v5, 1)
                Case 6
                    $s = StringReplace($s, "@PH2@", $v6, 1)
                Case 7
                    $s = StringReplace($s, "@PH2@", $v7, 1)
                Case 8
                    $s = StringReplace($s, "@PH2@", $v8, 1)
                Case 9
                    $s = StringReplace($s, "@PH2@", $v9, 1)
                Case 10
                    $s = StringReplace($s, "@PH2@", $v10, 1)
            EndSwitch
            If @extended = 0 Then ExitLoop
        Next
        $s = StringReplace($s, "@PH@", "$")
    EndIf
    Return $s
EndFunc

; Designed for NppExec Console
; Gets string before a newline from the console with a timeout
; Anything read after the newline in this function is discarded
Func cget($timeout = 2147483647)
    Local $s = ""
    Local $timer = TimerInit()
    While TimerDiff($timer) < $timeout
        $s = ConsoleRead()
        Local $pos = StringInStr($s, @LF, $STR_NOCASESENSEBASIC)
        If $pos <> 0 Then
            ConsoleRead()
            Return StringStripWS(StringLeft($s, $pos), $STR_STRIPTRAILING)
        EndIf
        Sleep(50)
    WEnd
    Return ""
EndFunc

; Consoleout Line
Func cl()
    If Not $_LD_Debug Then
        Return
    EndIf
    ConsoleWrite(@CRLF)
EndFunc

; Consoleout Variable
; Requires the name of variables without the $ as string
; Does not work when compiled using stripper param /rm "rename variables"
Func cv($nl = True, $v1 = 0x0, $v2 = 0x0, $v3 = 0x0, $v4 = 0x0, $v5 = 0x0, _
                        $v6 = 0x0, $v7 = 0x0, $v8 = 0x0, $v9 = 0x0, $v10 = 0x0)
    If Not $_LD_Debug Then
        Return
    EndIf
    Local $s = ""
    For $i = 1 To @NumParams - 1
        Switch ($i)
            Case 1
                $s &= "$" & $v1 & " = " & Eval($v1)
            Case 2
                $s &= "$" & $v2 & " = " & Eval($v2)
            Case 3
                $s &= "$" & $v3 & " = " & Eval($v3)
            Case 4
                $s &= "$" & $v4 & " = " & Eval($v4)
            Case 5
                $s &= "$" & $v5 & " = " & Eval($v5)
            Case 6
                $s &= "$" & $v6 & " = " & Eval($v6)
            Case 7
                $s &= "$" & $v7 & " = " & Eval($v7)
            Case 8
                $s &= "$" & $v8 & " = " & Eval($v8)
            Case 9
                $s &= "$" & $v9 & " = " & Eval($v9)
            Case 10
                $s &= "$" & $v10 & " = " & Eval($v10)
        EndSwitch
        If $i < @NumParams - 1 Then
            $s &= " | "
        EndIf
    Next
    If $nl Then
        $s &= @CRLF
    EndIf
    ConsoleWrite($s)
EndFunc

; Consoleout Array
; Set $out = False to get the string without printing, else returns the original array
Func ca(Const ByRef $a, $nl = True, $nlOnNewEle = False, $indentForNewEle = " ", $out = True)
    If Not IsArray($a) Then
        Return
    EndIf
    Local $dims = UBound($a, 0)
    Local $s = ""
    $s &= "{"
    ca_internal($s, $a, 1, $dims, "", $nlOnNewEle, $indentForNewEle)
    $s &= "}"
    If $nl Then
        $s &= @CRLF
    EndIf
    If $out Then
        ConsoleWrite($s)
        Return $a
    Else
        Return $s
    EndIf
EndFunc

Func ca_internal(ByRef $s, ByRef $a, $dim, $dims, $ref, $nlOnNewEle, $indentForNewEle)
    Local $count = UBound($a, $dim)
    If $dim = $dims Then
        Local $ele
        For $i = 0 To $count - 1
            $ele = Execute("$a" & $ref & "[" & $i & "]")
            Switch VarGetType($ele)
                Case "Double"
                    $s &= $ele
                    If Not IsFloat($ele) Then
                        $s &= ".0"
                    EndIf
                Case "String"
                    $s &= '"' & $ele & '"'
                Case "Array"
                    $s &= ca($ele, False, False, " ", False)
                Case "Map"
                    $s &= "Map"
                Case "Object"
                    $s &= ObjName($ele)
                Case "DLLStruct"
                    $s &= "Struct"
                Case "Keyword"
                    If IsKeyword($ele) = 2 Then
                        $s &= "Null"
                    Else
                        $s &= $ele
                    EndIf
                Case "Function"
                    $s &= FuncName($ele) & "()"
                Case "UserFunction"
                    $s &= FuncName($ele) & "()"
                Case Else
                    $s &= $ele
            EndSwitch
            If $i < $count - 1 Then
                $s &= "," & $indentForNewEle
            EndIf
        Next
    Else
        Local $indent = $indentForNewEle
        If $nlOnNewEle Then
            $indent = ""
            Local $indentBuf = $indentForNewEle
            Local $repeatCount = $dim
            While $repeatCount > 1
                If BitAND($repeatCount, 1) Then
                    $indent &= $indentBuf
                EndIf
                $indentBuf &= $indentBuf
                $repeatCount = BitShift($repeatCount, 1)
            WEnd
            $indent &= $indentBuf
        EndIf
        For $i = 0 To $count - 1
            If $nlOnNewEle Then
                $s &= @CRLF & $indent
            EndIf
            $s &= "["
            ca_internal($s, $a, $dim + 1, $dims, $ref & "[" & $i & "]", $nlOnNewEle, $indentForNewEle)
            If $nlOnNewEle And $dim + 1 < $dims Then
                $s &= @CRLF & $indent
            EndIf
            $s &= "]"
            If $i < $count - 1 Then
                $s &= "," & $indent
            EndIf
        Next
        If $nlOnNewEle And $dim = 1 Then
            $s &= @CRLF
        EndIf
    EndIf
EndFunc

; Consoleout Timerdiff
Func ct($t)
    c(TimerDiff($t))
EndFunc

; Consoleout Error
Func ce($e, $nl = True)
    If $nl Then
        ConsoleWrite("ERROR:" & $e & @CRLF)
    Else
        ConsoleWrite("ERROR:" & $e)
    EndIf
EndFunc

; Throws and exits if HotKeySet was failed (not every invalid hotkey is checked)
Func CheckedHotKeySet($key, $function = 0x0)
    If @NumParams = 1 Then
        If Not HotKeySet($key) Then
            Local $winError = _WinAPI_GetLastError()
            Throw("CheckedHotKeySet", _
                ($winError = 0) ? "Invalid or not registered hotkey " : _WinAPI_GetLastErrorMessage(), _
                'Hotkey: "' & $key & '"')
            Exit
        EndIf
        Return 1
    Else
        If Not HotKeySet($key, $function) Then
            Local $winError = _WinAPI_GetLastError()
            Local $functionType = IsFunc($function) ; Function reference
            If Not $functionType Then
                $functionType = IsFunc(Execute($function)) ; Function name or other
            EndIf
            Local $errorString = ""
            If $winError <> 0 Then
                $errorString = _WinAPI_GetLastErrorMessage()
            ElseIf $functionType = 2 Then
                $errorString = "Builtin function is not allowed"
            ElseIf $functionType = 0 Then
                $errorString = "Invalid function"
            Else
                $errorString = "Invalid hotkey"
            EndIf
            Throw("CheckedHotKeySet", _
                $errorString, _
                'Hotkey: "' & $key & '"', _
                'Function: "' & (IsFunc($function) ? FuncName($function) : $function) & '"')
            Exit
        EndIf
        Return 1
    EndIf
EndFunc

Func ArrayAdd(ByRef $a, $v)
	ReDim $a[UBound($a) + 1]
	$a[UBound($a) - 1] = $v
EndFunc

Func ArrayFind(ByRef $a, $v)
	For $i = 0 To UBound($a) - 1
		If $a[$i] = $v Then
			Return $i
		EndIf
	Next
	Return -1
EndFunc

Func Min($a, $b)
	Return $a < $b ? $a : $b
EndFunc

Func Max($a, $b)
	Return $a > $b ? $a : $b
EndFunc

; Throws an error msgbox, does not exit the script
Func Throw($funcName, $m1 = 0x0, $m2 = 0x0, $m3 = 0x0, $m4 = 0x0, $m5 = 0x0, _
                                 $m6 = 0x0, $m7 = 0x0, $m8 = 0x0, $m9 = 0x0, $m10 = 0x0)
    Local $s = "Exception on " & $funcName & "()"
    For $i = 1 To @NumParams - 1
        $s &= @CRLF & @CRLF
        Switch ($i)
            Case 1
                $s &= $m1
            Case 2
                $s &= $m2
            Case 3
                $s &= $m3
            Case 4
                $s &= $m4
            Case 5
                $s &= $m5
            Case 6
                $s &= $m6
            Case 7
                $s &= $m7
            Case 8
                $s &= $m8
            Case 9
                $s &= $m9
            Case 10
                $s &= $m10
        EndSwitch
    Next
    MsgBox($MB_ICONERROR + $MB_TOPMOST, StringTrimRight(@ScriptName, 4), $s)
EndFunc

; Profiler profile Add
Func pa($v)
    For $i = 0 to UBound($_Profile_Map, 1) - 1
        If $_Profile_Map[$i][0] = $v Then
            c("Profiler >> The profile name already exists: ""$""", 1, $v)
            Return
        EndIf
    Next
    ReDim $_Profile_Map[UBound($_Profile_Map, 1) + 1][3]
    $_Profile_Map[UBound($_Profile_Map, 1) - 1][0] = $v  ; Name
    $_Profile_Map[UBound($_Profile_Map, 1) - 1][1] = 0.0  ; Value
    $_Profile_Map[UBound($_Profile_Map, 1) - 1][2] = -1  ; Start time, -1 is not running
EndFunc

; Profiler profile Start
Func ps($v)
    For $i = 0 to UBound($_Profile_Map, 1) - 1
        If $_Profile_Map[$i][0] = $v Then
            If $_Profile_Map[$i][2] = -1 Then
                $_Profile_Map[$i][2] = TimerInit()
            Else
                c("Profiler >> The specified profile to start is already started: ""$""", 1, $v)
            EndIf
            Return
        EndIf
    Next
    c("Profiler >> The specified profile to start does not exist: ""$"", adding profile...", 1, $v)
    pa($v)
    ps($v)
EndFunc

; Profiler profile End
Func pe($v)
    For $i = 0 to UBound($_Profile_Map, 1) - 1
        If $_Profile_Map[$i][0] = $v Then
            If $_Profile_Map[$i][2] <> -1 Then
                $_Profile_Map[$i][1] += TimerDiff($_Profile_Map[$i][2])
                $_Profile_Map[$i][2] = -1
            Else
                c("Profiler >> The specified profile to end has not start: ""$""", 1, $v)
            EndIf
            Return
        EndIf
    Next
    c("Profiler >> The specified profile to end does not exist: ""$""", 1, $v)
EndFunc

; Profiler Print result
Func pp()
    c("Profiler >> Printing result ====================")
    For $i = 0 To UBound($_Profile_Map, 1) -1
        c("Profiler >> $ - $", 1, $i + 1, $_Profile_Map[$i][0])
        c("Profiler >> L $ ms", 1, Round($_Profile_Map[$i][1], 2))
    Next
    c("Profiler >> ====================================")
EndFunc

; Profiler Reset
Func pr()
    For $i = 0 To UBound($_Profile_Map, 1) -1
        $_Profile_Map[$i][1] = 0.0
        $_Profile_Map[$i][2] = -1
    Next
EndFunc