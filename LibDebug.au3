; Update history
; 10/6/2024 - Add StringJoin()
; 9/12/2024 - Add ArrayContains(), StringStartsWith() and StringEndsWith()
; 6/19/2024 - Correct cget() to peek during the `ConsoleRead()` loop
; 3/19/2024 - Make CheckedHotKeySet() exit optionally
; 3/14/2024 - Add `Default` support for ca(), and castr() as a macro
;             for getting the array string
; 3/13/2024 - Add `Const` to ca() array parameter
;             Add ArrayAdd(), ArrayFind(), Min() and Max()
;             Add function list
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

; Functions:
; _DebugOff()
; _DebugOn()
; CheckedHotKeySet($key, $function = Default, $exit = True)
; ArrayAdd(ByRef $a, $v)
; ArrayFind(ByRef $a, $v)
; ArrayContains(ByRef $a, $v)
; StringStartsWith($s, $v, $case = True)
; StringEndsWith($s, $v, $case = True)
; StringJoin(ByRef $a)
; Min($a, $b)
; Max($a, $b)
; Throw($funcName, $msg1 ... $msg10)
;
; Consoleout             c($v = "", $nl = True, $v1 ... $v10)
; Consoleout Mock        cmock($v = "", $nl = True, $v1 ... $v10)
; Insert Variable        iv($s = "", $v1 ... $v10)
; Console Get            cget($timeoutMs = 2147483647)
; Consoleout Line        cl()
; Consoleout Variable    cv($nl = True, $v1 ... $v10)
; Consoleout Array       ca(Const ByRef $a, $nl = True, $nlOnNewEle = False, $indentForNewEle = " ", $out = True)
; Consoleout Timerdiff   ct($t)
; Consoleout Error       ce($e, $nl = True)
; Profiler profile Add   pa($key)
; Profiler profile Start ps($key)
; Profiler profile End   pe($key)
; Profiler Print result  pp()
; Profiler Reset         pr()

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

; Throws and exits (optionally) if HotKeySet() was failed (not every invalid hotkey is checked)
; Returns 1 if successful, or 0 if failed and $exit is set to False
Func CheckedHotKeySet($key, $func = Default, $exit = True)
    If $func = Default Then
        If Not HotKeySet($key) Then
            Local $winError = _WinAPI_GetLastError()
            Throw("CheckedHotKeySet", _
                ($winError = 0) ? "Invalid or not registered hotkey " : _WinAPI_GetLastErrorMessage(), _
                'Hotkey: "' & $key & '"')
            If $exit Then
                Exit
            EndIf
            Return 0
        EndIf
    Else
        If Not HotKeySet($key, $func) Then
            Local $winError = _WinAPI_GetLastError()
            Local $varType = IsFunc($func) ; Function reference
            Local $funcType = (Not $varType) ? IsFunc(Execute($func)) : $varType ; Function name or other types
            Local $errorString = ""
            If $winError <> 0 Then
                $errorString = _WinAPI_GetLastErrorMessage()
            ElseIf $funcType = 2 Then
                $errorString = "Builtin function is not allowed"
            ElseIf $funcType = 0 Then
                $errorString = "Invalid function"
            Else
                $errorString = "Invalid hotkey"
            EndIf
            Throw("CheckedHotKeySet", _
                $errorString, _
                'Hotkey: "' & $key & '"', _
                'Function: "' & ($varType ? FuncName($func) : $func) & '"')
            If $exit Then
                Exit
            EndIf
            Return 0
        EndIf
    EndIf
    Return 1
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

Func ArrayContains(ByRef $a, $v)
    Return ArrayFind($a, $v) <> -1
EndFunc

Func StringStartsWith($s, $v, $case = True)
    If $case Then
        Return StringLeft($s, StringLen($v)) == $v
    Else
        Return StringLeft($s, StringLen($v)) = $v
    EndIf
EndFunc

Func StringEndsWith($s, $v, $case = True)
    If $case Then
        Return StringRight($s, StringLen($v)) == $v
    Else
        Return StringRight($s, StringLen($v)) = $v
    EndIf
EndFunc

Func StringJoin(ByRef $a)
    Local $s = ""
    Local $len = UBound($a)
    For $i = 0 To $len - 1
        $s &= $a[$i]
        If $i < $len - 1 Then
            $s &= ","
        EndIf
    Next
    Return $s
EndFunc

Func Min($a, $b)
    Return $a < $b ? $a : $b
EndFunc

Func Max($a, $b)
    Return $a > $b ? $a : $b
EndFunc

; Throws an error msgbox, does not exit the script
Func Throw($funcName, $msg1 = 0x0, $msg2 = 0x0, $msg3 = 0x0, $msg4 = 0x0, $msg5 = 0x0, _
                      $msg6 = 0x0, $msg7 = 0x0, $msg8 = 0x0, $msg9 = 0x0, $msg10 = 0x0)
    Local $s = "Exception on " & $funcName & "()"
    For $i = 1 To @NumParams - 1
        $s &= @CRLF & @CRLF
        Switch ($i)
            Case 1
                $s &= $msg1
            Case 2
                $s &= $msg2
            Case 3
                $s &= $msg3
            Case 4
                $s &= $msg4
            Case 5
                $s &= $msg5
            Case 6
                $s &= $msg6
            Case 7
                $s &= $msg7
            Case 8
                $s &= $msg8
            Case 9
                $s &= $msg9
            Case 10
                $s &= $msg10
        EndSwitch
    Next
    MsgBox($MB_ICONERROR + $MB_TOPMOST, StringTrimRight(@ScriptName, 4), $s)
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

; Console Get
; Designed for NppExec Console
; Gets the string before a newline from the console with a timeout
; Anything read after the newline in this function is discarded
Func cget($timeoutMs = 2147483647)
    Local $s = ""
    Local $timer = TimerInit()
    While TimerDiff($timer) < $timeoutMs
        $s = ConsoleRead(True)
        If @error Then
            Return ""
        EndIf
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

; Consoleout Array String
Func castr(Const ByRef $a)
    Return ca($a, False, Default, Default, False)
EndFunc

; Consoleout Array
; Set $out = False to get the string without printing, else returns the original array
Func ca(Const ByRef $a, $nl = True, $nlOnNewEle = False, $indentForNewEle = " ", $out = True)
    If Not IsArray($a) Then
        Return
    EndIf
    If $nl = Default Then $nl = True
    If $nlOnNewEle = Default Then $nlOnNewEle = False
    If $indentForNewEle = Default Then $indentForNewEle = " "
    If $out = Default Then $out = True
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

Func ca_internal(ByRef $s, Const ByRef $a, $dim, $dims, $ref, $nlOnNewEle, $indentForNewEle)
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

; Profiler profile Add
Func pa($key)
    For $i = 0 to UBound($_Profile_Map, 1) - 1
        If $_Profile_Map[$i][0] = $key Then
            c("Profiler >> The profile name already exists: ""$""", 1, $key)
            Return
        EndIf
    Next
    ReDim $_Profile_Map[UBound($_Profile_Map, 1) + 1][3]
    $_Profile_Map[UBound($_Profile_Map, 1) - 1][0] = $key  ; Name
    $_Profile_Map[UBound($_Profile_Map, 1) - 1][1] = 0.0  ; Value
    $_Profile_Map[UBound($_Profile_Map, 1) - 1][2] = -1  ; Start time, -1 is not running
EndFunc

; Profiler profile Start
Func ps($key)
    For $i = 0 to UBound($_Profile_Map, 1) - 1
        If $_Profile_Map[$i][0] = $key Then
            If $_Profile_Map[$i][2] = -1 Then
                $_Profile_Map[$i][2] = TimerInit()
            Else
                c("Profiler >> The specified profile to start is already started: ""$""", 1, $key)
            EndIf
            Return
        EndIf
    Next
    c("Profiler >> The specified profile to start does not exist: ""$"", adding profile...", 1, $key)
    pa($key)
    ps($key)
EndFunc

; Profiler profile End
Func pe($key)
    For $i = 0 to UBound($_Profile_Map, 1) - 1
        If $_Profile_Map[$i][0] = $key Then
            If $_Profile_Map[$i][2] <> -1 Then
                $_Profile_Map[$i][1] += TimerDiff($_Profile_Map[$i][2])
                $_Profile_Map[$i][2] = -1
            Else
                c("Profiler >> The specified profile to end has not start: ""$""", 1, $key)
            EndIf
            Return
        EndIf
    Next
    c("Profiler >> The specified profile to end does not exist: ""$""", 1, $key)
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