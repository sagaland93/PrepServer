; ---------------------------------------
; Function _GetNewestFile()
;   Call with:  _GetNewestFile($DirPath [, $DateType [, $Recurse]])
;   Where:  $DirPath is the directory to search
;           $DateType (0ptional) is the type of date to use [from FileGetTime()]:
;               0 = Modified (default)
;               1 = Created
;               2 = Accessed
;           $Recurse (Optional): If non-zero causes the search to be recursive.
;   On success returns the full path to the newest file in the directory.
;   On failure returns 0 and sets @error (see code below).
; ---------------------------------------
Func _GetNewestFile($DirPath, $DateType = 0, $Recurse = 0)
    Local $Found, $FoundRecurse, $FileTime
    Local $avNewest[2] = [0, ""]; [0] = time, [1] = file
    
    If StringRight($DirPath, 1) <> '\' Then $DirPath &= '\'
    If Not FileExists($DirPath) Then Return SetError(1, 0, 0)
    
    Local $First = FileFindFirstFile($DirPath & '*.*')
    If $First = -1 Or @error Then Return SetError(2, @error, 0)
    
    While 1
        $Found = FileFindNextFile($First)
        If @error Then ExitLoop
        If StringInStr(FileGetAttrib($DirPath & $Found), 'D') Then
            If $Recurse Then
                $FoundRecurse = _GetNewestFile($DirPath & $Found, $DateType, 1)
                If @error Then
                    ContinueLoop
                Else
                    $Found = StringReplace($FoundRecurse, $DirPath, "")
                EndIf
            Else
                ContinueLoop
            EndIf
        EndIf
        $FileTime = FileGetTime($DirPath & $Found, $DateType, 1)
        If $FileTime > $avNewest[0] Then
            $avNewest[0] = $FileTime
            $avNewest[1] = $DirPath & $Found
        EndIf
    WEnd
    
    If $avNewest[0] = 0 Then
        Return SetError(3, 0, 0)
    Else
        Return $avNewest[1]
    EndIf
EndFunc   ;==>_GetNewestFile