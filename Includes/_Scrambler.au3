; #INDEX# =======================================================================================================================
; Title .........: Simple Txt (un)Scrambler
; Script Version : 1.0.1
; AutoIt Version : 3.3.14.0 - Created with ISN AutoIt Studio v. 1.04
; Description ...: Takes a string and scrambles it, by creating an unicode Array, Random Multiply, Random letters and BitXOR
;                : The return string contains it's own "unscramble" keys, wich is used by the Unscramble function to reverse
;                : the scrambled string to the orginale string.
; Author(s) .....: Rex
; ===============================================================================================================================

; #CHANGE LOG# ==================================================================================================================
; Changes .......: Ascii -> Unicode, as jchd pointed out the script didn't work with "advance" unicode chars
; Date ..........: 12-02-2018
; ===============================================================================================================================

#include-once <Array.au3>
#include-once <String.au3>
#include-once <StringConstants.au3>

Func _Scramble($sMsg)
    ; Split the string
    Local $aMsgSplit = StringSplit($sMsg, '', $STR_NOCOUNT)
    ; Create a loop to get the Unicode char number
    Local $aiMsg[UBound($aMsgSplit)]
    For $i = 0 To UBound($aMsgSplit) - 1
        $aiMsg[$i] = AscW($aMsgSplit[$i])
    Next
    ; reverses the Array
    _ArrayReverse($aiMsg)

    ; Scramble the string
    Local $sSMsg, $iMultiplyWithFirst = Random(10, 99, 1), $iMultiplyWithLast = Random(100, 999, 1)

    ; Scramble the string by multiply some of the Ascii numbers with x and again others with y
    ; Looping thru the ascii array
    For $i = 0 To UBound($aiMsg) -1
        ; Multiplying some of the numbers
        If IsInt($i/3)  then ; If $i dived by 3 is an int. The divede by could be any number
            $aiMsg[$i] = $aiMsg[$i] * $iMultiplyWithFirst
        ElseIf IsInt($i /6) Then ; If $i diveded by 6 is int. The divede by could be any number
            $aiMsg[$i] = $aiMsg[$i] * $iMultiplyWithLast
        EndIf
        ; Creating the output string
        ; We do a bitXOR, and add Random letters to the sting
        $sSMsg &= BitXOR($aiMsg[$i],2) & __RandomChar()
    Next
    ; Combine and return scrambled string
    Return $iMultiplyWithFirst & $sSMsg & $iMultiplyWithLast
EndFunc

Func _DeScramble($sSMsg)
    ; Get the Multiplyer
    Local $iDivideWithFirst = StringLeft($sSMsg, 2)
    Local $iDivideWithLast = StringRight($sSMsg, 3)

    ;Removing the multiplyer
    $sSMsg = StringTrimLeft($sSMsg, 2)
    $sSMsg = StringTrimRight($sSMsg, 3)

    ; Replaces all Chars from the string with whitespaces
    Local $sNoChars = StringRegExpReplace($sSMsg, '\D', ' ')

    ; Removing 2+ whitspaces, stripping the last char from the string and splits the string into an array (Using the no count flag)
    Local $aSmsg = StringSplit(StringTrimRight(StringStripWS($sNoChars, $STR_STRIPSPACES), 1), ' ', $STR_NOCOUNT)

    ; Cleaning up the array
    For $i = 0 To UBound($aSmsg) -1
        ; Reversing the BitXOR
        $aSmsg[$i] = BitXOR($aSmsg[$i],2)
        ; Recalculating the orginal numbers
        If IsInt($i/3)  then
            $aSmsg[$i] = $aSmsg[$i] / $iDivideWithFirst
        ElseIf IsInt($i /6) Then
            $aSmsg[$i] = $aSmsg[$i] / $iDivideWithLast
        EndIf
    Next

    ; Now all we need is to change the numbers to chars, so we get the original and unscrambled string
    Local $sOrgMsg
    For $i = 0 To UBound($aSmsg) -1
            $sOrgMsg &= ChrW($aSmsg[$i])
    Next

    Return StringReverse($sOrgMsg)
EndFunc

Func __RandomChar()
    ; Internal use for Scramble function
    ; Generates a series of random chars (A-Z) in the range of 1 to 3 in a group
    Local $sRandChar
    For $i = 0 To Random(1, 3,1)
        $sRandChar &= ChrW(Random(65, 90, 1))
    Next
    Return $sRandChar
EndFunc
