^0::
; this ahk script replace manual line break with space, after copying text from pdf
;https://www.youtube.com/watch?v=ewCrO9jvcMg

; detect lang and switch to ENG
if InputLocaleID = 0x4090409 ; ENG
{
;msgbox ENG

}
else ;InputLocaleID = 0xF03D040D ; HEB
{
;msgbox HEB
SetInputLang(0x0409)
SetInputLang(Lang)
{
    WinExist("A")
    ControlGetFocus, CtrlInFocus
  PostMessage, 0x50, 0, % Lang, %CtrlInFocus%
}
}

send ^h
send +6
send l

send {tab}
send {space} ; fixing grammered error, happend by deleting breaking line
send {tab}
send {tab}
send {tab}
Sleep 100

send {Enter}

send {Enter}
Sleep 200
;https://www.autohotkey.com/board/topic/6383-send-esc-not-working-well/
SetKeyDelay, 10, 10
send {escape}
Send {Esc}
