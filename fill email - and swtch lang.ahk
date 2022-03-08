^`::

; detect lang
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


send , elirantal1985@gmail.com
return

