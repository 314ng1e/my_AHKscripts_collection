#Requires AutoHotkey v2.0

; Вставка текущей даты при вводе "сдн"
:*:сдн::
{
    CurrentDate := FormatTime(, "dd.MM.yy") 
    SendInput "{Backspace 3}" CurrentDate
}