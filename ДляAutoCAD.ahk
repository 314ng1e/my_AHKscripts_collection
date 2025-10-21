#Requires AutoHotkey v2.0
#SingleInstance


; Горячая клавиша работает ТОЛЬКО в AutoCAD
#HotIf WinActive("ahk_exe acad.exe") && GetKeyState("NumLock", "T")
NumpadDot::SendText(".")
#HotIf  ; Сбрасываем контекст