#Requires AutoHotkey v2.0
#SingleInstance Force

; Глобальные переменные для хранения данных
global g_orderNumber := ""     ; Номер заказа
global g_productName := ""     ; Наименование изделия
global g_quantity := ""        ; Количество изделий
global g_dueDate := ""         ; Дата готовности (в формате DD.MM.YY)
global g_customer := ""        ; заказчик
global g_addInfo := ""         ; Доп информация
global g_Path := ""            ; Путь к папке заказа

; Сочетания клавиш для сохранения параметров
^!1::SaveOrderNumber()        ; Ctrl+Alt+1 - сохранить номер заказа
^!2::SaveProductName()        ; Ctrl+Alt+2 - сохранить наименование
^!3::SaveQuantity()           ; Ctrl+Alt+3 - сохранить количество
^!4::SaveDueDate()            ; Ctrl+Alt+4 - сохранить дату готовности
^!5::SaveCustomer()           ; Ctrl+Alt+5 - сохранить заказчика
^!6::SaveAddInfo()            ; Ctrl+Alt+6 - сохранить доп. информацию
^!7::SavePath()               ; Ctrl+Alt+7 - сохранить путь к папке


SaveOrderNumber() {
    global g_orderNumber
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        text := Trim(A_Clipboard)
        if RegExMatch(text, "^\d{6}$") {
            g_orderNumber := text
            ToolTip "Заказ сохранен: " g_orderNumber
        } else {
            ToolTip "Некорректный номер заказа!"
        }
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}

SaveProductName() {
    global g_productName
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        text := Trim(A_Clipboard)
        g_productName := text
        ToolTip "Изделие сохранено: " SubStr(g_productName, 1, 30) (StrLen(g_productName) > 30 ? "..." : "")
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}

SaveQuantity() {
    global g_quantity
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        g_quantity := Trim(A_Clipboard)
        ToolTip "Количество сохранено: " g_quantity
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}

SaveDueDate() {
    global g_dueDate
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        text := Trim(A_Clipboard)
        
        ; преобразовать различные форматы дат
        if RegExMatch(text, "(\d{2})\.(\d{2})\.(\d{4})", &m) {
            ; Формат DD.MM.YYYY -> преобразуем в DD.MM.YY
            g_dueDate := m[1] "." m[2] "." SubStr(m[3], 3, 2)
            ToolTip "Дата сохранена: " g_dueDate
        }
        else if RegExMatch(text, "(\d{2})\.(\d{2})\.(\d{2})", &m) {
            ; Формат DD.MM.YY
            g_dueDate := text
            ToolTip "Дата сохранена: " g_dueDate
        }
        ;else if (dateValue := DateParse(text)) {
            ; Попытка автоматического разбора других форматов дат
        ;    FormatTime formattedDate, dateValue, "dd.MM.yy"
        ;    g_dueDate := formattedDate
        ;    ToolTip "Дата сохранена: " g_dueDate
        ;}
        else {
            ToolTip "Некорректный формат даты!"
        }
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}

; Функция для парсинга различных форматов дат
DateParse(str) {
    static formats := [
        "dd.MM.yyyy", "dd/MM/yyyy", "dd-MM-yyyy",
        "dd.MM.yy", "dd/MM/yy", "dd-MM-yy",
        "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
        "d.M.yyyy", "d/M/yyyy", "d-M-yyyy",
        "d.M.yy", "d/M/yy", "d-M-yy"
    ]
    
    loop formats.Length {
        try {
            return DateParseEx(str, formats[A_Index])
        } catch {
            continue
        }
    }
    return false
}

DateParseEx(str, fmt) {
    date := Date(str, fmt)
    if date = 16010101  ; Минимальная дата AHK
        return false
    return date
}

SaveCustomer() {
    global g_customer
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        text := Trim(A_Clipboard)
        g_customer := text
        ToolTip "Заказчик сохранен: " SubStr(g_customer, 1, 30) (StrLen(g_customer) > 30 ? "..." : "")
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}

SaveAddInfo() {
    global g_addInfo
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        text := Trim(A_Clipboard)
        g_addInfo := text
        ToolTip "Доп. инф. (открывание) сохранена: " SubStr(g_addInfo, 1, 30) (StrLen(g_addInfo) > 30 ? "..." : "")
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}

SavePath() {
    global g_Path
    try {
        ClipSaved := ClipboardAll()
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.5) {
            ToolTip "Ошибка копирования!"
            SetTimer () => ToolTip(), -1000
            return
        }
        
        text := Trim(A_Clipboard)
        g_Path := text
        ToolTip "Путь сохранен: " SubStr(g_Path, 1, 30) (StrLen(g_Path) > 30 ? "..." : "")
    } catch as e {
        ToolTip "Ошибка: " e.Message
    } finally {
        SetTimer () => ToolTip(), -1500
        A_Clipboard := ClipSaved
    }
}


:*:яяя:: {
    global g_orderNumber
    if g_orderNumber != "" {
        SendText g_orderNumber
    } else {
        SoundBeep 1500, 300
        ToolTip "Номер заказа не сохранен!"
        SetTimer () => ToolTip(), -1000
    }
}

:*:ннн:: {
    global g_productName
    if g_productName != "" {
        SendText g_productName
    } else {
        SoundBeep 1500, 300
        ToolTip "Наименование не сохранено!"
        SetTimer () => ToolTip(), -1000
    }
}

:*:ттт:: {
    global g_quantity
    if g_quantity != "" {
        SendText g_quantity
    } else {
        SoundBeep 1500, 300
        ToolTip "Количество не сохранено!"
        SetTimer () => ToolTip(), -1000
    }
}

:*:ддд:: {
    global g_dueDate
    if g_dueDate != "" {
        SendText g_dueDate
    } else {
        SoundBeep 1500, 300
        ToolTip "Дата не сохранена!"
        SetTimer () => ToolTip(), -1000
    }
}

:*:ззз:: {
    global g_customer
    if g_customer != "" {
        SendText g_customer
    } else {
        SoundBeep 1500, 300
        ToolTip "Заказчик не сохранен!"
        SetTimer () => ToolTip(), -1000
    }
}

:*:щщщ:: {
    global g_addInfo
    if g_addInfo != "" {
        SendText g_addInfo
    } else {
        SoundBeep 1500, 300
        ToolTip "Доп инфа не сохранена!"
        SetTimer () => ToolTip(), -1000
    }
}

:*:ппп:: {
    global g_Path
    if g_Path != "" {
        SendText g_Path
    } else {
        SoundBeep 1500, 300
        ToolTip "Путь не сохранен!"
        SetTimer () => ToolTip(), -1000
    }
}

; Просмотр сохраненных данных
^!0:: {
    global g_orderNumber, g_productName, g_quantity, g_dueDate, g_customer
    info := "ЗАКАЗ: " (g_orderNumber != "" ? g_orderNumber : "пусто") "`n"
          . "ИЗДЕЛИЕ: " (g_productName != "" ? SubStr(g_productName, 1, 40) : "пусто") "`n"
          . "КОЛИЧЕСТВО: " (g_quantity != "" ? g_quantity : "пусто") "`n"
          . "ДАТА ГОТОВНОСТИ: " (g_dueDate != "" ? g_dueDate : "пусто") "`n"
          . "ЗАКАЗЧИК: " (g_customer != "" ? SubStr(g_customer, 1, 40) : "пусто") "`n"
          . "ДОП. ИНФОРМАЦИЯ: " (g_addInfo != "" ? SubStr(g_addInfo, 1, 40) : "пусто") "`n"
          . "ПУТЬ К ПАПКЕ: " (g_Path != "" ? SubStr(g_Path, 1, 40) : "пусто") ; Новая строка

    MsgBox info, "Сохраненные параметры", "Iconi"
}

; Сброс параметров
^!9:: {
    global g_orderNumber := "", g_productName := "", g_quantity := "", g_dueDate := "", g_customer := "", g_addInfo := "", g_Path := ""
    ToolTip "Все параметры сброшены!"
    SetTimer () => ToolTip(), -1000
}