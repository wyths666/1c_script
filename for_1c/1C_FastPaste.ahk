#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; Горячая клавиша для запуска процесса (Ctrl+Shift+V)
^+v::
    ; Получаем данные из буфера обмена
    Clipboard := ""
    Send ^c
    ClipWait, 1
    
    if (ErrorLevel) {
        MsgBox, Не удалось получить данные из буфера обмена
        return
    }
    
    ; Разделяем строки на массив
    lines := StrSplit(Clipboard, "`n", "`r")
    
    if (lines.Length() = 0) {
        MsgBox, Буфер обмена пуст или не содержит текста
        return
    }
    
    ; Преобразуем в двумерный массив (каждая строка - массив из ячеек)
    global productData := []
    for i, line in lines {
        if (line != "") {
            ; Разделяем по табуляции
            cells := StrSplit(line, "`t")
            if (cells.Length() >= 2) {
                ; Добавляем в массив: [имя, количество]
                productData.Push([cells[1], cells[2]])
            }
        }
    }
    
    if (productData.Length() = 0) {
        MsgBox, Не найдено данных для вставки (ожидается 2 колонки: наименование и количество)
        return
    }
    
    global currentIndex := 1
    
    MsgBox, % "Готово! Загружено " productData.Length() " товаров.`nНажмите Alt+V для вставки следующего товара."
return

; Горячая клавиша для вставки следующего товара (Alt+V)
!v::
    if (!IsObject(productData) || productData.Length() = 0) {
        MsgBox, Нет товаров для вставки. Сначала скопируйте данные (Ctrl+Shift+V)
        return
    }
    
    ; Берем текущий товар (массив из двух элементов)
    currentProduct := productData[1]
    
    ; Вставляем наименование
    Clipboard := currentProduct[1]  ; первый элемент - имя
    Send ^v
    
    ; После вставки: Enter, Tab
    Sleep, 200
    Send {Enter}
    Sleep, 200
    Send {Tab}
    
    ; Вставляем количество
    Clipboard := currentProduct[2]  ; второй элемент - количество
    Send ^v
    
    ; После вставки: Down, Insert
    Sleep, 200
    Send {Down}
    Sleep, 200
    Send {Insert}
	Sleep, 200
	Send !v  ; Alt+V
    
    ; Удаляем обработанный товар из массива
    productData.RemoveAt(1)
    
    ; Показываем информацию о оставшихся товарах
    remainingCount := productData.Length()
    ToolTip, Осталось товаров: %remainingCount%
    SetTimer, RemoveToolTip, 1000
    
    ; Если товары закончились
    if (productData.Length() = 0) {
        SetTimer, ShowCompletionMessage, -500
    }
return

RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
return

ShowCompletionMessage:
    MsgBox, Все товары вставлены!
return

; Показать информацию о скрипте
^+i::
    MsgBox, Скрипт для последовательной вставки товаров из Excel`n`nCtrl+Shift+V - Загрузить данные из буфера`nAlt+V - Вставить следующий товар`n`nСтруктура: наименование (Enter, Tab) + количество (Down, Insert)
return