#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
#WinActivateForce

; === Глобальные переменные ===
global productNames := []

; === Ctrl+Shift+V: Загрузить список наименований из буфера (одна колонка) ===
^+v::
    productNames := []
    
    Clipboard := 
    SendInput ^c
    ClipWait, 2
    if (ErrorLevel) {
        MsgBox, 48, Ошибка, Не удалось получить данные из буфера обмена.
        return
    }

    text := Trim(Clipboard)
    lines := StrSplit(text, "`n", "`r")

    for i, line in lines {
        name := Trim(line)
        if (name != "") {
            productNames.Push(name)
        }
    }

    if (productNames.Length() == 0) {
        MsgBox, 48, Ошибка, Нет данных. Ожидается список наименований (по одному на строку).
        return
    }

    count := productNames.Length()
    MsgBox, 64, Готово, Загружено %count% наименований.`n`nНажмите Alt+V один раз — и все вставятся автоматически.
return


; === Alt+V: Запуск вставки по одному наименованию ===
!v::
    if (!IsObject(productNames) || productNames.Length() == 0) {
        MsgBox, 48, Внимание, Нет наименований для вставки.`nСначала Ctrl+Shift+V
        return
    }

    ; Запускаем таймер (первый элемент через 100 мс)
    SetTimer, InsertNextName, 100
return


; === Таймер: Вставить следующее наименование ===
InsertNextName:
    if (!IsObject(productNames) || productNames.Length() == 0) {
        SetTimer, InsertNextName, Off
        SetTimer, ShowCompletionMsg, -500
        return
    }

    name := productNames[1]
    productNames.RemoveAt(1)

    ; --- ВСТАВКА: Ctrl+V, Enter, Insert ---
    Clipboard := name
    SendInput ^v
    Sleep, 150
    SendInput {Enter}
    Sleep, 150
    SendInput {Insert}
    Sleep, 250  ; Дать 1С время на обработку (можно скорректировать)

    ; Таймер вызовется снова автоматически, если остались наименования
return


; === Сообщение об окончании ===
ShowCompletionMsg:
    MsgBox, 64, Готово, Все наименования успешно вставлены!
return


; === Справка: Ctrl+Shift+I ===
^+i::
    count := IsObject(productNames) ? productNames.Length() : 0
    MsgBox, 64, Скрипт — Вставка наименований в 1С,
    (
    • Ctrl+Shift+V — загрузить список наименований (по одному на строку)
    • Alt+V — НАЖМИТЕ ОДИН РАЗ → начнётся автоматическая вставка
    • Алгоритм: Ctrl+V → Enter → Insert (на каждое наименование)

    Наименований в очереди: %count%
    )
return


; === Остановка: Ctrl+Shift+X ===
^+x::
    SetTimer, InsertNextName, Off
    productNames := []
    MsgBox, 48, Остановлено, Скрипт остановлен. Очередь очищена.
return