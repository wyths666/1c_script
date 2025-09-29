#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
#WinActivateForce

; === Глобальные переменные ===
global productData := []

; === Ctrl+Shift+V: Загрузить данные из буфера ===
^+v::
    ; Очистка предыдущих данных
    productData := []

    ; Копируем выделенное
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
        line := Trim(line)
        if (line = "")
            continue
        cells := StrSplit(line, "`t")
        if (cells.Length() >= 2) {
            name := Trim(cells[1])
            qty  := Trim(cells[2])
            if (name != "" && qty != "") {
                productData.Push([name, qty])
            }
        }
    }

    if (productData.Length() == 0) {
        MsgBox, 48, Ошибка, Нет данных. Ожидается: Наименование[TAB]Количество
        return
    }

    count := productData.Length()
    MsgBox, 64, Готово, Загружено %count% товаров.`n`nНажмите Alt+V один раз — и все вставятся автоматически.
return


; === Alt+V: Запуск автопроцесса (одно нажатие!) ===
!v::
    ; Проверяем, есть ли что вставлять
    if (!IsObject(productData) || productData.Length() == 0) {
        MsgBox, 48, Внимание, Нет товаров для вставки.`nСначала Ctrl+Shift+V
        return
    }

    ; Запускаем автоматическую вставку всех товаров
    SetTimer, AutoInsertProduct, 100  ; Первый товар через 100 мс
return


; === Таймер: Вставить следующий товар ===
AutoInsertProduct:
    ; Если товаров больше нет — завершаем
    if (!IsObject(productData) || productData.Length() == 0) {
        SetTimer, AutoInsertProduct, Off
        SetTimer, ShowCompletionMsg, -500
        return
    }

    ; Берём первый товар
    item := productData[1]
    name := item[1]
    qty  := item[2]

    ; Удаляем его сразу
    productData.RemoveAt(1)

    ; --- ВСТАВКА НАИМЕНОВАНИЯ ---
    Clipboard := name
    SendInput ^v
    Sleep, 200
    SendInput {Enter}
    Sleep, 150
    SendInput {Tab}
    Sleep, 200

    ; --- ВСТАВКА КОЛИЧЕСТВА ---
    Clipboard := qty
    SendInput ^v
    Sleep, 150

    ; --- ЗАВЕРШЕНИЕ СТРОКИ ---
    SendInput {Down}
    Sleep, 150
    SendInput {Insert}
    Sleep, 300  ; Даем 1С время обработать вставку строки

    ; --- ПОДГОТОВКА К СЛЕДУЮЩЕМУ ---
    ; Таймер сам вызовется снова, если остались товары
    ; Задержка между товарами — можно уменьшить при необходимости
return


; === Сообщение об окончании ===
ShowCompletionMsg:
    MsgBox, 64, Готово, Все товары успешно вставлены!
return


; === Справка: Ctrl+Shift+I ===
^+i::
    count := IsObject(productData) ? productData.Length() : 0
    MsgBox, 64, Скрипт — Автовставка в 1С,
    (
    • Ctrl+Shift+V — загрузить список
    • Alt+V — НАЖМИТЕ ОДИН РАЗ → начнётся автоматическая вставка всех товаров
    • Формат: Наименование[TAB]Количество

    Товаров в очереди: %count%
    )
return


; === Остановка: Ctrl+Shift+X ===
^+x::
    SetTimer, AutoInsertProduct, Off
    ToolTip
    productData := []
    MsgBox, 48, Остановлено, Скрипт остановлен. Очередь очищена.
return