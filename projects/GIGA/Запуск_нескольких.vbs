' Имена и полные пути к скриптам, которые нужно запустить
Const SCRIPT1_PATH = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\назначить_автоматы_АЗД.vbs"
Const SCRIPT2_PATH = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\назначить_автоматы_ЧП.vbs"
Const SCRIPT3_PATH = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\ВСЕ_атрибуты_в_ООО.vbs"

' Создаем объект WScript.Shell для запуска внешних программ
Set WshShell = CreateObject("WScript.Shell")

' Запускаем первый скрипт
WshShell.Run Chr(34) & SCRIPT1_PATH & Chr(34), 1, True
' Chr(34) используется для добавления кавычек вокруг пути,
' чтобы корректно обрабатывать пути с пробелами.
' 1 - стиль окна (нормальное окно)
' True - дождаться завершения выполнения скрипта

' Запускаем второй скрипт
WshShell.Run Chr(34) & SCRIPT2_PATH & Chr(34), 1, True

' Запускаем третий скрипт
WshShell.Run Chr(34) & SCRIPT3_PATH & Chr(34), 1, True

' Сообщаем о завершении
WScript.Echo "Все скрипты успешно запущены по очереди."

' Освобождаем объект
Set WshShell = Nothing