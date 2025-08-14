
Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject()

' Создаём объект листа
Set Sheet = Job.CreateSheetObject()

' Параметры для создания листа
Dim moduleId, sheetName, symbolName, position, isBeforePosition
moduleId = 0                        ' Без модуля
sheetName = "Car"                  ' Имя листа
symbolName = "Формат_А3_гор_1лист" ' Формат листа
position = 0                       ' В конец проекта
isBeforePosition = 0              ' После указанной позиции

' Создаём новый лист
Dim sheetId, result, message
result = Sheet.Create(moduleId, sheetName, symbolName, position, isBeforePosition)

' Проверяем успешность создания
If result > 0 Then
    sheetId = result
    
    ' Создаем объект для рисования
    Set Graphic = Job.CreateGraphObject()
    
    ' Базовые координаты автомобиля
    Dim baseX, baseY
    baseX = 100  ' Начало автомобиля слева
    baseY = 100  ' Основание автомобиля

    ' Размеры автомобиля
    Dim carWidth, carHeight, wheelRadius
    carWidth = 120    ' Длина автомобиля
    carHeight = 40    ' Высота кузова
    wheelRadius = 12  ' Радиус колес

    ' Базовые точки для кузова
    Dim frontBottom : frontBottom = baseX + 20        ' Передняя нижняя точка
    Dim rearBottom : rearBottom = baseX + carWidth - 20  ' Задняя нижняя точка
    Dim bodyHeight : bodyHeight = baseY + 15          ' Высота основного кузова

    ' Рисуем основной кузов (нижняя часть)
    result = Graphic.CreateLine(sheetId, frontBottom, baseY, rearBottom, baseY)                ' Нижняя линия
    result = Graphic.CreateLine(sheetId, baseX, bodyHeight, frontBottom, baseY)               ' Передний скос
    result = Graphic.CreateLine(sheetId, rearBottom, baseY, baseX + carWidth, bodyHeight)     ' Задний скос
    
    ' Рисуем верхнюю часть кузова (кабину)
    Dim cabinStartX : cabinStartX = baseX + 30
    Dim cabinWidth : cabinWidth = 50
    ' Убедимся, что стойки начинаются от линии кузова
    result = Graphic.CreateLine(sheetId, cabinStartX, bodyHeight, cabinStartX, baseY + carHeight)         ' Передняя стойка
    result = Graphic.CreateLine(sheetId, cabinStartX + cabinWidth, bodyHeight, cabinStartX + cabinWidth, baseY + carHeight) ' Задняя стойка
    result = Graphic.CreateLine(sheetId, cabinStartX, baseY + carHeight, cabinStartX + cabinWidth, baseY + carHeight) ' Крыша
    
    ' Рисуем окна
    ' Лобовое стекло (наклонное) - начинаем от стойки кабины
    result = Graphic.CreateLine(sheetId, cabinStartX - 10, bodyHeight + 5, cabinStartX, baseY + carHeight)
    ' Заднее стекло (наклонное) - начинаем от крыши
    result = Graphic.CreateLine(sheetId, cabinStartX + cabinWidth, baseY + carHeight, cabinStartX + cabinWidth + 10, bodyHeight + 5)

    ' Рисуем колеса (передние и задние)
    ' Переднее колесо - выравниваем с передним скосом
    result = Graphic.CreateCircle(sheetId, frontBottom, baseY, wheelRadius)
    result = Graphic.CreateCircle(sheetId, frontBottom, baseY, wheelRadius - 4) ' Внутренний круг
    ' Заднее колесо - выравниваем с задним скосом
    result = Graphic.CreateCircle(sheetId, rearBottom, baseY, wheelRadius)
    result = Graphic.CreateCircle(sheetId, rearBottom, baseY, wheelRadius - 4) ' Внутренний круг

    ' Рисуем фары
    Dim headlightRadius : headlightRadius = 5
    ' Передняя фара - привязываем к переднему скосу
    result = Graphic.CreateCircle(sheetId, baseX + 5, bodyHeight, headlightRadius)
    ' Задняя фара - привязываем к заднему скосу
    result = Graphic.CreateCircle(sheetId, baseX + carWidth - 5, bodyHeight, headlightRadius)

    ' Рисуем бампера - выравниваем с фарами
    result = Graphic.CreateLine(sheetId, baseX - 5, bodyHeight, baseX + 10, bodyHeight)     ' Передний бампер
    result = Graphic.CreateLine(sheetId, baseX + carWidth - 10, bodyHeight, baseX + carWidth + 5, bodyHeight) ' Задний бампер

    ' Рисуем декоративные элементы
    ' Решетка радиатора
    For i = 0 To 3
        result = Graphic.CreateLine(sheetId, baseX + 5, baseY + 12 + i*3, baseX + 15, baseY + 12 + i*3)
    Next

    ' Ручка двери
    result = Graphic.CreateLine(sheetId, cabinStartX + 30, baseY + 25, cabinStartX + 40, baseY + 25)

    If result = 0 Then
        message = "Ошибка при создании графики"
    Else
        message = "Автомобиль успешно нарисован"
    End If

    App.PutInfo 0, message
End If

Set Graphic = Nothing
Set Sheet = Nothing
Set Job = Nothing
Set App = Nothing
