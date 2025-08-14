
Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject()

' Создаём объект листа
Set Sheet = Job.CreateSheetObject()

' Параметры для создания листа
Dim moduleId, sheetName, symbolName, position, isBeforePosition
moduleId = 0                        ' Без модуля
sheetName = "House"                ' Имя листа
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
    
    ' Базовые координаты дома
    Dim baseX, baseY
    baseX = 100  ' Начало дома слева
    baseY = 100  ' Основание дома

    ' Размеры дома
    Dim houseWidth, houseHeight, roofHeight
    houseWidth = 80   ' Ширина дома
    houseHeight = 60  ' Высота стен
    roofHeight = 30   ' Высота крыши

    ' Цвета для разных элементов
    Const COLOR_WALLS = 1      ' Синий для стен
    Const COLOR_ROOF = 13      ' Красный для крыши
    
    ' Рисуем основание дома (прямоугольник синими линиями)
    result = Graphic.CreateLine(sheetId, baseX, baseY, baseX + houseWidth, baseY)                          ' Нижняя линия
    Graphic.SetLineColour COLOR_WALLS
    result = Graphic.CreateLine(sheetId, baseX, baseY, baseX, baseY + houseHeight)                         ' Левая стена
    Graphic.SetLineColour COLOR_WALLS
    result = Graphic.CreateLine(sheetId, baseX + houseWidth, baseY, baseX + houseWidth, baseY + houseHeight) ' Правая стена
    Graphic.SetLineColour COLOR_WALLS
    result = Graphic.CreateLine(sheetId, baseX, baseY + houseHeight, baseX + houseWidth, baseY + houseHeight) ' Верхняя линия
    Graphic.SetLineColour COLOR_WALLS

    ' Рисуем крышу
    Set RoofGraphic = Job.CreateGraphObject()
    
    ' Рисуем основание крыши
    result = RoofGraphic.CreateLine(sheetId, baseX, baseY + houseHeight, baseX + houseWidth, baseY + houseHeight)
    RoofGraphic.SetLineColour COLOR_ROOF
    
    ' Рисуем левый скат
    result = RoofGraphic.CreateLine(sheetId, baseX, baseY + houseHeight, baseX + houseWidth/2, baseY + houseHeight + roofHeight)
    RoofGraphic.SetLineColour COLOR_ROOF
    
    ' Рисуем правый скат
    result = RoofGraphic.CreateLine(sheetId, baseX + houseWidth/2, baseY + houseHeight + roofHeight, baseX + houseWidth, baseY + houseHeight)
    RoofGraphic.SetLineColour COLOR_ROOF
    
    ' Рисуем все линии крыши красным цветом
    result = RoofGraphic.CreateLine(sheetId, baseX, baseY + houseHeight, baseX + houseWidth, baseY + houseHeight)             ' Основание
    result = RoofGraphic.CreateLine(sheetId, baseX, baseY + houseHeight, baseX + houseWidth/2, baseY + houseHeight + roofHeight)  ' Левый скат
    result = RoofGraphic.CreateLine(sheetId, baseX + houseWidth/2, baseY + houseHeight + roofHeight, baseX + houseWidth, baseY + houseHeight)  ' Правый скат
    
    ' Очищаем объект
    Set RoofGraphic = Nothing

    ' Рисуем дверь (коричневым цветом)
    Dim doorWidth, doorHeight
    doorWidth = 20
    doorHeight = 30
    Dim doorX : doorX = baseX + 15  ' Положение двери от левого края
    Dim doorY : doorY = baseY       ' Дверь стоит на земле
    Graphic.SetColour 6  ' Коричневый
    result = Graphic.CreateLine(sheetId, doorX, doorY, doorX, doorY + doorHeight)                   ' Левая сторона двери
    result = Graphic.CreateLine(sheetId, doorX + doorWidth, doorY, doorX + doorWidth, doorY + doorHeight) ' Правая сторона двери
    result = Graphic.CreateLine(sheetId, doorX, doorY + doorHeight, doorX + doorWidth, doorY + doorHeight) ' Верх двери

    ' Рисуем окно (квадрат с крестом голубым цветом)
    Dim windowSize : windowSize = 20
    Dim windowX : windowX = baseX + houseWidth - 35  ' Положение окна от левого края
    Dim windowY : windowY = baseY + 20               ' Высота окна от земли
    result = Graphic.CreateLine(sheetId, windowX, windowY, windowX + windowSize, windowY)                     ' Нижняя линия окна
    result = Graphic.CreateLine(sheetId, windowX, windowY, windowX, windowY + windowSize)                     ' Левая линия окна
    result = Graphic.CreateLine(sheetId, windowX + windowSize, windowY, windowX + windowSize, windowY + windowSize) ' Правая линия окна
    result = Graphic.CreateLine(sheetId, windowX, windowY + windowSize, windowX + windowSize, windowY + windowSize) ' Верхняя линия окна
    ' Крест в окне
    result = Graphic.CreateLine(sheetId, windowX, windowY + windowSize/2, windowX + windowSize, windowY + windowSize/2) ' Горизонтальная линия
    result = Graphic.CreateLine(sheetId, windowX + windowSize/2, windowY, windowX + windowSize/2, windowY + windowSize) ' Вертикальная линия

    ' Рисуем дерево
    Dim treeX : treeX = baseX + houseWidth + 30  ' Положение дерева справа от дома
    Dim treeY : treeY = baseY                    ' Дерево растет из земли
    Dim trunkHeight : trunkHeight = 40           ' Высота ствола
    Dim trunkWidth : trunkWidth = 8             ' Ширина ствола
    Dim crownRadius : crownRadius = 20          ' Радиус кроны

    ' Ствол дерева (коричневым)
    Graphic.SetColour 6  ' Коричневый
    result = Graphic.CreateLine(sheetId, treeX, treeY, treeX, treeY + trunkHeight)                         ' Ствол
    result = Graphic.CreateLine(sheetId, treeX - trunkWidth/2, treeY, treeX + trunkWidth/2, treeY)         ' Основание ствола

    ' Крона дерева (три круга разного размера зеленым цветом)
    Graphic.SetColour 2  ' Зеленый
    result = Graphic.CreateCircle(sheetId, treeX, treeY + trunkHeight + crownRadius/2, crownRadius)        ' Нижний круг
    result = Graphic.CreateCircle(sheetId, treeX - crownRadius/2, treeY + trunkHeight + crownRadius, crownRadius-5)  ' Левый верхний
    result = Graphic.CreateCircle(sheetId, treeX + crownRadius/2, treeY + trunkHeight + crownRadius, crownRadius-5)  ' Правый верхний

    If result = 0 Then
        message = "Ошибка при создании графики"
    Else
        message = "Домик успешно нарисован"
    End If

    App.PutInfo 0, message
End If

Set Graphic = Nothing
Set Sheet = Nothing
Set Job = Nothing
Set App = Nothing
