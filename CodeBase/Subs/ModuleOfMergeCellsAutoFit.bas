Attribute VB_Name = "ModuleOfMergeCellsAutoFit"
Option Explicit
'Модуль выравнивния журнала (подбора высоты строк по содержимому ячеек)
Sub MergeCellsAutoFit(sheetName As String)
      'Отдельная ячейка
      Dim cell As Range
      'Диапазон объединения ячеек
      Dim MRng As Range
      'Высота верхней строки в диапазоне объединения
      Dim HRow1 As Double
      'Исходная высота по совокупности всех строк объединения
      Dim H1 As Double
      'Наименьшая необходимая высота для показа текста в объединённой ячейке
      Dim H2 As Double
      'Исходная ширина левого столбца в диапазоне объединения
      Dim WCol1 As Double
      'Исходная ширина по совокупности всех столбцов объединения
      Dim W1 As Double
      'счетчик
      Dim i As Integer
      Dim sheet As Worksheet
      'последняя заполненная строка на листе
      Dim lastFilledLine As Integer
      'переменные для задания ширины столбцов
      Dim j As Range
      Dim d As Integer
      'проверка даты
      Dim dat As Variant

      'Ускорение работы
      With Application
            .DisplayAlerts = False
            .ScreenUpdating = False
            .DisplayStatusBar = False
            .EnableEvents = False
      End With
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'задает лист для выравнивания
      Set sheet = ThisWorkbook.Worksheets(sheetName)
                       
      sheet.DisplayPageBreaks = False
      sheet.ResetAllPageBreaks
                  
      'добавление пункта перечня и действительного наименования вч
      sheet.Range("A2:S2").Value = "XXXXXXXX"
      'изменение заголовка журнала в зависимости от участка
      If UserForm2.CheckBox5 Then
            sheet.Range("A4:S4").Value = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      End If
      sheet.Range("A5:S5").Value = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      sheet.Range("A5:S5").Font.Color = vbBlack
                  
      'получение последней заполненной строки на листе
      lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
                  
      'отображает скрытый заголовок и копирует его
      ThisWorkbook.Worksheets("Программный лист").Range("K2:AC4").EntireColumn.Hidden = False
      ThisWorkbook.Worksheets("Программный лист").Range("K2:AC4").Copy
      ActiveSheet.Paste Destination:=sheet.Range("A7:S9")
                  
      'задаем нужную ширину столбцов
      d = 1
      For Each j In ThisWorkbook.Worksheets("Программный лист").Range("K1:AC1")
            sheet.Cells(1, d).ColumnWidth = j.ColumnWidth
            d = d + 1
      Next
                  
      'скрывает отредактированный заголовок с программного листа
      ThisWorkbook.Worksheets("Программный лист").Range("K2:AC4").EntireColumn.Hidden = True
                  
      'подбираем высоту для всех не merge ячеек
      sheet.Range("A1:S" & lastFilledLine).Rows.AutoFit
                  
      'переборка всего диапазона для выравнивания
      For Each cell In Application.Union(sheet.Range("E10:E" & lastFilledLine), sheet.Range("P10:Q" & lastFilledLine))
                        
            'выравнивание и задание даты дня
            If CBool(InStr(cell.Columns.Address, "E")) Then
                  If IsDate(cell) Then
                        dat = cell.Value
                        sheet.Range("A" & cell.Row & ":S" & cell.Row).MergeCells = False
                        sheet.Range("A" & cell.Row & ":S" & cell.Row).MergeCells = True
                        sheet.Range("A" & cell.Row & ":S" & cell.Row) = dat
                        sheet.Range("A" & cell.Row & ":S" & cell.Row).Font.Bold = True
                  End If
            End If
                        
            'Определяет диапазон объединения, в который входит ячейка
            Set MRng = cell.MergeArea
                        
            'ЕСЛИ ячейка принадлежит диапазону объединённых ячеек
            'И эта ячейка является левой верхней ячейкой в этом диапазоне
            If cell.MergeCells And (cell.Address = MRng.Cells(1, 1).Address) Then
                              
                  'Высота верхней строки в диапазоне объединения
                  HRow1 = MRng.Rows(1).RowHeight
                              
                  'Подсчитывает исходную высоту диапазона объединения по совокупности всех его строк
                  H1 = HRow1
                  If MRng.Rows.Count > 1 Then
                        For i = 2 To MRng.Rows.Count
                            H1 = H1 + MRng.Rows(i).RowHeight
                        Next i
                  End If
                                    
                  'Ширина левого столбца в диапазоне объединения
                  WCol1 = MRng.Columns(1).ColumnWidth
      
                  'Подсчитывает исходную ширину диапазона объединения по совокупности всех его столбцов
                  W1 = WCol1
                  If MRng.Columns.Count > 1 Then
                        For i = 2 To MRng.Columns.Count
                              W1 = W1 + MRng.Columns(i).ColumnWidth
                        Next i
                  End If
      
                  'Разъединяет ячейки (unmerge)
                  MRng.MergeCells = False
      
                  'Делает ширину левого столбца равной исходной ширине всего диапазона объединения
                  cell.ColumnWidth = W1
      
                  'Задает режим переноса текста по словам
                  cell.WrapText = True
      
                  'Выполняет подгон высоты верхней строки
                  cell.Rows.AutoFit
                              
                  'Выполняет замер получившейся высоты верхней строки. Это наименьшая высота, пригодная для показа текста
                  H2 = cell.Rows(1).RowHeight
      
                  'ЕСЛИ исходная высота диапазона объединения оказалась меньше, чем наименьшая пригодная высота
                  'ТО увеличиваем высоту верхней строки на соответствующую величину
                  If HRow1 > H2 Then
                        cell.Rows(1).RowHeight = HRow1
                  ElseIf H1 < H2 Then
                        cell.Rows(1).RowHeight = HRow1 + (H2 - H1)
                  End If
      
                  'Возвращает левому столбцу диапазона его прежнюю ширину
                  cell.ColumnWidth = WCol1
      
                  'Объединяет все нужные ячейки
                  MRng.MergeCells = True
            End If
      Next
      
      'работа с печатью
      With sheet.PageSetup
            'задает повторяющиеся строки заголовка
            .PrintTitleRows = "$7:$9"
            'задает формат бумаги
            .PaperSize = xlPaperA3
            'задает ориентацию печати
            .Orientation = xlLandscape
                        
            'задает верхние и нижние границы отступов
            .TopMargin = Application.CentimetersToPoints(1.3)
            .BottomMargin = Application.CentimetersToPoints(1.5)
                        
            'задает отступ от верхнего и нижнего краев листа до колонтитулов соотвественно
            .HeaderMargin = Application.CentimetersToPoints(0.5)
            .FooterMargin = Application.CentimetersToPoints(0.6)
                        
            'задает левые и правые отступы краев листа
            .LeftMargin = Application.CentimetersToPoints(2.5)
            .RightMargin = Application.CentimetersToPoints(0.5)
      End With
  
      'Восстанавливает UI эффекты
      sheet.DisplayPageBreaks = True
      With Application
            .DisplayAlerts = True
            .ScreenUpdating = True
            .DisplayStatusBar = True
            .EnableEvents = True
      End With
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле выравнивания. Обратитесь к разработчику", vbCritical
End Sub
