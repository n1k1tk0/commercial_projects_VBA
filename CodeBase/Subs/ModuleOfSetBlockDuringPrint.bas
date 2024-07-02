Attribute VB_Name = "ModuleOfSetBlockDuringPrint"
Option Explicit
'Модуль предназначен для обработки печати в runtime последних листов
Sub SettingBlockDuringPrinting(sheet As Worksheet, page As Integer, Optional numberPage As Integer)
      
      'номер верхнего разделителя страницы для переданной страницы
      Dim numberTopHPageBreaksInPage As Integer
      numberTopHPageBreaksInPage = page - 1
      'характеристика переданной страницы
      'true - четная; false - нечетная
      Dim oddOrEven As Boolean
      oddOrEven = page Mod 2 = 0
      'правый нижний колонтитул
      Dim footerRight As String
      footerRight = sheet.PageSetup.RightFooter
      'количество печатных страниц на листе
      Dim pagesCount As Integer
      pagesCount = sheet.PageSetup.Pages.Count
      'последняя заполненная строка на листе (в ней должна быть программная запись "Hello, world!"
      Dim lastFilledLine As Range
      Set lastFilledLine = sheet.Cells(Rows.Count, 1).End(xlUp)
      'последняя заполненная строка на листе в начале модуля
      Dim oldLastFilledLine As Long
      oldLastFilledLine = lastFilledLine.Row
      'верхняя строка фигуры до печати
      Dim addressOldRowFigure As Range
      Set addressOldRowFigure = sheet.Shapes.Item("programm figure").TopLeftCell
      'номер последней строки фигуры
      Dim lastFilledLineRow As Long
      
      'если первая строка после pagebreaks содержит данные, то спускаемся вниз, если строка не содержит данные, то лист пустой
      'получаем общее количество страниц на листе чтобы понять на нем фигура или нет
      'Если не на нем, значит печатается предпоследний нечетный лист ....
      'Если на нем, значит печатается последний четный лист
      
      'ЕСЛИ на листе есть программная фигура
      If findTextBoxInSheet(sheet) Then
            With sheet
                  Select Case oddOrEven
                        Case True:
                              'печать последнего четного листа не требует проверок, т.к. по дефолту программа не попадет сюда, если не будет выставлен учетный блок программными средствами
                              'убирает сквозные строки
                              .PageSetup.PrintTitleRows = ""
                              
                              'перемещает "съехавший" учетный блок на свое место
                              'ПОКА количество страниц меньше необходимого
                              While .PageSetup.Pages.Count < pagesCount + 1
                  
                                    'добавляет текст в ячейку для того, чтобы счетчик считал страницы
                                    With lastFilledLine
                                          .Value = "Hello, world!"
                                          .Font.Color = vbWhite
                                    End With
                                          
                                    'прыгает на одну ячейку вниз (имитация стрелки вниз)
                                    Set lastFilledLine = lastFilledLine.Offset(1)
                                    .PageSetup.PrintArea = lastFilledLine
                  
                                    'не дает ui умереть при выполнении
                                    DoEvents
                              Wend
                              
                              'прыгает на нужный лист (метод .Select не работает без этого)
                              .Select
                              'удаляет выделенную ячейку (т.к. цикл заканчивается на первой ячейке листа, не входящего в нужный диапазон)
                              lastFilledLine.Offset(-1).Delete
                                          
                              'новая последняя строка после удаления сквозных строк
                              Set lastFilledLine = .Cells(Rows.Count, 1).End(xlUp)
                              lastFilledLineRow = lastFilledLine.Row
                                          
                              'перемещает фонарик на нужное место
                              'делается отступ на 7 ячеек вверх (такое количество занимает фонарик)
                              .Shapes("programm figure").Left = .Range("A" & lastFilledLine.Offset(-7).Row).Left
                              .Shapes("programm figure").Top = .Range("A" & lastFilledLine.Offset(-7).Row).Top
                              
                              'удаляет все программные записи
                              .Range("A" & oldLastFilledLine & ":A" & lastFilledLine.Row).Delete
                              'удаляет первую запись
                              .Range("A" & oldLastFilledLine).Delete
                              'добавляет в последнюю ячейку последнего листа запись для корректной печати
                              With .Cells(lastFilledLineRow, 1)
                                    .Value = "Hello, world!"
                                    .Font.Color = vbWhite
                              End With
                              
                              'выравнивает фигуру (при удалении записей сбивается ширина фигуры)
                              .Shapes("programm figure").Width = Application.CentimetersToPoints(6.69)
                              
                              'печать
                              .PrintOut From:=page, To:=page
                              
                              'возвращает все обратно
                              .Shapes("programm figure").Left = addressOldRowFigure.Left
                              .Shapes("programm figure").Top = addressOldRowFigure.Top
                              .PageSetup.PrintTitleRows = "$7:$9"
                              
                              'прыгает обратно
                              ThisWorkbook.Worksheets("Программный лист").Select
                              
                        Case False:
                        
                              'ЕСЛИ ячейка под разделителем не merge И пустая, то печатается предпоследний нечетный лист
                              'в этом случае необходимо убрать: сквозные строки, правый нижний колонтитул и продолжить нумерацию
                              If Not CBool(.HPageBreaks(numberTopHPageBreaksInPage).Location.MergeCells) And .HPageBreaks(numberTopHPageBreaksInPage).Location.Value = "" And Not oddOrEven Then
                                    
                                    'убирает сквозные строки
                                    .PageSetup.PrintTitleRows = ""
                                    'убирает учетные данные от летника
                                    .PageSetup.RightFooter = ""
                                    'продолжает нумерацию
                                    .PageSetup.RightHeader = "&""Times New Roman""&12" & " " & numberPage
                                    
                                    'печать
                                    .PrintOut From:=page, To:=page
                                    
                                    'возвращает данные обратно
                                    .PageSetup.PrintTitleRows = "$7:$9"
                                    .PageSetup.RightFooter = footerRight
                                    
                              'в ином случае, если на листе есть данные
                              Else
                                    'продолжает нумерацию
                                    .PageSetup.RightHeader = "&""Times New Roman""&12" & " " & numberPage
                                    
                                    'печать
                                    .PrintOut From:=page, To:=page
                              End If
                  End Select
            End With
      End If
End Sub
