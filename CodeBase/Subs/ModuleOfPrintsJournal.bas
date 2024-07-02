Attribute VB_Name = "ModuleOfPrintsJournal"
Option Explicit
'Модуль печати
Sub PrintsJournal(sheetName As String, startPage As Integer, totalPage As Integer, Optional ByRef numberPage As Integer)

      'переменная листа
      Dim sheet As Worksheet
      Set sheet = ThisWorkbook.Worksheets(sheetName)
      'последняя строка на листе
      Dim lastFilledLine As Integer
      lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
      'переменная-флаг для указания характеристики нумерации
      Dim oddOrEven As Boolean
      oddOrEven = False
      'переменная-флаг для указания на изменения данных печати из-за учетного блока
      Dim accountingBlockPrint As Boolean
      accountingBlockPrint = False
      'переменная для печати
      Dim page As Integer
      'переменная для фиксации правого колонтитула
      Dim fixFooterRight As String
      'верхняя левая строка ячейки, на которой находится фигура
      Dim addressRowFigure As Integer
      'верхняя ячейка фигуры до удаления сквозных строк
      Dim topCellShape As Range
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'задает левые и правые границы отступов в зависимости от характеристики нумерации
      'четные
      If startPage Mod 2 = 0 Then
            oddOrEven = True
            With sheet.PageSetup
                  .LeftMargin = Application.CentimetersToPoints(0.5)
                  .RightMargin = Application.CentimetersToPoints(2.5)
            End With
      'НЕчетные
      Else
            oddOrEven = False
            With sheet.PageSetup
                  .LeftMargin = Application.CentimetersToPoints(2.5)
                  .RightMargin = Application.CentimetersToPoints(0.5)
            End With
      End If
      
      'печать отдельно четных и нечетных страниц
      'печать от начала к концу
      With sheet
            If startPage < totalPage Then
                  For page = startPage To totalPage Step 2
                  
                        'печать НЕчетных страниц
                        If Not oddOrEven Then
                        
                              'задает номер страницы и учетный номер в колонтитулах соотвественно
                              .PageSetup.RightHeader = "&""Times New Roman""&12" & " " & numberPage
                        
                              'увелечение счетчика программной нумерации
                              numberPage = numberPage + 1
                                                      
                              'ЕСЛИ печатается НЕчетная предпоследняя страница
                              If page = .PageSetup.Pages.Count - 1 Then
                              
                                    'ЕСЛИ на листе есть программная фигура
                                    If findTextBoxInSheet(sheet) Then
                                          'вызов обработки печати последних страниц
                                          Call SettingBlockDuringPrinting(sheet, page, numberPage - 1)
                                    Else
                                          'печать
                                          .PrintOut From:=page, To:=page
                                    End If
                              Else
                                    'печать
                                    .PrintOut From:=page, To:=page
                              End If
                              
                        'печать четных страниц
                        Else
                              
                              'ЕСЛИ печатается последняя четная страница
                              If page = .PageSetup.Pages.Count Then
                              
                                    'ЕСЛИ на листе есть программная фигура
                                    If findTextBoxInSheet(sheet) Then
                                          
                                          'вызов обработки печати последних страниц
                                          Call SettingBlockDuringPrinting(sheet, page)
                                    End If
                              Else
                                    'печать
                                    .PrintOut From:=page, To:=page
                              End If
                        End If
                  Next
            
            'печать от конца к началу или одного листа
            Else
                  For page = startPage To totalPage Step -2
                  
                        'печать НЕчетных страниц
                        If Not oddOrEven Then
                        
                              'задает номер страницы
                              .PageSetup.RightHeader = "&""Times New Roman""&12" & " " & numberPage
                        
                              'изменяет программный счетчик нумерации
                              numberPage = numberPage - 1
                        
                              'ЕСЛИ печатается НЕчетная предпоследняя страница
                              If page = .PageSetup.Pages.Count - 1 Then
                              
                                    'ЕСЛИ на листе есть программная фигура
                                    If findTextBoxInSheet(sheet) Then
                                    
                                          'вызов обработки печати последних страниц
                                          Call SettingBlockDuringPrinting(sheet, page, numberPage + 1)
                                    Else
                                    
                                          'печать
                                          .PrintOut From:=page, To:=page
                                    End If
                              Else
                              
                                    'печать
                                    .PrintOut From:=page, To:=page
                              End If

                        'печать четных страниц
                        Else
            
                              'ЕСЛИ печатается последняя четная страница
                              If page = .PageSetup.Pages.Count Then
                              
                                    'ЕСЛИ на листе есть программная фигура
                                    If findTextBoxInSheet(sheet) Then
                                          
                                          'вызов обработки печати последних страниц
                                          Call SettingBlockDuringPrinting(sheet, page)
                                    End If
                              Else
                        
                                    'печать
                                    .PrintOut From:=page, To:=page
                              End If
                        End If
                  Next
            End If
      End With
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле печати. Обратитесь к разработчику", vbCritical, "Модуль печати"
End Sub

