Attribute VB_Name = "ModuleOfAdjustPrintBoundaries"
Option Explicit
'adjust the print boundaries
Sub AdjustPrintBoundaries(sheetName As String)
      Dim sheet As Worksheet
      'счетчик горизонтальных разрывов печати
      Dim nextPageBreakNumber As Integer
      Dim pageBreakFirstLine As Range
      Dim lineNumber As Long
      Dim lastFilledLine As Long

      'отключает UI
      Application.ScreenUpdating = False
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'задает лист для выравнивания
      Set sheet = ThisWorkbook.Worksheets(sheetName)
                   
      'последняя заполненная строка в документе
      lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
      sheet.Activate
      ActiveWindow.View = xlPageBreakPreview
      sheet.ResetAllPageBreaks
      nextPageBreakNumber = 1
      
      'добавление символа для корректной работы алгоритма (ссылка на документацию microsoft)
      sheet.Cells(sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row + 1, 1) = 1
      
      'HPageBreaks - каждый горизонтальный разрыв страницы печати
      'HPageBreaks(...).Location - возвращает или задает горизонтальный разрыв страницы (по верхнему краю ячейки (диапазона))
      While nextPageBreakNumber <= sheet.HPageBreaks.Count
            Set pageBreakFirstLine = sheet.HPageBreaks(nextPageBreakNumber).Location
            lineNumber = pageBreakFirstLine.Row
            
            If sheet.Cells(lineNumber, 1).MergeCells Then
                  If IsDate(sheet.Cells(sheet.Cells(lineNumber, 1).MergeArea.Cells(1, 1).Row - 1, 1)) Then
                        Set sheet.HPageBreaks(nextPageBreakNumber).Location = sheet.Cells(sheet.Cells(lineNumber, 1).MergeArea.Cells(1, 1).Row - 1, 1)
                  Else
                        Set sheet.HPageBreaks(nextPageBreakNumber).Location = sheet.Cells(sheet.Cells(lineNumber, 1).MergeArea.Row, 1)
                  End If
            Else
                  Set sheet.HPageBreaks(nextPageBreakNumber).Location = sheet.Cells(lineNumber, 1)
            End If
                        
            nextPageBreakNumber = nextPageBreakNumber + 1
      Wend
            
      'выделяет все границы черным
      sheet.Range("A7:S" & lastFilledLine).Borders.LineStyle = True
      
      'удаление спец. символа
      sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp) = ""
      
      'включает UI
      Application.ScreenUpdating = True
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле задания границ печати. Обратитесь к разработчику", vbCritical
End Sub
