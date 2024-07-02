Attribute VB_Name = "ModuleOfFillingKeyDocumentaion"
Option Explicit
'Модуль заполнения XXXXXX XXXXXXXX
Sub FillingOfKeyDocumentaion()
Attribute FillingOfKeyDocumentaion.VB_ProcData.VB_Invoke_Func = "q\n14"
      
      'обработчик ошибок
      On Error GoTo ErrorHandler

      'variable
      
      'переменная программного листа
      Dim programmSheet As Worksheet
      Set programmSheet = ThisWorkbook.Worksheets("Программный лист")
      'последняя заполненная строка на листе
      Dim lastFilledLine As Integer
      lastFilledLine = ActiveSheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
      'номер первой ячейки заполненной строки на листе
      Dim startFilledLine As Integer
      startFilledLine = 11
      'переменная-флаг для корректной выдачи уведомлений пользователю
      Dim resultWork As Boolean
      resultWork = False
      
      'создание массивов для работы
      Dim arrayOfAllVulture As Variant
      Dim arrayOfNotAllVulture As Variant
      
      'массив для рандомной выборки массива ключей
      Dim arrRandom() As Variant
      ReDim arrRandom(1)
      
      'заносим данные в массивы
      arrayOfAllVulture = programmSheet.Range("B124:D" & programmSheet.Cells(Rows.Count, Range("B1").Column).End(xlUp).Row)
      arrayOfNotAllVulture = programmSheet.Range("E124:G" & programmSheet.Cells(Rows.Count, Range("G1").Column).End(xlUp).Row)
      
      arrRandom(0) = arrayOfAllVulture
      arrRandom(1) = arrayOfNotAllVulture
      
      'проход циклом по листу
      While startFilledLine <= lastFilledLine
       
            'ЕСЛИ ячейка с XXXXXX XXXXXXXX пустая
            'И имеется номер
            'И информация НЕ дата
            'И ЕСТЬ ФИО обрабатывавшего
            'И объем файла НЕ прочерк
            If ActiveSheet.Cells(startFilledLine, 7) = "" _
            And ActiveSheet.Cells(startFilledLine, 1) <> "" _
            And Not IsDate(ActiveSheet.Cells(startFilledLine, 1)) _
            And ActiveSheet.Cells(startFilledLine, 10) <> "" _
            And ActiveSheet.Cells(startFilledLine, 6) <> "-" Then
                  resultWork = True
                  Select Case getVulture(ActiveSheet.Cells(startFilledLine, 1))
                        Case "":
                              'получает случайный массив из массива, далее в полученном массиве получает случайный XXXXXX набор
                              ActiveSheet.Range("G" & startFilledLine & ":I" & startFilledLine) = WorksheetFunction.Index(arrRandom(WorksheetFunction.RandBetween(0, 1)), _
                              WorksheetFunction.RandBetween(1, UBound(arrayOfAllVulture) - LBound(arrayOfAllVulture) + 1))
                        Case "xxx":
                              'получает случайный массив из массива, далее в полученном массиве получает случайный XXXXXX набор
                              ActiveSheet.Range("G" & startFilledLine & ":I" & startFilledLine) = WorksheetFunction.Index(arrRandom(WorksheetFunction.RandBetween(0, 1)), _
                              WorksheetFunction.RandBetween(1, UBound(arrayOfAllVulture) - LBound(arrayOfAllVulture) + 1))
                        Case "x":
                              'получает случайный массив из массива, далее в полученном массиве получает случайный XXXXXX набор
                              ActiveSheet.Range("G" & startFilledLine & ":I" & startFilledLine) = WorksheetFunction.Index(arrRandom(WorksheetFunction.RandBetween(0, 1)), _
                              WorksheetFunction.RandBetween(1, UBound(arrayOfAllVulture) - LBound(arrayOfAllVulture) + 1))
                        Case "xx":
                              'получает случайный XXXXXX набор
                              ActiveSheet.Range("G" & startFilledLine & ":I" & startFilledLine) = WorksheetFunction.Index(arrayOfNotAllVulture, WorksheetFunction.RandBetween(1, UBound(arrayOfNotAllVulture) - LBound(arrayOfNotAllVulture) + 1))
                  End Select
            End If
                        
            'увеличивает счетчик строк excel документа
            startFilledLine = startFilledLine + 1
      Wend
            
      If resultWork Then
            MsgBox "XXXXXXXXXXXX на листе успешно заполнена", vbInformation, "Модуль заполнения XXXXXX XXXXXXXX"
      Else
            MsgBox "Незаполненная XXXXXXXXXXXX на листе отсутствует", vbInformation, "Модуль заполнения XXXXXX XXXXXXXX"
      End If
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле заполнения XXXXXX XXXXXXXX. Обратитесь к разработчику", vbCritical, "Модуль заполнения XXXXXX XXXXXXXX"
End Sub
