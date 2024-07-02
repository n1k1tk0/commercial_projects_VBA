Attribute VB_Name = "ModuleOfFindNumberException"
Option Explicit
'модуль поиска пропущенных номеров
Sub findNumberException(startRangeNumber As Long, endRangeNumber As Long)
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'variable
      
      Dim sheet As Variant
      'последняя заполненная строка на листе
      Dim lastFilledLine As Integer
      'первая рабочая строка на листе excel
      Dim startFilledLine As Integer
      startFilledLine = 11
      'массив листов
      Dim sheetArray() As Worksheet
      ReDim sheetArray(ThisWorkbook.Worksheets.Count - 2)
      'индекс массива
      Dim i As Integer
      i = 0
      'индекс массива значений
      Dim j As Long
      'массив всех значений ячеек первого столбца
      Dim cellsValueArray As Variant
      'значение ячейки в массиве
      Dim cellValue As Variant
      'текущий входящий номер
      Dim currentAccountingNumber As Variant
      'повторяющийся (искусственно) текущий номер для проверки на корректную последовательность нумерации
      Dim testingCurrentAccountingNumber As Variant
      testingCurrentAccountingNumber = startRangeNumber
      'количество за учетный период (последний номер в учетном периоде не считается)
      Dim countEncryptionTelegrams As Long
      countEncryptionTelegrams = endRangeNumber - startRangeNumber
      'коллекция проблемных номеров
      Dim problemTelegramsCollection As New Collection
      'коллекция всех номеров для проверки на уникальность
      Dim accountingNumberCollection As New Collection
      'счетчик количества
      Dim currentCountTelegrams As Long
      currentCountTelegrams = 0
      'проблемный номер
      Dim problemTelegram As Variant
      'формируемая строка с проблемными номерами (если они есть)
      Dim stringWithProblemTelegram As String
      stringWithProblemTelegram = ""
      
      'процедура предназначена для уменьшения количества кода
      'заполняет массив именами рабочих листов
      For Each sheet In ThisWorkbook.Worksheets
            If sheet.Name <> "Программный лист" Then
                  Set sheetArray(i) = sheet
                  i = i + 1
            End If
      Next
      
      'проход по всем листам в книге
      For Each sheet In sheetArray
            
            'заполнение массива всеми значениями с листов
            cellsValueArray = sheet.Cells(11, 1).Resize(sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row).Value
            
            'проход по массиву значений
            For Each cellValue In cellsValueArray
                        
                  'ЕСЛИ ячейка с номером НЕ пустая И ячейка с номером НЕ дата
                  If cellValue <> "" And Not IsDate(cellValue) Then
                        
                        'получение номера
                        currentAccountingNumber = getNumber(CStr(cellValue))
                        
                        'ЕСЛИ функция отбора номера вернуло значение
                        If currentAccountingNumber <> "" Then
                        
                              'ЕСЛИ полученный входящий номер входит в диапазон счетчика
                              If startRangeNumber <= currentAccountingNumber And currentAccountingNumber < endRangeNumber Then
                                    
                                    'проверка на корректную последовательность нумерации
                                    If CLng(testingCurrentAccountingNumber) <> CLng(currentAccountingNumber) Then
                                          MsgBox "Номер идет не по порядку: " & testingCurrentAccountingNumber & Chr(10) + Chr(10) _
                                          & "Выполнение будет остановлено. Перенесите номер на корректное место и повторите работу алгоритма.", vbCritical, "Модуль проверки нумерации"
                                          
                                           '----модуль переносит пользователя к "проблемному" номеру
                                          Dim sheetFind As Variant
                                          Dim rangeTest As Range
                                          For Each sheetFind In ThisWorkbook.Worksheets
                                                If Not (sheetFind.Range("A10:A" & sheetFind.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row).Find(testingCurrentAccountingNumber) Is Nothing) Then
                                                      Set rangeTest = sheetFind.Range("A10:A" & sheetFind.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row).Find(testingCurrentAccountingNumber)
                                                      sheetFind.Activate
                                                      rangeTest.Activate
                                                End If
                                          Next
                                          UserForm4.Hide
                                          '-------------------------------------------------------------------------------------
                                          
                                          Exit Sub
                                    End If
                                    
                                    'подсчет фактического количества в выгрузке
                                    currentCountTelegrams = currentCountTelegrams + 1
                                    
                                    'искусственная последовательность
                                    testingCurrentAccountingNumber = testingCurrentAccountingNumber + 1
                                    
                                    'добавление номера для проверки на актуальность
                                    On Error GoTo ErrorKey
                                    accountingNumberCollection.Add currentAccountingNumber, currentAccountingNumber
                              Else
                                    
                                    'подсчет фактического количества в выгрузке
                                    currentCountTelegrams = currentCountTelegrams + 1
                                    
                                    'добавление неучтенных номеров в коллекцию
                                    problemTelegramsCollection.Add currentAccountingNumber
                                    
                                    'добавление номера для проверки на актуальность
                                    On Error GoTo ErrorKey
                                    accountingNumberCollection.Add currentAccountingNumber, currentAccountingNumber
                              End If
                        End If
                  End If
            Next
      Next
      
      
      For Each problemTelegram In problemTelegramsCollection
            stringWithProblemTelegram = stringWithProblemTelegram & problemTelegram & ", "
      Next
      
      If stringWithProblemTelegram <> "" Then
            stringWithProblemTelegram = Left(stringWithProblemTelegram, Len(stringWithProblemTelegram) - 2)
            MsgBox "Полученное количество номеров с счетчика: " & countEncryptionTelegrams & Chr(10) & Chr(10) _
            & "Общее количество номеров, полученное алгоритмом: " & currentCountTelegrams & Chr(10) & Chr(10) _
            & "Номера, не вошедшие в данный номерный диапазон: " & stringWithProblemTelegram, vbInformation + vbDefaultButton1, "Модуль проверки нумерации"
      Else
            MsgBox "Полученное количество номеров с счетчика: " & countEncryptionTelegrams & Chr(10) & Chr(10) _
            & "Общее количество номеров, полученное алгоритмом: " & currentCountTelegrams, vbInformation + vbDefaultButton1, "Модуль проверки нумерации"
      End If
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле проверки нумерации. Обратитесь к разработчику", vbCritical, "Модуль проверки нумерации"
ErrorKey:
      MsgBox "Номер повторяется: " & currentAccountingNumber & ". Устраните неуникальность нумерации и повторите поиск", vbInformation, "Модуль проверки нумерации"
End Sub
