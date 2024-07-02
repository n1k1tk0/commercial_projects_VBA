Attribute VB_Name = "ModuleOfSpecialCounting"
Option Explicit

'общее количество экземпляров
Public totalNumberOfCopies As Long
'общее количество XXXXX
Public totalNumberOfHemmedCopies As Long
'общее количество XXXXX
Public totalNumberOfDestroyedCopies As Long
'общее количество XXXXX
Public totalNumberOfRepostedCopies As Long
'общее количество XXXXX
Public totalNumberOfCopiesSentIrrevocably As Long
'общее количество пXXXX
Public totalNumberOfCopiesPutOnInventory As Long
'Модуль подсчета для акта годовой
Sub SpecializedCounting()
      
      On Error GoTo ErrorHandler
      
      'result variable
      totalNumberOfCopies = 0
      totalNumberOfHemmedCopies = 0
      totalNumberOfDestroyedCopies = 0
      totalNumberOfRepostedCopies = 0
      totalNumberOfCopiesSentIrrevocably = 0
      totalNumberOfCopiesPutOnInventory = 0
      
      'переменная рабочего листа
      Dim sheet As Variant
      'массив листов
      Dim sheetArray() As Worksheet: ReDim sheetArray(ThisWorkbook.Worksheets.Count - 2)
      'индекс массива
      Dim i As Integer: i = 0
      
      'первая заполненная строка на листе
      Dim startFilledLine As Long
      'последняя заполненная строка на листе
      Dim lastFilledLine As Long
      
      'количество строк в объединенной ячейке количества экземпляров
      Dim rowsCountInMergeCell As Integer
      
      'процедура предназначена для уменьшения количества кода
      'заполняет массив именами рабочих листов
      For Each sheet In ThisWorkbook.Worksheets
            If sheet.Name <> "Программный лист" Then
                  Set sheetArray(i) = sheet
                  i = i + 1
            End If
      Next
      
      'номера, которые не вошли ни в одну группу
      Dim problemTelegramsCollection As New Collection
      'переменная-флаг для определения проблемных номеров
      Dim flag As Boolean
      
      'проход по всем листам в книге
      For Each sheet In sheetArray
            
            startFilledLine = 11
            
            'полученное последней заполненной строки на текущем листе
            lastFilledLine = sheet.Cells(Rows.Count, 1).End(xlUp).Row
            
            'проход циклом по всем строкам на листе
            While startFilledLine <= lastFilledLine
                  
                  'сброс флага
                  flag = False
                  
                  'ветвление для общего подсчета
                  'ЕСЛИ столбец 13 содержит данные
                  If sheet.Cells(startFilledLine, 13) <> "" Then
                        
                        'общее количество
                        totalNumberOfCopies = totalNumberOfCopies + CInt(sheet.Cells(startFilledLine, 13))
                        
                        'ЕСЛИ ячейка объединенная, значит в ней есть строка с электронным файлом, который учитывать не нужно
                        If sheet.Cells(startFilledLine, 13).MergeCells Then
                        
                              'получение количества строк в объединенной ячейке количества без учета первой строки
                              rowsCountInMergeCell = sheet.Cells(startFilledLine, 13).MergeArea.Rows.Count - 1
                              
                              'проход по всем строкам в объединенной ячейке количества
                              For i = 0 To rowsCountInMergeCell
                                    
                                    'ЕСЛИ 16 столбец ячейки содержит пометку XXXX И эта ячейка является объединенной
                                    If CBool(InStr(sheet.Cells(startFilledLine + i, 16), "xxxxxxx")) And sheet.Cells(startFilledLine + i, 16).MergeCells Then
                                          
                                          'проверяет была ли XXXXXX.
                                          'ЕСЛИ ДА, то добавляет в XXXXXX
                                          If CBool(InStr(sheet.Cells(startFilledLine + i, 18), "xxxxxxx")) Then
                                                totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                                flag = True
                                          
                                          'ЕСЛИ НЕТ, то добавляет в XXXXX
                                          Else
                                                totalNumberOfCopiesSentIrrevocably = totalNumberOfCopiesSentIrrevocably + 1
                                                flag = True
                                          End If
                                    
                                    'ЕСЛИ 16 столбец ячейки содержит "Реестр" И она объединена, НО НЕ содержит безвозвратно И 18 столбец содержит ВОЗВРАТ
                                    ElseIf CBool(InStr(sheet.Cells(startFilledLine + i, 16), "XXXX")) _
                                    And sheet.Cells(startFilledLine + i, 16).MergeCells _
                                    And CBool(InStr(sheet.Cells(startFilledLine + i, 18), "XXXX")) _
                                    And Not CBool(InStr(sheet.Cells(startFilledLine + i, 16), "XXXXX")) Then
                                          totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                          flag = True
                                    
                                    'ЕСЛИ 16 столбец ячейки НЕ содержит "XXXX" ИЛИ 17 столбец ячейки НЕ содержит "XXX"
                                    ElseIf Not (CBool(InStr(sheet.Cells(startFilledLine + i, 16), "1 фшт")) Or CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXX"))) Then
                                          
                                          'проверяет 17 столбец ячейки на определение конечного состояния
                                          'ЕСЛИ XXXX
                                          If CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXX")) Then
                                                totalNumberOfDestroyedCopies = totalNumberOfDestroyedCopies + 1
                                                flag = True
                                          
                                          'ЕСЛИ XXXX (множество условий исключает повторное добавление)
                                          ElseIf (CBool(InStr(sheet.Cells(startFilledLine + i, 17), "xxx")) Or InStr(sheet.Cells(startFilledLine + i, 17), "/")) _
                                          And Not (CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XX")) Or CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XX")) Or CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXXXX"))) Then
                                                totalNumberOfHemmedCopies = totalNumberOfHemmedCopies + 1
                                                flag = True
                                          
                                          'ЕСЛИ XXXXXX
                                          ElseIf CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXXXXXX")) Then
                                                
                                                'проверяет была ли XXXXX XXXXX XXX.
                                                'ЕСЛИ ДА, то добавляет в XXXXXX
                                                If CBool(InStr(sheet.Cells(startFilledLine + i, 18), "XXXXX.")) Then
                                                      totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                                      flag = True
                                                
                                                'ЕСЛИ НЕТ, то добавляет в XXXXXX
                                                Else
                                                      totalNumberOfCopiesSentIrrevocably = totalNumberOfCopiesSentIrrevocably + 1
                                                      flag = True
                                                End If
                                          
                                          'ЕСЛИ xxxxx
                                          ElseIf (CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXXX")) _
                                          Or CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXXXX")) _
                                          Or CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXXX"))) _
                                          And Not CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXXXX")) Then
                                                totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                                flag = True
                                          
                                          'ЕСЛИ поставлен на xxxxxxxx
                                          ElseIf CBool(InStr(sheet.Cells(startFilledLine + i, 17), "XXX")) Then
                                                totalNumberOfCopiesPutOnInventory = totalNumberOfCopiesPutOnInventory + 1
                                                flag = True
                                          End If
                                    End If
                              Next
                        
                        'ЕСЛИ ячейка не объединена
                        Else
                        
                              'ЕСЛИ 16 столбец ячейки содержит пометку XXXXX И эта ячейка является объединенной
                              If CBool(InStr(sheet.Cells(startFilledLine, 16), "xxxxxxx")) And sheet.Cells(startFilledLine, 16).MergeCells Then
                                          
                                    'проверяет была ли XXXXXX.
                                    'ЕСЛИ ДА, то добавляет в xxxxxxxx
                                    If CBool(InStr(sheet.Cells(startFilledLine, 18), "XXXXX")) Then
                                          totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                          flag = True
                                    
                                    'ЕСЛИ НЕТ, то добавляет в xxxxxxxx
                                    Else
                                          totalNumberOfCopiesSentIrrevocably = totalNumberOfCopiesSentIrrevocably + 1
                                          flag = True
                                    End If
                              
                               'ЕСЛИ 16 столбец ячейки содержит "XXXXX" И она объединена, НО НЕ содержит XXXX И 18 столбец содержит XXXXX
                              ElseIf CBool(InStr(sheet.Cells(startFilledLine, 16), "XXXX")) _
                              And sheet.Cells(startFilledLine, 16).MergeCells _
                              And CBool(InStr(sheet.Cells(startFilledLine, 18), "XXXX")) _
                              And Not CBool(InStr(sheet.Cells(startFilledLine, 16), "XXXXX")) Then
                                    totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                    flag = True
                                    
                              'ЕСЛИ 16 столбец ячейки НЕ содержит "XXXX" ИЛИ 17 столбец ячейки НЕ содержит "XXX"
                              ElseIf Not (CBool(InStr(sheet.Cells(startFilledLine, 16), "XXXX")) Or CBool(InStr(sheet.Cells(startFilledLine, 17), "XXX"))) Then
                                          
                                    'проверяет 17 столбец ячейки на определение конечного состояния
                                    'ЕСЛИ xxxxxxx
                                    If CBool(InStr(sheet.Cells(startFilledLine, 17), "XXXX")) Then
                                          totalNumberOfDestroyedCopies = totalNumberOfDestroyedCopies + 1
                                          flag = True
                                    
                                    'ЕСЛИ XXXXX (множество условий исключает повторное добавление переданных XXXXX)
                                    ElseIf (CBool(InStr(sheet.Cells(startFilledLine, 17), "xxx")) Or InStr(sheet.Cells(startFilledLine, 17), "/")) _
                                    And Not (CBool(InStr(sheet.Cells(startFilledLine, 17), "xx")) Or CBool(InStr(sheet.Cells(startFilledLine, 17), "XX")) Or CBool(InStr(sheet.Cells(startFilledLine, 17), "XXXX"))) Then
                                          totalNumberOfHemmedCopies = totalNumberOfHemmedCopies + 1
                                          flag = True
                                          
                                    'ЕСЛИ отправлен XXXX
                                    ElseIf CBool(InStr(sheet.Cells(startFilledLine, 17), "xxxxxx")) Then
                                                
                                          'проверяет была ли XXXXX.
                                          'ЕСЛИ ДА, то добавляет в xxxxxx
                                          If CBool(InStr(sheet.Cells(startFilledLine, 18), "XXXXXXX.")) Then
                                                totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                                flag = True
                                          
                                          'ЕСЛИ НЕТ, то добавляет в XXXXXX
                                          Else
                                                totalNumberOfCopiesSentIrrevocably = totalNumberOfCopiesSentIrrevocably + 1
                                                flag = True
                                          End If
                                    
                                    'ЕСЛИ перерег.
                                    ElseIf (CBool(InStr(sheet.Cells(startFilledLine, 17), "xxxxxxx")) _
                                    Or CBool(InStr(sheet.Cells(startFilledLine, 17), "XXXXXXX")) _
                                    Or CBool(InStr(sheet.Cells(startFilledLine, 17), "XXXXXXX"))) _
                                    And Not CBool(InStr(sheet.Cells(startFilledLine, 17), "xxxxxxx")) Then
                                          totalNumberOfRepostedCopies = totalNumberOfRepostedCopies + 1
                                          flag = True
                                          
                                    'ЕСЛИ поставлен на xxxxx
                                    ElseIf CBool(InStr(sheet.Cells(startFilledLine, 17), "XXX")) Then
                                          totalNumberOfCopiesPutOnInventory = totalNumberOfCopiesPutOnInventory + 1
                                          flag = True
                                    End If
                              End If
                        End If
                        
                        If Not flag Then
                              problemTelegramsCollection.Add getNumber(sheet.Cells(startFilledLine, 1))
                        End If
                        
                  End If
                  
                  startFilledLine = startFilledLine + 1
            Wend
      Next
      
      
      'если есть проблемные номера
      If problemTelegramsCollection.Count > 0 Then
            Dim problemTelegram As Variant
            Dim stringWithProblemTelegram As String
            
            For Each problemTelegram In problemTelegramsCollection
                  stringWithProblemTelegram = stringWithProblemTelegram & problemTelegram & ", "
            Next
            
            stringWithProblemTelegram = Left(stringWithProblemTelegram, Len(stringWithProblemTelegram) - 2)
            
            MsgBox "Номера, требующие внимания: " & stringWithProblemTelegram, vbExclamation, "Модуль специального подсчета"
      End If
Exit Sub

ErrorHandler:
      MsgBox "Произошла ошибка в модуле специального подсчета. Обратитесь к разработчику", vbCritical, "Модуль специального подсчета"
End Sub
