Attribute VB_Name = "ModuleOfSearchDraftMaterial"
Option Explicit
'Модуль поиска XXXXX
Sub SearchDraftMaterial(Optional sheetName As Variant)

      'обработчик ошибок
      On Error GoTo ErrorHandler

      'variable
      
      'переменные для работы с Word
      Dim wordApp As Word.Application, wordDoc As Word.Document
      'переменная отражает фактическое создание процесса или подключение к уже созданному
      Dim newWordProccess As Boolean
      newWordProccess = False
      On Error Resume Next
      Set wordApp = GetObject(, "Word.Application")
      If wordApp Is Nothing Then
            Set wordApp = New Word.Application
            newWordProccess = True
      End If
      'переменная для работы с файловой системой
      Dim fso As FileSystemObject
      Set fso = New FileSystemObject
      'путь до рабочей директории
      Dim homeDir As String
      homeDir = ThisWorkbook.Path
      'последняя заполненная строка на листе
      Dim lastFilledLine As Integer
      'номер первой ячейки заполненной строки на листе
      Dim startFilledLine As Integer
      startFilledLine = 11
      'номер первой строки для заполнения в word документе
      Dim stringInWordDoc As Integer
      stringInWordDoc = 2
      'переменная, отражающая результат текущего поиска
      Dim isSearchCurrentResult As Boolean
      isSearchCurrentResult = False
      'массив листов
      Dim sheetArray() As Worksheet
      'индекс массива
      Dim i As Integer
      i = 0
      'массив для разделения каждого диапазона словаря
      Dim rangeArray() As String
      ReDim rangeArray(2)
      'индекс переборки полученного словаря
      Dim j As Double
      'индекс для отслеживания перехода алгоритма на следующий месяц
      Dim numberMonth As Integer
      numberMonth = 1
      'переменная-флаг для отображения перехода на следующий месяц
      Dim monthIsNow As Boolean
      monthIsNow = False
      Dim sheet As Variant
      'фиксирует первое открытие word документа алгоритмом
      Dim firstOpen As Boolean
      firstOpen = False
      'словарь для разделения диапазона номеров на листах
      Dim rangeNumbersDict As New Dictionary
      'горизонтальный разрыв страницы
      Dim horPageBreak As Integer
      'переменная для уменьшения количества отработки глобальных функций, записывает полученный номер в типе Long
      Dim numberWithDraftMaterial As Long
      
      'процедура предназначена для уменьшения количества кода
      'заполняет массив рабочими листками в зависимости от переданных условий
      If Not IsMissing(sheetName) Then
            ReDim sheetArray(0)
            Set sheetArray(i) = ThisWorkbook.Worksheets(sheetName)
      Else
            ReDim sheetArray(ThisWorkbook.Worksheets.Count - 2)
            For Each sheet In ThisWorkbook.Worksheets
                  If sheet.Name <> "Программный лист" Then
                        Set sheetArray(i) = sheet
                        i = i + 1
                  End If
            Next
      End If
      
      'перебор всех листов в массиве
      For Each sheet In sheetArray
            
            'если месяц не первый, то увеличивает счетчик месяцев
            If numberMonth <> 1 Then
                  monthIsNow = True
            End If
            
            'проход циклом по всем горизонтальным разделителям для определения их номерного диапазона
            For horPageBreak = 1 To sheet.HPageBreaks.Count
                  'ЕСЛИ это первый разрыв, значит значения ОТ нужно брать с 11 строки
                  If horPageBreak = 1 Then
                        'добавление диапазона номеров в словарь с ключом номера страницы
                        'правая граница берется на одну строчку выше (т.к. свойство location выдает по верхнему краю ячейки)
                        'также берется первая ячейка в merge диапазоне
                        rangeNumbersDict.Add horPageBreak, getNumber(sheet.Cells(11, 1)) & "..." & getNumber( _
                                                                                                                                                                             sheet.Cells( _
                                                                                                                                                                                                sheet.Cells( _
                                                                                                                                                                                                                  sheet.HPageBreaks.Item(horPageBreak).Location.Row - 1, 1).MergeArea.Row, 1))
                  Else
                        'проверка попадания разделителя на дату
                        If IsDate(sheet.Cells(sheet.HPageBreaks.Item(horPageBreak).Location.Row, 1)) Then
                              rangeNumbersDict.Add horPageBreak, getNumber( _
                                                                                                                  sheet.Cells( _
                                                                                                                                    sheet.HPageBreaks.Item(horPageBreak).Location.Row + 1, 1)) & "..." & _
                                                                                              getNumber( _
                                                                                                                  sheet.Cells( _
                                                                                                                                    sheet.Cells( _
                                                                                                                                                      sheet.HPageBreaks.Item(horPageBreak).Location.Row - 1, 1).MergeArea.Row, 1))
                        Else
                              rangeNumbersDict.Add horPageBreak, getNumber( _
                                                                                                                  sheet.Cells( _
                                                                                                                                     IIf( _
                                                                                                                                          sheet.HPageBreaks.Item(horPageBreak - 1).Location.MergeCells, _
                                                                                                                                          sheet.HPageBreaks.Item(horPageBreak - 1).Location.MergeArea.Row, _
                                                                                                                                          sheet.HPageBreaks.Item(horPageBreak - 1).Location.Row), 1)) & "..." & _
                                                                                              getNumber( _
                                                                                                                  sheet.Cells( _
                                                                                                                                    sheet.Cells( _
                                                                                                                                                      sheet.HPageBreaks.Item(horPageBreak).Location.Row - 1, 1).MergeArea.Row, 1))
                        End If
                  End If
            Next
            
            'номер последней заполненной строки на листе
            lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
            
            'проход циклом по листу
            While startFilledLine <= lastFilledLine
                        
                  'ЕСЛИ ячейка с XX НЕ пустая
                  'И ЕСЛИ имеется XX номер
                  If sheet.Cells(startFilledLine, 11) <> "" And sheet.Cells(startFilledLine, 1) <> "" Then
                        isSearchCurrentResult = True
                        
                        If monthIsNow Then
                              'добавляет название месяца в начале номеров
                              With wordDoc.Tables(1)
                                    .Rows.Add
                                    .cell(stringInWordDoc, 1).Merge .cell(stringInWordDoc, 3)
                                    .Rows.Item(stringInWordDoc).Range.Text = "" & sheet.Name
                                    .Rows.Item(stringInWordDoc).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                              End With
                        
                              'увеличивает счетчик строк word документа
                              stringInWordDoc = stringInWordDoc + 1
                              
                              monthIsNow = False
                        End If
                        
                        'получение номера страницы в журнале проверкой вхождения номера в каждый range листа
                        For j = 1 To rangeNumbersDict.Count
                        
                              'разеделение каждого диапазона словаря
                              rangeArray = Split(rangeNumbersDict.Item(j), "...", 2)
                              
                              numberWithDraftMaterial = CLng(getNumber(sheet.Cells(startFilledLine, 1)))
                              'проверка вхождения в range
                              If CLng(rangeArray(0)) <= numberWithDraftMaterial And numberWithDraftMaterial <= CLng(rangeArray(1)) Then
                                    Exit For
                              End If
                        Next
                        
                        'проверка наличия общей папки для отчетов (в случае отсутствия - ее создание)
                        If Not fso.FolderExists(homeDir & "\XXXXX") Then
                              fso.CreateFolder (homeDir & "\XXXXX")
                        End If
                        
                        'ЕСЛИ НЕ передано название месяца
                        If IsMissing(sheetName) Then
                              
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\XXXXX\0.Общий отчет.docx") Then
                                    
                                    'проверяет есть ли образец в программном репозитории
                                    If Not fso.FileExists(homeDir & "\Программные файлы\Образец для XXXXX за весь период.docx") Then
                                          MsgBox "Отсутствует образец файла в программном репозитории. Работа будет остановлена", vbCritical, "Модуль поиска XXXXX"
                                          Exit Sub
                                    End If
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для XXXXX за весь период.docx", homeDir & "\XXXXX\0.Общий отчет.docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\XXXXX\0.Общий отчет.docx", ReadOnly:=False)
                                    
                                    'добавляет название месяца в начале номеров
                                    With wordDoc.Tables(1)
                                          .Rows.Add
                                          .cell(stringInWordDoc, 1).Merge .cell(stringInWordDoc, 3)
                                          .Rows.Item(stringInWordDoc).Range.Text = "" & sheet.Name
                                          .Rows.Item(stringInWordDoc).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                                    End With
                                    
                                    'увеличивает счетчик строк word документа
                                    stringInWordDoc = stringInWordDoc + 1
                                    
                                    firstOpen = True
                                    
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\XXXXX\0.Общий отчет.docx.docx") _
                              And Not fileWordDocIsOpen(homeDir & "\XXXXX\0.Общий отчет.docx", firstOpen) Then
                                    MsgBox "Общий отчет уже существует. Удалите его и запустите программу снова", vbExclamation, "Модуль поиска XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    Exit Sub
                              End If
                        Else
                        
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\XXXXX\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx") Then
                                    
                                    'проверяет есть ли образец в программном репозитории
                                    If Not fso.FileExists(homeDir & "\Программные файлы\Образец для XXXXX месячный.docx") Then
                                          MsgBox "Отсутствует образец файла в программном репозитории. Работа будет остановлена", vbCritical, "Модуль поиска XXXXX"
                                          Exit Sub
                                    End If
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для XXXXX месячный.docx", _
                                    homeDir & "\XXXXX\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\XXXXX\" _
                                    & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx", ReadOnly:=False)
                                    
                                    firstOpen = True
                                    
                                    'подстановка месяца в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&month", ReplaceWith:=sheetName
                              
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\XXXXX\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx") _
                              And Not fileWordDocIsOpen(homeDir & "\XXXXX\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx", firstOpen) Then
                                    
                                    MsgBox "Отчет за " & sheetName & " уже существует. Удалите его и запустите программу снова", vbExclamation, "Модуль поиска XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        End If
                        
                        'высвобождает процессы worda и останавливает работу алгоритма
                        If Not firstOpen Then
                              'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                              If newWordProccess Then
                                    wordApp.Quit
                                    Set wordApp = Nothing
                              End If
                              Exit Sub
                        End If
                              
                        'добавляет строку в документе word
                        wordDoc.Tables(1).Rows.Add
                        
                        'вызов функции копирования данных в документ word
                        'передается ссылка на вордДок, вх. номер, нумерация, номер строки в ворде, номер листа в журнале
                        'для номера страницы применяется тернарный оператор
                        If IsMissing(sheetName) Then
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)), stringInWordDoc - (numberMonth + 1), stringInWordDoc, IIf(j Mod 2 = 0, j / 2 + 1 & "/2", j / 2 + 1.5 & "/1"))
                        Else
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)), stringInWordDoc - 1, stringInWordDoc, IIf(j Mod 2 = 0, j / 2 + 1 & "/2", j / 2 + 1.5 & "/1"))
                        End If
                        
                        'увеличивает счетчик строк word документа
                        stringInWordDoc = stringInWordDoc + 1
                  End If
        
                  'увеличивает счетчик строк excel документа
                  startFilledLine = startFilledLine + 1
            Wend
            
            'возвращает переменной первоначальное значение при переходе на новый лист
            startFilledLine = 11
            'увеличивает счетчик месяцев
            numberMonth = numberMonth + 1
            
            'удаляет из словаря все значения старого листа
            rangeNumbersDict.RemoveAll
      Next
            
      If isSearchCurrentResult Then
            
            'добавляет в последнюю строку word документа общее количество XXXXX
            With wordDoc.Tables(1)
                  .cell(stringInWordDoc, 1).Merge .cell(stringInWordDoc, 2)
                  
                  'если не передано название месяца
                  If IsMissing(sheetName) Then
                        .Rows.Last.Range.Text = "Общее количество: " & stringInWordDoc - (numberMonth + 1)
                  Else
                        .Rows.Last.Range.Text = "Общее количество: " & stringInWordDoc - 2
                  End If
                  
                  .Rows.Last.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With
            
            'закрывает word документ и процессы word
            wordDoc.Save
            wordDoc.Close
            Set wordDoc = Nothing
            If newWordProccess Then
                  wordApp.Quit
                  Set wordApp = Nothing
            End If
            
            If Not IsMissing(sheetName) Then
                  MsgBox "Общий отчет за " & sheetName & " сформирован", vbInformation, "Модуль поиска XXXXX"
            Else
                  MsgBox "Общий отчет за весь период сформирован", vbInformation, "Модуль поиска XXXXX"
            End If
      Else
            If Not IsMissing(sheetName) Then
                  MsgBox "Пропуски за " & sheetName & " отсутствуют", vbInformation, "Модуль поиска XXXXX"
            Else
                  MsgBox "Пропуски за весь период отсутствуют", vbInformation, "Модуль поиска XXXXX"
            End If
      End If
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле поиска XXXXX. Обратитесь к разработчику", vbCritical, "Модуль поиска XXXXX"
End Sub
