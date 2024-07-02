Attribute VB_Name = "ModuleOfUnclosedDocumentation"
Option Explicit
'модуль поиска подвисших персональный
Sub CreateReportPersonalUnclosedDocumentation(fioIsResponsible As String, dispAlerts As Boolean, ByRef flag3 As Boolean, Optional sheetName As Variant)
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      Application.DisplayAlerts = dispAlerts
      Application.ScreenUpdating = dispAlerts

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
      startFilledLine = 9
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
      Dim sheet As Variant
      'фиксирует первое открытие word документа алгоритмом
      Dim firstOpen As Boolean
      firstOpen = False
      
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
      
      'перебор всех листов в книге
      For Each sheet In sheetArray
                  
            'номер последней заполненной строки на листе
            lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
            
            'проход циклом по листу
            While startFilledLine < lastFilledLine
                        
                  'основное условие ветвления
                  'ЕСЛИ ячейка с ФИО responsible НЕ пустая
                  'И ЕСЛИ ячейка с конечным состоянием пустая ИЛИ в ней нет цифр
                  'И ЕСЛИ ячейка не объединена (not merge) И ЕСЛИ ФИО responsible совпадает с выбранным
                  If sheet.Cells(startFilledLine, 16) <> "" And Not sheet.Cells(startFilledLine, 16).MergeCells And (sheet.Cells(startFilledLine, 17) = "" Or getNumber(sheet.Cells(startFilledLine, 17)) = "") _
                  And InStr(getFIONotGap(Trim(Mid(sheet.Cells(startFilledLine, 16), 12))), getFIONotGap(fioIsResponsible)) Then
                        isSearchCurrentResult = True
                        flag3 = True
                        
                        'создает основную папку по недостаткам, если она еще не создана
                        If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX") Then
                              fso.CreateFolder (homeDir & "\Незакрытые XXXXX")
                        End If
                              
                        'ЕСЛИ НЕ передано название месяца
                        If IsMissing(sheetName) Then
                              
                              'проверка наличия папки для отчетов за весь период (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX\Отчеты за весь период") Then
                                    fso.CreateFolder (homeDir & "\Незакрытые XXXXX\Отчеты за весь период")
                              End If
                        
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\" & fioIsResponsible & ".docx") Then
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для незакрытых XXXXX персональный за весь период.docx", _
                                    homeDir & "\Незакрытые XXXXX\Отчеты за весь период\" & fioIsResponsible & ".docx")
                                    
                                    'открывает скопированный файл и присваивает ему имя, в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\" & fioIsResponsible & ".docx", ReadOnly:=False)
                                    firstOpen = True
                                    
                                    'подставляет фио в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&responsible", ReplaceWith:=fioIsResponsible
                              
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\" & fioIsResponsible & ".docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\" & fioIsResponsible & ".docx", firstOpen) Then
                                    MsgBox "Отчет на " & fioIsResponsible & " за весь период уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль незакрытых XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        Else
                              
                              'проверка наличия папки для отчетов за выбранный месяц (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам") Then
                                    fso.CreateFolder (homeDir & "\Незакрытые XXXXX\Отчеты по месяцам")
                              End If
                              
                              If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName) Then
                                    fso.CreateFolder (homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName)
                              End If
                              
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx") Then
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для незакрытых XXXXX персональный месячный.docx", _
                                    homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx", ReadOnly:=False)
                                    firstOpen = True
                                    
                                    'подставляет фио в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&responsible", ReplaceWith:=fioIsResponsible
                                    'подставляет месяц в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&month", ReplaceWith:=sheetName
                              
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx", firstOpen) Then
                                    MsgBox "Отчет на " & fioIsResponsible & " за " & sheetName & " уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль незакрытых XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        End If
                        
                        'ЕСЛИ документ word не был открыт алгоритмом - высвобождает процессы worda и останавливает работу алгоритма
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
                                    
                        'проверка на merge и изъятие данных merga
                        'вызов функции копирования данных в документ word
                        If sheet.Cells(startFilledLine, 1).MergeCells Then
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1).MergeArea.Cells(1, 1)) & " / " & sheet.Cells(startFilledLine, 15), stringInWordDoc - 1, stringInWordDoc)
                        Else
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)) & " / " & sheet.Cells(startFilledLine, 15), stringInWordDoc - 1, stringInWordDoc)
                        End If
                              
                        'увеличивает счетчик строк word документа
                        stringInWordDoc = stringInWordDoc + 1
                  End If
        
                  'увеличивает счетчик строк excel документа
                  startFilledLine = startFilledLine + 1
            Wend
            
            'возвращает переменной первоначальное значение при переходе на новый лист
            startFilledLine = 9
      Next
        
      If isSearchCurrentResult Then
      
            'добавляет в последнюю строку word документа общее количество
            With wordDoc.Tables(1)
                  .cell(stringInWordDoc, 1).Merge .cell(stringInWordDoc, 3)
                  .Rows.Last.Range.Text = "Общее количество: " & stringInWordDoc - 2
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
                        
            If dispAlerts Then
                  If Not IsMissing(sheetName) Then
                        MsgBox "Отчет на " & fioIsResponsible & " за " & sheetName & " сформирован", vbInformation, "Модуль незакрытых XXXXX"
                  Else
                        MsgBox "Отчет на " & fioIsResponsible & " за весь период сформирован", vbInformation, "Модуль незакрытых XXXXX"
                  End If
            End If
      Else
            If dispAlerts Then
                  If Not IsMissing(sheetName) Then
                        MsgBox "Телеграммы на уточнении у " & fioIsResponsible & " за " & sheetName & " отсутствуют", vbInformation, "Модуль незакрытых XXXXX"
                  Else
                        MsgBox "Телеграммы на уточнении отсутствуют у " & fioIsResponsible & " за весь период отсутствуют", vbInformation, "Модуль незакрытых XXXXX"
                  End If
            End If
      End If
      
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True
End Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле поиска незакрытых XXXXX. Обратитесь к разработчику", vbCritical, "Модуль незакрытых XXXXX"
End Sub
'модуль поиска подвисших XXXXX общий
Sub CreateReportAllUnclosedDocumentation(Optional sheetName As Variant)
      
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
      startFilledLine = 9
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
      Dim sheet As Variant
      'фиксирует первое открытие word документа алгоритмом
      Dim firstOpen As Boolean
      firstOpen = False
    
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
      
      'перебор всех листов в книге
      For Each sheet In sheetArray
                  
            'номер последней заполненной строки на листе
            lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
            
            'проход циклом по листу
            While startFilledLine < lastFilledLine
                        
                  'основное условие ветвления
                  'ЕСЛИ фио responsible НЕ пустая
                  'И ЕСЛИ ячейка с конечным состоянием XXXXX пустая ИЛИ в ней нет цифр
                  'И ЕСЛИ ячейка НЕ merge (состояние подвисшей XXXXX)
                  If sheet.Cells(startFilledLine, 16) <> "" And Not sheet.Cells(startFilledLine, 16).MergeCells And (sheet.Cells(startFilledLine, 17) = "" Or getNumber(sheet.Cells(startFilledLine, 17)) = "") Then
                        isSearchCurrentResult = True
                        
                        'создает основную папку по недостаткам, если она еще не создана
                        If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX") Then
                              fso.CreateFolder (homeDir & "\Незакрытые XXXXX")
                        End If
                        
                        'ЕСЛИ НЕ передано название месяца
                        If IsMissing(sheetName) Then
                        
                              'проверка наличия папки для отчетов за весь период (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX\Отчеты за весь период") Then
                                    fso.CreateFolder (homeDir & "\Незакрытые XXXXX\Отчеты за весь период")
                              End If
                              
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\0.Общий отчет.docx") Then
                                          
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для незакрытых XXXXX общий за весь период.docx", _
                                    homeDir & "\Незакрытые XXXXX\Отчеты за весь период\0.Общий отчет.docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\0.Общий отчет.docx", ReadOnly:=False)
                                    firstOpen = True
                              
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\0.Общий отчет.docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Незакрытые XXXXX\Отчеты за весь период\0.Общий отчет.docx", firstOpen) Then
                                    MsgBox "Общий отчет за весь период уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль незакрытых XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        Else
                              
                              'проверка наличия папки для отчетов за выбранный месяц (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам") Then
                                    fso.CreateFolder (homeDir & "\Незакрытые XXXXX\Отчеты по месяцам")
                              End If
                              
                              If Not fso.FolderExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName) Then
                                    fso.CreateFolder (homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName)
                              End If
                              
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\0.Общий отчет.docx") Then
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для незакрытых XXXXX общий месячный.docx", _
                                    homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\0.Общий отчет.docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\0.Общий отчет.docx", ReadOnly:=False)
                                    firstOpen = True
                                    
                                    'подставляет месяц в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&month", ReplaceWith:=sheetName
                                    
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\0.Общий отчет.docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Незакрытые XXXXX\Отчеты по месяцам\Отчеты за " & sheetName & "\0.Общий отчет.docx", firstOpen) Then
                                    MsgBox "Общий отчет за " & sheetName & " уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль незакрытых XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        End If
                        
                        'ЕСЛИ документ word не был открыт алгоритмом - высвобождает процессы worda и останавливает работу алгоритма
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
                                    
                        'проверка на merge и изъятие данных merga
                        'вызов функции копирования данных в документ word
                        If sheet.Cells(startFilledLine, 1).MergeCells Then
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1).MergeArea.Cells(1, 1)) & " / " _
                              & sheet.Cells(startFilledLine, 15) & " / " & Mid(sheet.Cells(startFilledLine, 16), 12), stringInWordDoc - 1, stringInWordDoc)
                        Else
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)) & " - " & sheet.Cells(startFilledLine, 15), stringInWordDoc - 1, stringInWordDoc)
                        End If
                              
                        'увеличивает счетчик строк word документа
                        stringInWordDoc = stringInWordDoc + 1
                  End If
        
                  'увеличивает счетчик строк excel документа
                  startFilledLine = startFilledLine + 1
            Wend
            
            'возвращает переменной первоначальное значение при переходе на новый лист
            startFilledLine = 9
      Next
        
      If isSearchCurrentResult Then
      
            'добавляет в последнюю строку word документа общее количество XXXXX
            With wordDoc.Tables(1)
                  .cell(stringInWordDoc, 1).Merge .cell(stringInWordDoc, 2)
                  .Rows.Last.Range.Text = "Общее количество: " & stringInWordDoc - 2
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
                  MsgBox "Общий отчет за " & sheetName & " сформирован", vbInformation, "Модуль незакрытых XXXXX"
            Else
                  MsgBox "Общий отчет за весь период сформирован", vbInformation, "Модуль незакрытых XXXXX"
            End If
      Else
            If Not IsMissing(sheetName) Then
                  MsgBox "Телеграммы на уточнении за " & sheetName & " отсутствуют", vbInformation, "Модуль незакрытых XXXXX"
            Else
                  MsgBox "Телеграммы на уточнении за весь период отсутствуют", vbInformation, "Модуль незакрытых XXXXX"
            End If
      End If
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле поиска незакрытых XXXXX. Обратитесь к разработчику", vbCritical, "Модуль незакрытых XXXXX"
End Sub
