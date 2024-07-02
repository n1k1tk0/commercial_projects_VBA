Attribute VB_Name = "ModuleOfKeyDocumentaion"
Option Explicit
'Модуль поиска незаполненной документации персональный
Sub CreateReportOfKeyDocumentation(fioIsResponsible As String, dispAlerts As Boolean, ByRef isSearchResultAllPersonalReport As Boolean, Optional sheetName As Variant)
      
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
      'первая рабочая строка на листе excel
      Dim startFilledLine As Integer
      startFilledLine = 9
      'первая рабочая строка в word документе
      Dim stringInWordDoc As Integer
      stringInWordDoc = 2
      'переменная, отражающая результат текущего поиска
      Dim isSearchCurrentResult As Boolean
      isSearchCurrentResult = False
      'переменная, отражающая хотя бы один положительный результат поиска
      Dim isSearchResult As Boolean
      isSearchResult = False
      'массив листов
      Dim sheetArray() As Worksheet
      'индексация массива
      Dim i As Integer
      i = 0
      'рабочий лист
      Dim sheet As Variant
      'фиксирует первое открытие алгоритмом
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
      
      'перебор всех листов в массиве
      For Each sheet In sheetArray
      
            'последняя заполненная строка на листе
            lastFilledLine = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
            
            'проход циклом по листу
            While startFilledLine < lastFilledLine
             
                  'если ячейка с XXXXXX XXXXXXXX пустая и имеется вх номер
                  If sheet.Cells(startFilledLine, 7) = "" And sheet.Cells(startFilledLine, 1) <> "" Then
                              
                        'если выбранное ФИО соответствует ФИО responsible
                        'ВЛОЖЕННАЯ ПРОВЕРКА наличия или отсутствия "отп." в cell
                        Select Case InStr(sheet.Cells(startFilledLine, 10), "отп.")
                              Case Is <> 0
                                    isSearchCurrentResult = CBool(InStr(Mid(sheet.Cells(startFilledLine, 10), 1, InStr(sheet.Cells(startFilledLine, 10), "отп.") - 1), fioIsResponsible))
                              Case Is = 0
                                    isSearchCurrentResult = CBool(InStr(sheet.Cells(startFilledLine, 10), fioIsResponsible))
                        End Select
                                    
                        If isSearchCurrentResult Then
                              isSearchResult = True
                              isSearchResultAllPersonalReport = True
                              
                              'проверка наличия общей папки для отчетов (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Недостатки по XXXX документации") Then
                                    fso.CreateFolder (homeDir & "\Недостатки по XXXX документации")
                              End If
                        
                              'ЕСЛИ НЕ передано название месяца
                              If IsMissing(sheetName) Then
                              
                                    'проверка наличия папки для отчетов за весь период (в случае отсутствия - ее создание)
                                    If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период") Then
                                          fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период")
                                    End If
                        
                                    'создает word документ на сотрудника и открывает его, если документ еще не создан
                                    If Not fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\" & fioIsResponsible & ".docx") Then

                                          'копирует образец и присваивает ему имя
                                          Call FileCopy(homeDir & "\Программные файлы\Образец для XXXXX персональный за весь период.docx", _
                                          homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\" & fioIsResponsible & ".docx")
                                          
                                          'открывает скопированный файл в режиме rw
                                          Set wordDoc = wordApp.Documents.Open(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\" & fioIsResponsible & ".docx", ReadOnly:=False)
                                          firstOpen = True
                                          
                                          'подставляет фио в заголовке
                                          wordDoc.Range.Find.Execute FindText:="&responsible", ReplaceWith:=fioIsResponsible
                              
                                    'проверяет создан ли документ
                                    'ЕСЛИ создан И закрыт
                                    ElseIf fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\" & fioIsResponsible & ".docx") _
                                    And Not fileWordDocIsOpen(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\" & fioIsResponsible & ".docx", firstOpen) Then
                                          MsgBox "Отчет за весь период на " & fioIsResponsible & " уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль поиска XXXXX документации"
                                          
                                          'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                          If newWordProccess Then
                                                wordApp.Quit
                                                Set wordApp = Nothing
                                          End If
                                          
                                          Exit Sub
                                    End If
                              Else
                              
                                    'проверка наличия папки для отчетов за выбранный месяц (в случае отсутствия - ее создание)
                                    If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам") Then
                                          fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам")
                                    End If
                                    
                                    If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName) Then
                                          fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName)
                                    End If
                              
                                    'создает word документ на сотрудника и открывает его, если документ еще не создан
                                    If Not fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" _
                                    & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx") Then
                                          
                                          'проверяет есть ли образец в программном репозитории
                                          If Not fso.FileExists(homeDir & "\Программные файлы\Образец для XXXXX персональный месячный.docx") Then
                                                MsgBox "Отсутствует образец файла в программном репозитории. Работа будет остановлена", vbCritical, "Модуль поиска XXXXX документации"
                                                Exit Sub
                                          End If

                                          'копирует образец и присваивает ему имя
                                          Call FileCopy(homeDir & "\Программные файлы\Образец для XXXXX персональный месячный.docx", _
                                          homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx")
                                          
                                          'открывает скопированный файл в режиме rw
                                          Set wordDoc = wordApp.Documents.Open(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" _
                                          & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx", ReadOnly:=False)
                                          
                                          firstOpen = True
                                          
                                          'подставляет фио в заголовке
                                          wordDoc.Range.Find.Execute FindText:="&responsible", ReplaceWith:=fioIsResponsible
                                          'подстановка месяца в заголовке
                                          wordDoc.Range.Find.Execute FindText:="&month", ReplaceWith:=sheetName
                              
                                    'проверяет создан ли документ
                                    'ЕСЛИ создан И закрыт
                                    ElseIf fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx") _
                                    And Not fileWordDocIsOpen(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\" & fioIsResponsible & ".docx", firstOpen) Then
                                          MsgBox "Отчет за " & sheetName & " на " & fioIsResponsible & " уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль поиска XXXXX документации"
                                          
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
                        
                              'добавляется строка в word документ
                              wordDoc.Tables(1).Rows.Add
                                    
                              'вызов функции копирования данных в документ word
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)), stringInWordDoc - 1, stringInWordDoc)
                                    
                              'увеличивает счетчик строк word документа
                              stringInWordDoc = stringInWordDoc + 1
                        End If
                  End If
        
                  'увеличивает счетчик строк excel документа
                  startFilledLine = startFilledLine + 1
            Wend
            
            'возвращает переменной первоначальное значение при переходе на новый лист
            startFilledLine = 9
      Next
        
      If isSearchResult Then
            'добавляет в последнюю строку word документа общее количество и закрывает его
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
                        MsgBox "Отчет на " & fioIsResponsible & " за " & sheetName & " сформирован", vbInformation, "Модуль поиска XXXXX документации"
                  Else
                        MsgBox "Отчет на " & fioIsResponsible & " за весь период сформирован", vbInformation, "Модуль поиска XXXXX документации"
                  End If
            End If
      Else
            If dispAlerts Then
                  If Not IsMissing(sheetName) Then
                        MsgBox "Пропуски у " & fioIsResponsible & " за " & sheetName & " отсутствуют", vbInformation, "Модуль поиска XXXXX документации"
                  Else
                        MsgBox "Пропуски у " & fioIsResponsible & " за весь период отсутствуют", vbInformation, "Модуль поиска XXXXX документации"
                  End If
            End If
      End If
      
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле поиска незаполненной XXXXX документации для выбранного сотрудника. Обратитесь к разработчику", vbCritical, "Модуль поиска XXXXX документации"
End Sub
'Модуль поиска незаполненной XXXXX документации общий
Sub CreateReportOfAllKeyDocumentation(Optional sheetName As Variant)

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
                        
                  'ЕСЛИ ячейка с XXXXX документами пустая и имеется vh N
                  If sheet.Cells(startFilledLine, 7) = "" And sheet.Cells(startFilledLine, 1) <> "" And Not IsDate(sheet.Cells(startFilledLine, 1)) Then
                        isSearchCurrentResult = True
                        
                        'проверка наличия общей папки для отчетов по ключам (в случае отсутствия - ее создание)
                        If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации") Then
                              fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации")
                        End If
                        
                        'ЕСЛИ НЕ передано название месяца
                        If IsMissing(sheetName) Then
                              
                              'проверка наличия папки для отчетов за весь период (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период") Then
                                    fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период")
                              End If
                              
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\0.Общий отчет.docx") Then
                                    
                                    'проверяет есть ли образец в программном репозитории
                                    If Not fso.FileExists(homeDir & "\Программные файлы\Образец для XXXXX общий за весь период.docx") Then
                                          MsgBox "Отсутствует образец файла в программном репозитории. Работа будет остановлена", vbCritical, "Модуль поиска XXXXX документации"
                                          Exit Sub
                                    End If
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для XXXXX общий за весь период.docx", _
                                    homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\0.Общий отчет.docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\0.Общий отчет.docx", ReadOnly:=False)
                                    firstOpen = True
                                    
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\0.Общий отчет.docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Недостатки по XXXXX документации\Отчеты за весь период\0.Общий отчет.docx", firstOpen) Then
                                    MsgBox "Общий отчет за весь период уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль поиска XXXXX документации"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        Else
                        
                              'проверка наличия папки для отчетов за выбранный месяц (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам") Then
                                    fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам")
                              End If
                              
                              If Not fso.FolderExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName) Then
                                    fso.CreateFolder (homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName)
                              End If
                        
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\0.Общий отчет.docx") Then
                                    
                                    'проверяет есть ли образец в программном репозитории
                                    If Not fso.FileExists(homeDir & "\Программные файлы\Образец для XXXXX общий месячный.docx") Then
                                          MsgBox "Отсутствует образец файла в программном репозитории. Работа будет остановлена", vbCritical, "Модуль поиска XXXXX документации"
                                          Exit Sub
                                    End If
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для XXXXX общий месячный.docx", _
                                    homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\0.Общий отчет.docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" _
                                    & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\0.Общий отчет.docx", ReadOnly:=False)
                                    
                                    firstOpen = True
                                    
                                    'подстановка месяца в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&month", ReplaceWith:=sheetName
                              
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName & "\0.Общий отчет.docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Недостатки по XXXXX документации\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчеты за " & sheetName _
                              & "\0.Общий отчет.docx", firstOpen) Then
                                    
                                    MsgBox "Общий отчет за " & sheetName & " уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль поиска XXXXX документации"
                                    
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
                        Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)), stringInWordDoc - 1, stringInWordDoc)
                        
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
                  MsgBox "Общий отчет за " & sheetName & " сформирован", vbInformation, "Модуль поиска XXXXX документации"
            Else
                  MsgBox "Общий отчет за весь период сформирован", vbInformation, "Модуль поиска XXXXX документации"
            End If
      Else
            If Not IsMissing(sheetName) Then
                  MsgBox "Пропуски за " & sheetName & " отсутствуют", vbInformation, "Модуль поиска XXXXX документации"
            Else
                  MsgBox "Пропуски за весь период отсутствуют", vbInformation, "Модуль поиска XXXXX документации"
            End If
      End If
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле поиска незаполненной XXXXX документации. Обратитесь к разработчику", vbCritical, "Модуль поиска XXXXX документации"
End Sub
