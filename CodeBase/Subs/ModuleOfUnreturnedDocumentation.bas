Attribute VB_Name = "ModuleOfUnreturnedDocumentation"
Option Explicit
'модуль поиска невозвращенных XXXXX
Sub CreateReportAllUnreturnedDocumentation(Optional sheetName As Variant)
      
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
      'номер последней заполненной строки на листе
      Dim lastFilledLine As Integer
      'номер первой ячейки заполненной строки на листе
      Dim startFilledLine As Integer
      startFilledLine = 9
      'номер первой строки для заполнения в word документе
      Dim stringInWordDoc As Integer
      stringInWordDoc = 2
      'массива рабочих листов
      Dim sheetArray() As Worksheet
      'array index
      Dim i As Integer
      i = 0
      Dim sheet As Variant
      'переменная, отражающая результат текущего поиска
      Dim isSearchCurrentResult As Boolean
      isSearchCurrentResult = False
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
                  'ЕСЛИ ячейка с ФИО делопроизводителя пустая
                  'И ячейка с конечным состоянием XXXXX содержит ФИО руководителя или орган XXXXX
                  'И ячейка с конечным состоянием XXXXX не merge
                  'И ячейка с конечным состоянием не содержит "XXX"
                  If sheet.Cells(startFilledLine, 16) = "" _
                  And sheet.Cells(startFilledLine, 17) <> "" _
                  And Not sheet.Cells(startFilledLine, 17).MergeCells _
                  And Not CBool(InStr(sheet.Cells(startFilledLine, 17), "XXX")) Then
                        isSearchCurrentResult = True
                        
                        'создает папку по недостаткам, если она еще не создана
                        If Not fso.FolderExists(homeDir & "\Невозвращенные XXXXX") Then
                              fso.CreateFolder (homeDir & "\Невозвращенные XXXXX")
                        End If
                        
                        'ЕСЛИ НЕ передано название месяца
                        If IsMissing(sheetName) Then
                        
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Невозвращенные XXXXX\Отчет за весь период.docx") Then
                                    
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для невозвращенных XXXXX за весь период.docx", homeDir & "\Невозвращенные XXXXX\Отчет за весь период.docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Невозвращенные XXXXX\Отчет за весь период.docx", ReadOnly:=False)
                                    firstOpen = True
                                    
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Невозвращенные XXXXX\Отчет за весь период.docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Невозвращенные XXXXX\Отчет за весь период.docx", firstOpen) Then
                                    MsgBox "Отчет за весь период уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль невозвращенных XXXXX"
                                    
                                    'ЕСЛИ процесс ворда инициализирован алгоритмом - высвобождает его
                                    If newWordProccess Then
                                          wordApp.Quit
                                          Set wordApp = Nothing
                                    End If
                                    
                                    Exit Sub
                              End If
                        Else
                              
                              'проверка наличия папки для отчетов за выбранный месяц (в случае отсутствия - ее создание)
                              If Not fso.FolderExists(homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам") Then
                                    fso.CreateFolder (homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам")
                              End If
                                                            
                              'создает word документ и открывает его, если документ еще не создан
                              If Not fso.FileExists(homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx") Then
                           
                                    'копирует образец и присваивает ему имя
                                    Call FileCopy(homeDir & "\Программные файлы\Образец для невозвращенных XXXXX месячный.docx", _
                                    homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx")
                                    
                                    'открывает скопированный файл в режиме rw
                                    Set wordDoc = wordApp.Documents.Open(homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам\" _
                                    & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx", ReadOnly:=False)
                                    firstOpen = True
                                    
                                    'подставляет месяц в заголовке
                                    wordDoc.Range.Find.Execute FindText:="&month", ReplaceWith:=sheetName
                              
                              'проверяет создан ли документ
                              'ЕСЛИ создан И закрыт
                              ElseIf fso.FileExists(homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx") _
                              And Not fileWordDocIsOpen(homeDir & "\Невозвращенные XXXXX\Отчеты по месяцам\" & month(DateValue("08/" & sheetName & "/1998")) & ".Отчет за " & sheetName & ".docx", firstOpen) Then
                                    MsgBox "Отчет за " & sheetName & " уже существует. Удалите его и запустите программу снова", vbInformation, "Модуль невозвращенных XXXXX"
                                    
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
                              & sheet.Cells(startFilledLine, 15) & " / " & vbNewLine & Replace(sheet.Cells(startFilledLine, 17), Chr(10), " "), stringInWordDoc - 1, stringInWordDoc)
                        Else
                              Call copyDataToWord(wordDoc, getNumber(sheet.Cells(startFilledLine, 1)) & " / " & sheet.Cells(startFilledLine, 15) & " / " _
                              & sheet.Cells(startFilledLine, 17), stringInWordDoc - 1, stringInWordDoc)
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
                  MsgBox "Отчет за " & sheetName & " сформирован", vbInformation, "Модуль невозвращенных XXXXX"
            Else
                  MsgBox "Общий отчет за весь период сформирован", vbInformation, "Модуль невозвращенных XXXXX"
            End If
      Else
            If Not IsMissing(sheetName) Then
                  MsgBox "Невозвращенные телеграммы за " & sheetName & " отсутствуют", vbInformation, "Модуль невозвращенных XXXXX"
            Else
                  MsgBox "Невозвращенные телеграммы за весь период отсутствуют", vbInformation, "Модуль невозвращенных XXXXX"
            End If
      End If
Exit Sub

'обработчик ошибок
ErrorHandler:
      MsgBox "Произошла ошибка в модуле поиска невозвращенных XXXXX. Обратитесь к разработчику", vbCritical, "Модуль невозвращенных XXXXX"
End Sub
