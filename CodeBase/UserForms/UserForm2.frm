VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Модуль подготовки к печати"
   ClientHeight    =   8988.001
   ClientLeft      =   105
   ClientTop       =   448
   ClientWidth     =   7007
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'checkbox vh
Private Sub CheckBox4_Click()
      With Me
            'on/off ui для дальнейшей работы
            'кнопки
            .CommandButton1.Enabled = .CheckBox4
            .CommandButton2.Enabled = .CheckBox4
            .CommandButton3.Enabled = .CheckBox4
                  
            'combobox'ы
            .ComboBox2.Enabled = .CheckBox4
            .ComboBox3.Enabled = .CheckBox4
                  
            'checkbox'ы
            .CheckBox2.Enabled = .CheckBox4
            .CheckBox3.Enabled = .CheckBox4
                  
            'textbox
            .TextBox1.Enabled = .CheckBox4
            'блокировка другого варианта
            .CheckBox5.Enabled = Not .CheckBox4
      End With
End Sub
'checkbox vh XXXX
Private Sub CheckBox5_Click()
      With Me
            'on/off ui для дальнейшей работы
            'кнопки
            .CommandButton1.Enabled = .CheckBox5
            .CommandButton2.Enabled = .CheckBox5
            .CommandButton3.Enabled = .CheckBox5
                  
            'combobox'ы
            .ComboBox2.Enabled = .CheckBox5
            .ComboBox3.Enabled = .CheckBox5
                  
            'checkbox'ы
            .CheckBox2.Enabled = .CheckBox5
            .CheckBox3.Enabled = .CheckBox5
                  
            'textbox
            .TextBox1.Enabled = .CheckBox5
            'блокировка другого варианта
            .CheckBox4.Enabled = Not .CheckBox5
      End With
End Sub
'button alignment
Private Sub CommandButton1_Click()
      With Me
            If .ComboBox3.Text <> "" Then
                  Call MergeCellsAutoFit(.ComboBox3.Text)
                  MsgBox "Выравнивание успешно завершенно"
            Else
                  MsgBox "Выберите месяц для выравнивания"
            End If
      End With
End Sub
'button adjust the print boundaries
Private Sub CommandButton2_Click()
      With Me
            If .ComboBox3.Text <> "" Then
                  Call AdjustPrintBoundaries(.ComboBox3.Text)
                  MsgBox "Границы печати заданы успешно"
            Else
                  MsgBox "Выберите месяц для задания границ печати"
            End If
      End With
End Sub
'button adjust accounting number
Private Sub CommandButton3_Click()
      'отключает UI
      Application.ScreenUpdating = False
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'последняя заполненная строка на листе
      Dim lastFilledLine As Integer
      Dim sheet As Worksheet
      Dim textBoxShape As Excel.Shape
      
      If Me.ComboBox2.Text <> "" Then
            lastFilledLine = ThisWorkbook.Worksheets(Me.ComboBox2.Text).Cells(Rows.Count, Range("A10").Column).End(xlUp).Row
            Set sheet = ThisWorkbook.Worksheets(Me.ComboBox2.Text)
      End If
      'массив для разделения колонтитула
      Dim str() As String
      
            
      If Me.ComboBox2.Text <> "" Then
            If Me.TextBox1.Text <> "" Then
                  
                  'ЕСЛИ учетный номер является числом
                  If IsNumeric(Me.TextBox1.Text) Then
                  
                        'проверяет на пустоту вводимых строк
                        If Me.CheckBox3 And (Me.TextBox6.Text = "" Or Me.TextBox7.Text = "") Then
                              MsgBox "Задайте необходимые учетные данные или отмените установку параметра"
                              Exit Sub
                        End If
                        
                        'устанавливает параметры печати колонтитулов для четных и нечетных страниц
                        With sheet.PageSetup
                              
                              'устанавливает разные колонтитулы для четных и нечетных страниц
                              .OddAndEvenPagesHeaderFooter = True
                              'устанавливает левый нижний колонтитул для НЕчетных страниц
                              .LeftFooter = "&""Times New Roman""&12" & "Уч. № " & CInt(Me.TextBox1.Text)
                              'устанавливает левый нижний колонтитул для четных страниц
                              .EvenPage.LeftFooter.Text = "&""Times New Roman""&12" & "Уч. № " & Me.TextBox1.Text & Chr(10) & Trim(sheet.Cells(1, 1))
                                    
                              'корректирует учетные данные
                              'ЕСЛИ данные нужно скорректировать
                              If Me.CheckBox2 Then
                                          
                                    'Устанавливает правый нижний колонтитул нечетной страницы с учетом отредактированного
                                    .RightFooter = "&""Times New Roman""&10" & " " & Me.TextBox5.Text
                                          
                                    'изменяет фамилию в последней строке в соответствии с колонтитулом
                                    If CBool(InStr(Me.TextBox5.Text, "г. ")) Then
                                                
                                          'разделяет отредактированную строку на ДО и ПОСЛЕ разделителя
                                          str = Split(Me.TextBox5.Text, "г. ", 2)
                                                      
                                          'задает изменения в последнюю строку
                                          With sheet.Cells(lastFilledLine, 1)
                                                .Value = str(1)
                                                .Font.Name = "Times New Roman"
                                                .Font.Size = 10
                                          End With
                                    Else
                                          MsgBox "Последняя строка с ФИО не была изменена по причине отсутствия формализованного разделителя", vbCritical
                                    End If
                              Else
                                    
                                    'Форматирует правый нижний колонтитул нечетной страницы
                                    .RightFooter = "&""Times New Roman""&10" & " " & Me.TextBox5.Text
                                          
                                    'задает форматирование в последнюю строку
                                    With sheet.Cells(lastFilledLine, 1)
                                          .Font.Name = "Times New Roman"
                                          .Font.Size = 10
                                    End With
                              End If
                              
                              'ЕСЛИ нужно установить
                              If Me.CheckBox3 Then
                                    
                                    'очищает все прежние
                                    Dim i As Integer
                                    If sheet.Shapes.Count > 0 Then
                                          For i = 1 To sheet.Shapes.Count
                                                sheet.Shapes.Item(i).Delete
                                          Next
                                    End If
                                    
                                    'добавляет на лист учетные данные и тут же задает текст и параметры
                                    With sheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                                                                       Application.CentimetersToPoints(0.5), _
                                                                                       Application.CentimetersToPoints(1.6), _
                                                                                       Application.CentimetersToPoints(6.69), _
                                                                                       Application.CentimetersToPoints(3.31)).TextFrame
                                                                                       
                                          With .Characters
                                                'если данные от летника корректируются
                                                If Me.CheckBox2 Then
                                                      .Text = "XXXX " & Chr(10) & _
                                                                  "XXXX" & Chr(10) & _
                                                                  "XXXX " & str(1) & Chr(10) & _
                                                                  "XXXX " & Me.TextBox6.Text & Chr(10) & _
                                                                  "XXXX " & Me.TextBox7.Text & Chr(10) & _
                                                                  DateValue(str(0))
                                                Else
                                                      str = Split(Me.TextBox5.Text, "г. ", 2)
                                                      .Text = "XXXXX " & Chr(10) & _
                                                                  "XXXXX" & Chr(10) & _
                                                                  "XXXXX " & str(1) & Chr(10) & _
                                                                  "XXXXX " & Me.TextBox6.Text & Chr(10) & _
                                                                  "XXXXX " & Me.TextBox7.Text & Chr(10) & _
                                                                  DateValue(str(0))
                                                End If
                                                            
                                                .Font.Size = 12
                                                .Font.Color = vbBlack
                                                .Font.Name = "Times New Roman"
                                          End With
                                    End With
                                    
                                    'убирает границы у фигуры
                                    sheet.Shapes.Item(1).Line.Visible = msoFalse
                                    'задает имя фугуре для дальнейшей идентификации
                                    sheet.Shapes.Item(1).Name = "programm figure"
                                    
                                    'Нижеидущий кодовый блок предназначен для корректной установки вновь созданного XXXX
                                    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                    
                                    'первая пустая строка после ФИО
                                    Dim aC As Range
                                    Set aC = sheet.Cells(lastFilledLine + 1, 1)
                                    'количество листов до прохода инструкцией
                                    Dim countPage As Integer
                                    countPage = sheet.PageSetup.Pages.Count
                                    'характеристика последнего листа (false - нечетные; true - четные)
                                    Dim oddOrEven As Boolean
                                    oddOrEven = (countPage Mod 2 = 0)
                                    
                                    'ЕСЛИ последняя страница НЕчетная
                                    If Not oddOrEven Then
                                          'ПОКА количество страниц меньше необходимого
                                          While sheet.PageSetup.Pages.Count < countPage + 2
                                                'добавляет текст в ячейку для того, чтобы счетчик считал страницы
                                                With aC
                                                      .Value = "Hello, world!"
                                                      .Font.Color = vbWhite
                                                End With
                                                'прыгает на одну ячейку вниз (имитация стрелки вниз)
                                                Set aC = aC.Offset(1)
                                                sheet.PageSetup.PrintArea = aC
                                          Wend
                                          
                                          'прыгает на нужный лист (метод .Select не работает без этого)
                                          sheet.Select
                                          'выделяет удаляемую ячейку
                                          aC.Offset(-1).Select
                                          'удаляет выделенную ячейку
                                          Selection.EntireRow.Delete
                                          
                                          'новая последняя строка
                                          Set aC = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp)
                                          
                                          'перемещает фонарик на нужное место
                                          With sheet
                                                .Shapes(1).Left = .Range("A" & aC.Offset(-7).Row).Left
                                                .Shapes(1).Top = .Range("A" & aC.Offset(-7).Row).Top
                                                
                                                'удаляет все программные записи
                                                .Range("A" & lastFilledLine + 1 & ":A" & .Cells(Rows.Count, Range("A10").Column).End(xlUp).Row).Delete
                                                .Range("A" & lastFilledLine + 1).Delete
                                                'вновь выравнивает фигуру
                                                .Shapes(1).Width = Application.CentimetersToPoints(6.69)
                                          End With
                                          
                                          'прыгает обратно
                                          ThisWorkbook.Worksheets("Программный лист").Select
                                          
                                    'ЕСЛИ последняя страница четная
                                    Else
                                          'ПОКА количество страниц меньше необходимого
                                          While sheet.PageSetup.Pages.Count < countPage + 3
                                                'добавляет текст в ячейку для того, чтобы счетчик считал страницы
                                                With aC
                                                      .Value = "Hello, world!"
                                                      .Font.Color = vbWhite
                                                End With
                                                'прыгает на одну ячейку вниз (имитация стрелки вниз)
                                                Set aC = aC.Offset(1)
                                                sheet.PageSetup.PrintArea = aC
                                          Wend
                                          
                                          'прыгает на нужный лист (метод .Select не работает без этого)
                                          sheet.Select
                                          'выделяет удаляемую ячейку
                                          aC.Offset(-1).Select
                                          'удаляет выделенную ячейку
                                          Selection.EntireRow.Delete
                                          
                                          Set aC = sheet.Cells(Rows.Count, Range("A10").Column).End(xlUp)
                                          'перемещает XXXX на нужное место
                                          With sheet
                                                .Shapes(1).Left = .Range("A" & aC.Offset(-7).Row).Left
                                                .Shapes(1).Top = .Range("A" & aC.Offset(-7).Row).Top
                                                
                                                'удаляет все программные записи
                                                .Range("A" & lastFilledLine + 1 & ":A" & .Cells(Rows.Count, Range("A10").Column).End(xlUp).Row).Delete
                                                .Range("A" & lastFilledLine + 1).Delete
                                                'вновь выравнивает фигуру
                                                .Shapes(1).Width = Application.CentimetersToPoints(6.69)
                                          End With
                                          
                                          'прыгает обратно
                                          ThisWorkbook.Worksheets("Программный лист").Select
                                    End If
                                    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                              End If
                                                      
                              'ставит галку при успешном завершении установки параметров
                              Me.CheckBox1 = True
                        End With
                  Else
                        MsgBox "Учетный номер не может быть строковым литералом"
                        Me.TextBox1.Text = ""
                  End If
            Else
                  MsgBox "Введите учетный номер"
            End If
      Else
            MsgBox "Выберите месяц"
      End If
      Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
      MsgBox "Ошибка пользовательской формы подготовки к печати. Обратитесь к разработчику", vbCritical
End Sub
'checkbox click true
Private Sub CheckBox2_Click()
      With Me
            If .CheckBox2 Then
                  .TextBox5.Enabled = True
            Else
                  .TextBox5.Enabled = False
            End If
      End With
End Sub
'checkbox click true
Private Sub CheckBox3_Click()
      With Me
            If .CheckBox3 Then
                  .TextBox6.Enabled = True
                  .TextBox7.Enabled = True
            Else
                  .TextBox6.Enabled = False
                  .TextBox7.Enabled = False
            End If
      End With
End Sub
'отображает данные для редактирования на пользовательской форме в зависимости от месяца
Private Sub ComboBox2_Change()
      Dim str() As String
      Dim sheet As Worksheet
      
      On Error Resume Next
      Set sheet = ThisWorkbook.Worksheets(Me.ComboBox2.Text)
      
      With Me
            If sheet Is Nothing Then
                  .TextBox5.Text = ""
            Else
                  'ЕСЛИ правый нижний колонтитул листа содержит "&10"
                  If CBool(InStr(ThisWorkbook.Worksheets(.ComboBox2.Text).PageSetup.RightFooter, "&10")) Then
                        'Разделяет строку на до проверяемого символа и после и добавляет обе строки в массив
                        str = Split(ThisWorkbook.Worksheets(.ComboBox2.Text).PageSetup.RightFooter, "&10", 2)
                        'отображает в TextBox вторую часть строки
                        .TextBox5.Text = LTrim(str(1))
                  '-/- для символа "&6" (дефолтная выгрузка журнала)
                  ElseIf CBool(InStr(ThisWorkbook.Worksheets(.ComboBox2.Text).PageSetup.RightFooter, "&6")) Then
                         str = Split(ThisWorkbook.Worksheets(.ComboBox2.Text).PageSetup.RightFooter, "&6", 2)
                        .TextBox5.Text = LTrim(str(1))
                  End If
            End If
      End With
End Sub
'button cancel
Private Sub CommandButton4_Click()
      Me.Hide
End Sub
'userForm completion
Private Sub UserForm_Initialize()

      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      Dim sheetNameArray() As String
      'индекс
      Dim i As Integer
      i = 0
      Dim sheet As Worksheet
      
      'переменная для проверки названий месяцев
      Dim numberMonth As Variant
      
      If ThisWorkbook.Worksheets.Count < 2 Then
            MsgBox "Добавьте листы для работы"
            Me.CommandButton1.Enabled = False
            Me.CommandButton2.Enabled = False
            Me.CommandButton3.Enabled = False
            Exit Sub
      Else
            ReDim sheetNameArray(ThisWorkbook.Worksheets.Count - 2)
      End If
      
      For Each sheet In ThisWorkbook.Worksheets
            If sheet.Name <> "Программный лист" Then
                  On Error Resume Next
                  numberMonth = month(DateValue("08/" & sheet.Name & "/1998"))
                  If IsEmpty(numberMonth) Then
                        MsgBox "Переменуйте лист "" & sheet.Name & "" в словесно-действительное значение месяца, иначе работа будет невозможна"
                        Me.CommandButton1.Enabled = False
                        Me.CommandButton2.Enabled = False
                        Me.CommandButton3.Enabled = False
                        Exit Sub
                  Else
                        sheetNameArray(i) = sheet.Name
                        numberMonth = Empty
                        i = i + 1
                  End If
            End If
      Next

      With Me
            'checkbox выбора журнала
            .CheckBox4 = False
            .CheckBox5 = False
      
            'отключение всех ui-элементов до выбора журнала
            .ComboBox2.Enabled = False
            .ComboBox3.Enabled = False
            .CheckBox2.Enabled = False
            .CheckBox3.Enabled = False
            .CheckBox1.Enabled = False
            .CheckBox2 = False
            .CheckBox3 = False
            .TextBox1.Enabled = False
            .TextBox1.Font.Size = 10
            .TextBox5.Enabled = False
            .TextBox6.Enabled = False
            .TextBox7.Enabled = False
            .CommandButton1.Enabled = False
            .CommandButton2.Enabled = False
            .CommandButton3.Enabled = False
      
            With .ComboBox2
                  .ControlTipText = "Выберите месяц для установки учетного номера"
                  .Style = fmStyleDropDownList
                  .Font.Size = 10
                  .List = sheetNameArray
            End With
            
            With .ComboBox3
                  .ControlTipText = "Выберите месяц для выравнивания"
                  .Style = fmStyleDropDownList
                  .Font.Size = 10
                  .List = sheetNameArray
            End With
            'ставит курсор на логически первую операцию
            .ComboBox3.SetFocus
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка инициализации пользовательской формы подготовки к печати. Обратитесь к разработчику", vbCritical
End Sub
