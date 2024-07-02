VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Модуль печати"
   ClientHeight    =   7966
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4984
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'button select a printer for book
Private Sub CommandButton6_Click()
      Application.Dialogs(xlDialogPrinterSetup).Show
      Me.TextBox5 = Application.ActivePrinter
End Sub
'button print
Private Sub CommandButton4_Click()

      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      With Me
            'ЕСЛИ НЕ выбран месяц
            If .ComboBox1.Text = "" Then
                  MsgBox "Выберите месяц"
            'ЕСЛИ ввод нумерация НЕ отключен И в ней пустое значение
            ElseIf .TextBox4.Enabled And .TextBox4.Text = "" Then
                  MsgBox "Задайте начальное значение нумерации"
            'ЕСЛИ ввод нумерации НЕ отключен И введен строковый литерал
            ElseIf .TextBox4.Enabled And Not IsNumeric(.TextBox4.Text) Then
                  MsgBox "Нумерация листов не может быть строковым литералом"
            'ЕСЛИ границы печати пустые
            ElseIf .TextBox2.Text = "" Or .TextBox3.Text = "" Then
                  MsgBox "Задайте границы печати"
            'ЕСЛИ введены буквенные значения
            ElseIf Not IsNumeric(.TextBox2.Text) Or Not IsNumeric(.TextBox3.Text) Then
                  MsgBox "Границы печати не могут быть строковым литералом"
            Else
                  'ЕСЛИ начальное значение четной страницы меньше, чем конечное (печать от начало к концу четных страниц (не совсем корректная обработка журнала))
                  If .TextBox2.Text Mod 2 = 0 And .TextBox2.Text < .TextBox3.Text Then
                        Select Case MsgBox("Вы уверены, что хотите напечатать четные страницы от начала к концу?" & Chr(10) + Chr(10) _
                        & "Удобнее напечатать журнал с корректировкой параметра ""Вывод"" у принтера.", vbYesNo + vbQuestion + vbDefaultButton2, "Модуль печати")
                        Case vbYes
                              Select Case MsgBox("Вы уверены, что настройки принтера заданы корректно?" & Chr(10) + Chr(10) _
                              & "При выборе ""Нет"" откроется окно выбора принтера, в нем необходимо нажать клавишу ""Установка"" для настройки принтера для выбранного месяца." & Chr(10), _
                              vbYesNoCancel + vbQuestion + vbDefaultButton3, "Модуль печати")
                              Case vbYes
                                    'если печатаются четные листы
                                    If .TextBox4.Text = "" Or Not .TextBox4.Enabled Then
                                          Call PrintsJournal(.ComboBox1.Text, CInt(.TextBox2.Text), CInt(.TextBox3.Text))
                                          Exit Sub
                                    Else
                                          Call PrintsJournal(.ComboBox1.Text, CInt(.TextBox2.Text), CInt(.TextBox3.Text), CInt(.TextBox4.Text))
                                          Exit Sub
                                    End If
                              Case vbNo
                                    Application.ScreenUpdating = False
                                    ThisWorkbook.Worksheets(.ComboBox1.Text).Select
                                    Application.Dialogs(xlDialogPrint).Show
                                    ThisWorkbook.Worksheets("Программный лист").Select
                                    Application.ScreenUpdating = True
                                    Exit Sub
                              Case vbCancel
                              End Select
                        Case vbNo
                              Exit Sub
                        End Select
                  End If
                  
                  Select Case MsgBox("Вы уверены, что настройки принтера заданы корректно?" & Chr(10) + Chr(10) _
                  & "При выборе ""Нет"" откроется окно выбора принтера, в нем необходимо нажать клавишу ""Установка"" для настройки принтера для выбранного месяца." & Chr(10), _
                  vbYesNoCancel + vbQuestion + vbDefaultButton3, "Модуль печати")
                  Case vbYes
                        If .TextBox4.Text = "" Or Not .TextBox4.Enabled Then
                              Call PrintsJournal(.ComboBox1.Text, CInt(.TextBox2.Text), CInt(.TextBox3.Text))
                              Exit Sub
                        Else
                              Call PrintsJournal(.ComboBox1.Text, CInt(.TextBox2.Text), CInt(.TextBox3.Text), CInt(.TextBox4.Text))
                              Exit Sub
                        End If
                  Case vbNo
                        Application.ScreenUpdating = False
                        ThisWorkbook.Worksheets(.ComboBox1.Text).Select
                        Application.Dialogs(xlDialogPrint).Show
                        ThisWorkbook.Worksheets("Программный лист").Select
                        Application.ScreenUpdating = True
                        Exit Sub
                  Case vbCancel
                  End Select
            End If
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка пользовательской формы модуля печати. Обратитесь к разработчику", vbCritical
End Sub
'button close
Private Sub CommandButton5_Click()
      Me.Hide
End Sub
'action enabled/disabled list number
Private Sub TextBox2_Change()
      With Me
            If IsNumeric(.TextBox2.Text) Then
                  On Error Resume Next
                  If CInt(.TextBox2.Text) Mod 2 = 0 Then
                        .TextBox4.Enabled = False
                  Else
                        .TextBox4.Enabled = True
                  End If
            End If
      End With
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
            Me.CommandButton4.Enabled = False
            Me.CommandButton6.Enabled = False
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
                        Me.CommandButton4.Enabled = False
                        Me.CommandButton6.Enabled = False
                        Exit Sub
                  Else
                        sheetNameArray(i) = sheet.Name
                        numberMonth = Empty
                        i = i + 1
                  End If
            End If
      Next

      With Me
            With .ComboBox1
                  .ControlTipText = "Выберите месяц для печати"
                  .Style = fmStyleDropDownList
                  .Font.Size = 10
                  .List = sheetNameArray
            End With
            
            .TextBox4.Enabled = False
            .TextBox5 = Application.ActivePrinter
            .TextBox5.AutoSize = True
            .TextBox5.Locked = True
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка инициализации пользовательской формы печати. Обратитесь к разработчику", vbCritical
End Sub
