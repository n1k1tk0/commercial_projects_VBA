VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Модуль проверки нумерации"
   ClientHeight    =   2975
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4200
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'button ok
Private Sub CommandButton1_Click()
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      Dim sheet As Worksheet
      'массив для получения всех листов из книги
      Dim sheetNameArray() As String
      ReDim sheetNameArray(ThisWorkbook.Worksheets.Count - 2)
      
      'индекс массива
      Dim i As Integer
      i = 0
      
      For Each sheet In ThisWorkbook.Worksheets
            If sheet.Name <> "Программный лист" Then
                  sheetNameArray(i) = sheet.Name
                  i = i + 1
            End If
      Next
      
      With Me
      
            'ЕСЛИ границы числа
            If .TextBox1.Text = "" Or .TextBox2.Text = "" Then
                  MsgBox "Введите границы номеров"
            ElseIf IsNumeric(.TextBox1.Text) And IsNumeric(.TextBox2.Text) Then
                  If .TextBox1.Text < .TextBox2.Text Then
                        Call findNumberException(.TextBox1.Text, .TextBox2.Text)
                  'ЕСЛИ левая граница больше правой
                  Else
                        MsgBox "Левая граница не может быть больше правой"
                  End If
            Else
                  MsgBox "Границы не могут быть строковым литералом"
            End If
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка пользовательской формы проверки нумерации. Обратитесь к разработчику", vbCritical, "Пользовательская форма проверки нумерации"
End Sub
'button close
Private Sub CommandButton2_Click()
      Me.Hide
End Sub
'userForm completion
Private Sub UserForm_Initialize()
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'массив для получения всех листов из книги
      Dim sheetNameArray() As String
      
      'индекс массива
      Dim i As Integer
      i = 0
      
      Dim sheet As Worksheet
      'переменная для проверки названий месяцев
      Dim numberMonth As Variant
      
      With Me
            .TextBox1.Font.Size = 10
            .TextBox2.Font.Size = 10
            
            If ThisWorkbook.Worksheets.Count < 2 Then
                  MsgBox "Добавьте листы для работы"
                  .CommandButton1.Enabled = False
                  .TextBox1.Enabled = False
                  .TextBox2.Enabled = False
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
                              .CommandButton1.Enabled = False
                              .TextBox1.Enabled = False
                              .TextBox2.Enabled = False
                              Exit Sub
                        Else
                              sheetNameArray(i) = sheet.Name
                              numberMonth = Empty
                              i = i + 1
                        End If
                  End If
            Next
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка инициализации пользовательской формы проверки нумерации. Обратитесь к разработчику", vbCritical, "Пользовательская форма проверки нумерации"
End Sub
