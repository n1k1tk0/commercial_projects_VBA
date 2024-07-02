VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Модуль выгрузки XX"
   ClientHeight    =   3696
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4844
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'enable/disable view element
Private Sub CheckBox1_Click()
      With Me
            If .CheckBox1 Then
                  .CheckBox2.Enabled = False
            Else
                  .CheckBox2.Enabled = True
            End If
      End With
End Sub
'enable/disable view element
Private Sub CheckBox2_Click()
      With Me
            If .CheckBox2 Then
                  .CheckBox1.Enabled = False
                  .ComboBox1.Enabled = True
            Else
                  .CheckBox1.Enabled = True
                  .ComboBox1.Enabled = False
            End If
      End With
End Sub
'enable/disable view element
Private Sub ComboBox1_Change()
      With Me
            If .ComboBox1.Text <> "" Then
                  .CheckBox1.Enabled = False
            Else
                  .CheckBox1.Enabled = True
            End If
      End With
End Sub
'button ok
Private Sub CommandButton1_Click()
      With Me
            If Not .CheckBox1 And Not .CheckBox2 Then
                  MsgBox "Выберите период отчета", vbExclamation, "Пользовательская форма выгрузки черного материала"
            Else
                  If .CheckBox1 Then
                        Call SearchDraftMaterial
                  ElseIf .CheckBox2 Then
                        If .ComboBox1.Text = "" Then
                              MsgBox "Выберите месяц", vbExclamation, "Пользовательская форма выгрузки черного материала"
                        Else
                              Call SearchDraftMaterial(.ComboBox1.Text)
                        End If
                  End If
            End If
      End With
End Sub
'button cancel
Private Sub CommandButton2_Click()
      Me.Hide
End Sub
'userform completion
Private Sub UserForm_Initialize()
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      'массив для получения всех месяцев из книги
      Dim sheetNameArray() As String
      
      'индекс
      Dim i As Integer
      i = 0
      
      Dim sheet As Worksheet
      'переменная для проверки названий месяцев
      Dim numberMonth As Variant
      
      With Me
            If ThisWorkbook.Worksheets.Count < 2 Then
                  MsgBox "Добавьте листы для работы"
                  .ComboBox1.Enabled = False
                  .CheckBox1.Enabled = False
                  .CommandButton1.Enabled = False
                  Exit Sub
            Else
                  ReDim sheetNameArray(ThisWorkbook.Worksheets.Count - 2)
            End If
      
            For Each sheet In ThisWorkbook.Worksheets
                  If sheet.Name <> "Программный лист" Then
                        On Error Resume Next
                        numberMonth = month(DateValue("08/" & sheet.Name & "/1998"))
                        If IsEmpty(numberMonth) Then
                              MsgBox "Переменуйте лист "" & sheet.Name & "" в словесно-действительное значение месяца, иначе работа будет невозможна", vbCritical, "Модуль выгрузки черного материала"
                              .ComboBox1.Enabled = False
                              .CheckBox1.Enabled = False
                              .CommandButton1.Enabled = False
                              Exit Sub
                        Else
                              sheetNameArray(i) = sheet.Name
                              numberMonth = Empty
                              i = i + 1
                        End If
                  End If
            Next
            
            With .ComboBox1
                  .Enabled = False
                  .ControlTipText = "Выберите ФИО из списка"
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .List = sheetNameArray
            End With
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка инициализации пользовательской формы. Обратитесь к разработчику", vbCritical, "Пользовательская форма выгрузки черного материала"
End Sub
