VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Модуль создания отчетов"
   ClientHeight    =   7665
   ClientLeft      =   105
   ClientTop       =   448
   ClientWidth     =   7588
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------
'МОДУЛЬ KEY
'----------------------------
'enabled/disabled view all in
Private Sub CheckBox2_Click()
      With Me
            Select Case .CheckBox2.Value
                  Case True
                        .CheckBox3.Enabled = False
                        .ComboBox1.Enabled = False
                  Case False
                        .CheckBox3.Enabled = True
                        .ComboBox1.Enabled = True
            End Select
      End With
End Sub
'enabled/disabled view personal in
Private Sub CheckBox3_Click()
      With Me
            Select Case .CheckBox3.Value
                  Case True
                        .CheckBox2.Enabled = False
                        .ComboBox1.Enabled = True
                  Case False
                        .CheckBox2.Enabled = True
                        .ComboBox1.Enabled = False
            End Select
      End With
End Sub
'enabled/disabled view month
Private Sub CheckBox8_Click()
      With Me
            If .CheckBox8 Then
                  .ComboBox3.Enabled = False
            Else
                  .ComboBox3.Enabled = True
            End If
      End With
End Sub
'enabled/disabled view all month
Private Sub CheckBox9_Click()
      With Me
            If .CheckBox9 Then
                  .ComboBox4.Enabled = False
            Else
                  .ComboBox4.Enabled = True
            End If
      End With
End Sub
'button ok
Private Sub CommandButton1_Click()
      
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      Dim responsible As Variant
      
      With Me
      
            'responsible all frame
            'ЕСЛИ НЕ выбран ни один вид отчета
            If Not .CheckBox1 And Not .CheckBox2 And Not .CheckBox3 Then
                  MsgBox "Выберите вид отчета"
            Else
            
                  'responsible frame1
                  If .CheckBox2 Or .CheckBox3 Then
                        
                        'ЕСЛИ НЕ выбран месяц И НЕ отмечен весь период
                        If .ComboBox3.Text = "" And Not .CheckBox8 Then
                              MsgBox "Выберите период для персонального отчета"
                              
                        'ЕСЛИ НЕ выбрана фамилия
                        ElseIf .CheckBox3 And .ComboBox1.Text = "" Then
                              MsgBox "Выберите ФИО сотрудника"
                              
                        'ЕСЛИ выбрана фамилия
                        ElseIf .CheckBox3 And .ComboBox1.Text <> "" Then
                        
                              'ЕСЛИ выбран месяц И выбрана фамилия
                              If .ComboBox3.Text <> "" And .ComboBox3.Enabled And .ComboBox1.Text <> "" Then
                                    Call CreateReportOfKeyDocumentation(.ComboBox1.Text, True, False, .ComboBox3.Text)
                              
                              'ЕСЛИ выбран весь период И выбрана фамилия
                              Else
                                    Call CreateReportOfKeyDocumentation(.ComboBox1.Text, True, False)
                              End If
                              
                        'ЕСЛИ выбран общий персональный отчет
                        ElseIf .CheckBox2 Then
                              
                              'ЕСЛИ НЕ выбран месяц И НЕ отмечен весь период
                              If .ComboBox3.Text = "" And Not .CheckBox8 Then
                                    MsgBox "Выберите период для персональных отчетов"
                              
                              'ЕСЛИ выбран месяц и НЕ отмечен весь период
                              ElseIf .ComboBox3.Text <> "" And Not .CheckBox8 Then
                                    
                                    'переменная для корректного оповещения о наличии пропусков
                                    Dim flag3 As Boolean
                                    flag3 = False
                                    
                                    For Each responsible In ThisWorkbook.Worksheets("Программный лист").Range("A10:A" & Cells(Rows.Count, Range("A1").Column).End(xlUp).Row)
                                          'вызов обработки общего списка для выбранного месяца
                                          Call CreateReportOfKeyDocumentation(CStr(responsible), False, flag3, .ComboBox3.Text)
                                    Next
                              Else
                                    
                                    For Each responsible In ThisWorkbook.Worksheets("Программный лист").Range("A10:A" & Cells(Rows.Count, Range("A1").Column).End(xlUp).Row)
                                          'вызов обработки общего списка за весь период
                                          Call CreateReportOfKeyDocumentation(CStr(responsible), False, flag3)
                                    Next
                              End If
                              
                              If flag3 Then
                                    If .ComboBox3.Text = "" Then
                                          MsgBox "Персональные отчеты за весь период сформированы"
                                    Else
                                          MsgBox "Персональные отчеты за " & .ComboBox3.Text & " сформированы"
                                    End If
                              Else
                                    If .ComboBox3.Text = "" Then
                                          MsgBox "Пропуски за весь период отсутствуют"
                                    Else
                                          MsgBox "Пропуски за " & .ComboBox3.Text & " отсутствуют"
                                    End If
                              End If
                        End If
                  End If
        
                  'responsible frame2
                  If .CheckBox1 Then
                        
                        'ЕСЛИ месяц НЕ выбран И НЕ выбран весь период
                        If .ComboBox4 = "" And Not .CheckBox9 Then
                              MsgBox "Выберите период для общего отчета"
                        
                        'ЕСЛИ месяц выбран И НЕ выбран весь период
                        ElseIf .ComboBox4 <> "" And Not .CheckBox9 Then
                              
                              'вызов обработки общего отчета на месяц
                              Call CreateReportOfAllKeyDocumentation(.ComboBox4.Text)
                        
                        'ЕСЛИ месяц НЕ выбран И выбран весь период
                        Else
                              
                              'вызов обработки общего отчета на весь год
                              Call CreateReportOfAllKeyDocumentation
                        End If
                  End If
            End If
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка пользовательской формы. Обратитесь к разработчику"
End Sub
'button close
Private Sub CommandButton2_Click()
    UserForm1.Hide
End Sub
'---------------------------------------
'МОДУЛЬ ПОДВИСШИХ XXXX
'---------------------------------------
'enabled/disabled view personal in
Private Sub CheckBox4_Click()
      With Me
            Select Case .CheckBox4.Value
                  Case True
                        .CheckBox5.Enabled = False
                        .ComboBox2.Enabled = False
                  Case False
                        .CheckBox5.Enabled = True
                        .ComboBox2.Enabled = True
            End Select
      End With
End Sub
'enabled/disabled view all in
Private Sub CheckBox5_Click()
      With Me
            Select Case .CheckBox5.Value
                  Case True
                        .CheckBox4.Enabled = False
                        .ComboBox2.Enabled = True
                  Case False
                        .CheckBox4.Enabled = True
                        .ComboBox2.Enabled = False
            End Select
      End With
End Sub
'enabled/disabled view month
Private Sub CheckBox10_Click()
      With Me
            If .CheckBox10 Then
                  .ComboBox5.Enabled = False
            Else
                  .ComboBox5.Enabled = True
            End If
      End With
End Sub
'enabled/disabled view month all
Private Sub CheckBox11_Click()
      With Me
            If .CheckBox11 Then
                  .ComboBox6.Enabled = False
            Else
                  .ComboBox6.Enabled = True
            End If
      End With
End Sub
'button ok
Private Sub CommandButton8_Click()
      'обработчик ошибок
      On Error GoTo ErrorHandler
      
      Dim responsible As Variant
      
      With Me
      
            'responsible all frame
            'ЕСЛИ НЕ выбран ни один вид отчета
            If Not .CheckBox4 And Not .CheckBox5 And Not .CheckBox6 Then
                  MsgBox "Выберите вид отчета"
            Else
            
                  'responsible frame1
                  If .CheckBox4 Or .CheckBox5 Then
                  
                        'ЕСЛИ НЕ выбран месяц И НЕ отмечен весь период
                        If .ComboBox5.Text = "" And Not .CheckBox10 Then
                              MsgBox "Выберите период для персонального отчета"
                        
                        'ЕСЛИ НЕ выбрана фамилия
                        ElseIf .ComboBox2.Text = "" And .CheckBox5 Then
                              MsgBox "Выберите ФИО сотрудника"
                                          
                        'ЕСЛИ выбрана фамилия
                        ElseIf .ComboBox2.Text <> "" And .CheckBox5 Then
                              
                              'ЕСЛИ выбран месяц И выбрана фамилия
                              If .ComboBox2.Text <> "" And .ComboBox2.Enabled And .ComboBox5.Text <> "" Then
                                    Call CreateReportPersonalUnclosedDocumentation(.ComboBox2.Text, True, False, .ComboBox5.Text)
                              
                              'ЕСЛИ выбран весь период И выбрана фамилия
                              Else
                                    Call CreateReportPersonalUnclosedDocumentation(.ComboBox2.Text, True, False)
                              End If
                             
                        'ЕСЛИ выбран общий персональный отчет
                        ElseIf .CheckBox4 Then
                        
                              'ЕСЛИ НЕ выбран месяц И НЕ отмечен весь период
                              If .ComboBox5.Text = "" And Not .CheckBox10 Then
                                    MsgBox "Выберите период для персональных отчетов"
                                    
                              'ЕСЛИ выбран месяц и НЕ отмечен весь период
                              ElseIf .ComboBox5.Text <> "" And Not .CheckBox10 Then
                              
                                    'переменная для корректного оповещения о наличии пропусков
                                    Dim flag3 As Boolean
                                    flag3 = False
                              
                                    For Each responsible In ThisWorkbook.Sheets("Программный лист").Range("A10:A" & Cells(Rows.Count, Range("A1").Column).End(xlUp).Row)
                                          'вызов обработки общего списка для выбранного месяца
                                          Call CreateReportPersonalUnclosedDocumentation(CStr(responsible), False, flag3, .ComboBox5.Text)
                                    Next
                              Else
                                    For Each responsible In ThisWorkbook.Sheets("Программный лист").Range("A10:A" & Cells(Rows.Count, Range("A1").Column).End(xlUp).Row)
                                          'вызов обработки общего списка для выбранного месяца
                                          Call CreateReportPersonalUnclosedDocumentation(CStr(responsible), False, flag3)
                                    Next
                              End If
                              
                              If flag3 Then
                                    MsgBox "Отчеты сформированы"
                              Else
                                    MsgBox "Пропуски отсутствуют"
                              End If
                        End If
                  End If
        
                  'responsible frame2
                  If .CheckBox6 Then
                  
                        'ЕСЛИ месяц НЕ выбран И НЕ выбран весь период
                        If .ComboBox6.Text = "" And Not .CheckBox11 Then
                              MsgBox "Выберите период для общего отчета"
                        
                        'ЕСЛИ месяц выбран И НЕ выбран весь период
                        ElseIf .ComboBox6.Text <> "" And Not .CheckBox11 Then
                        
                              'вызов обработки общего отчета на месяц
                              Call CreateReportAllUnclosedDocumentation(.ComboBox6.Text)
                        
                        'ЕСЛИ месяц НЕ выбран И выбран весь период
                        Else
                              
                              'вызов обработки общего отчета на весь год
                              Call CreateReportAllUnclosedDocumentation
                        End If
                  End If
            End If
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка пользовательской формы. Обратитесь к разработчику"
End Sub
'-------------------------------------------------
'МОДУЛЬ НЕВОЗВРАЩЕННЫХ XXXX
'-------------------------------------------------
'enabled/disabled view month all
Private Sub CheckBox12_Click()
      With Me
            If .CheckBox12 Then
                  .ComboBox7.Enabled = False
            Else
                  .ComboBox7.Enabled = True
            End If
      End With
End Sub
'button ok
Private Sub CommandButton6_Click()
      With Me
            'ЕСЛИ не выбран вид отчета
            If Not .CheckBox7 Then
                  MsgBox "Выберите вид отчета"
            'ЕСЛИ НЕ выбран месяц И НЕ отмечен весь период
            ElseIf .ComboBox7.Text = "" And Not .ComboBox7.Enabled And Not .CheckBox12 Then
                  MsgBox "Выберите период для отчета"
            Else
                  'ЕСЛИ отмечен весь период
                  If .CheckBox12 Then
                        Call CreateReportAllUnreturnedDocumentation
                  'ЕСЛИ выбран месяц И НЕ отмечен весь период
                  ElseIf Not .CheckBox12 And .ComboBox7.Text <> "" Then
                        Call CreateReportAllUnreturnedDocumentation(.ComboBox7.Text)
                  End If
            End If
      End With
End Sub
'button cancel
Private Sub CommandButton7_Click()
      Me.Hide
End Sub
'userForm completion
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
      
      If ThisWorkbook.Worksheets.Count < 2 Then
            MsgBox "Добавьте листы для работы"
            Me.CommandButton1.Enabled = False
            Me.CommandButton6.Enabled = False
            Me.CommandButton8.Enabled = False
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
                        Me.CommandButton6.Enabled = False
                        Me.CommandButton8.Enabled = False
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
                  .ControlTipText = "Выберите ФИО из списка"
                  .Enabled = False
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .RowSource = ThisWorkbook.Sheets("Программный лист").Range("A10:A" & Cells(Rows.Count, Range("A1").Column).End(xlUp).Row).Address
            End With
    
            With .ComboBox2
                  .ControlTipText = "Выберите ФИО из списка"
                  .Enabled = False
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .RowSource = ThisWorkbook.Sheets("Программный лист").Range("A10:A" & Cells(Rows.Count, Range("A1").Column).End(xlUp).Row).Address
            End With
    
            With .ComboBox3
                  .ControlTipText = "Выберите месяц"
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .List = sheetNameArray
            End With
            
            With .ComboBox4
                  .ControlTipText = "Выберите месяц"
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .List = sheetNameArray
            End With
            
            With .ComboBox5
                  .ControlTipText = "Выберите месяц"
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .List = sheetNameArray
            End With
            
            With .ComboBox6
                  .ControlTipText = "Выберите месяц"
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .List = sheetNameArray
            End With
            
            With .ComboBox7
                  .ControlTipText = "Выберите месяц"
                  .Style = fmStyleDropDownList
                  .Font.Size = 11
                  .Font = "Times New Roman"
                  .List = sheetNameArray
            End With
    
            .CheckBox1 = False
            .CheckBox2 = False
            .CheckBox3 = False
            .CheckBox4 = False
            .CheckBox5 = False
            .CheckBox6 = False
            .CheckBox7 = False
            .CheckBox8 = False
            .CheckBox10 = False
            .CheckBox11 = False
            .CheckBox12 = False
      End With
Exit Sub

ErrorHandler:
      MsgBox "Ошибка пользовательской формы. Обратитесь к разработчику"
End Sub
'button close
Private Sub CommandButton4_Click()
      UserForm1.Hide
End Sub
