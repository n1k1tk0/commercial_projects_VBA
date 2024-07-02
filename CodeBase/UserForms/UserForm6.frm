VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "Модуль специального подсчета"
   ClientHeight    =   7819
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4921
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'button ok
Private Sub CommandButton1_Click()
      With Me
            'если входные данные не пустые
            If .TextBox1.Text <> "" And .TextBox2.Text <> "" Then
                  'если входные данные числа
                  If IsNumeric(.TextBox1.Text) And IsNumeric(.TextBox2.Text) Then
                        'вызов модуля подсчета
                        Call SpecializedCounting
                        'вставляет разность (кол-во номеров) в лэйбл
                        .Label20 = CLng(.TextBox2.Text) - CLng(.TextBox1.Text)
                        'общее кол-во xxx
                        .Label19 = ModuleOfSpecialCounting.totalNumberOfCopies
                        'xxxxxxx
                        .Label18 = ModuleOfSpecialCounting.totalNumberOfHemmedCopies
                        'xxxxxxx
                        .Label17 = ModuleOfSpecialCounting.totalNumberOfDestroyedCopies
                        'xxxxxxx
                        .Label15 = ModuleOfSpecialCounting.totalNumberOfRepostedCopies
                        'xxxxxxx
                        .Label10 = ModuleOfSpecialCounting.totalNumberOfCopiesSentIrrevocably
                        'xxxxxxx
                        .Label16 = ModuleOfSpecialCounting.totalNumberOfCopiesPutOnInventory
                  Else
                        MsgBox "Числовой диапазон не может быть строковым литералом", vbExclamation, "Модуль специального подсчета"
                        .TextBox1.Text = ""
                        .TextBox2.Text = ""
                  End If
            Else
                  'вызов модуля подсчета
                  Call SpecializedCounting
                  .Label20 = "-"
                   'xxxxxxxx
                  .Label19 = ModuleOfSpecialCounting.totalNumberOfCopies
                  'xxxxxxxx
                  .Label18 = ModuleOfSpecialCounting.totalNumberOfHemmedCopies
                  'xxxxxxxx
                  .Label17 = ModuleOfSpecialCounting.totalNumberOfDestroyedCopies
                  'xxxxxxxx
                  .Label15 = ModuleOfSpecialCounting.totalNumberOfRepostedCopies
                  'xxxxxxxx
                  .Label10 = ModuleOfSpecialCounting.totalNumberOfCopiesSentIrrevocably
                  'xxxxxxxx
                  .Label16 = ModuleOfSpecialCounting.totalNumberOfCopiesPutOnInventory
            End If
      End With
End Sub
'button close
Private Sub CommandButton2_Click()
      Me.Hide
End Sub
'user form completion
Private Sub UserForm_Initialize()

      'массив для получения всех листов из книги
      Dim sheetNameArray() As String
      
      'индекс массива
      Dim i As Integer
      i = 0
      
      Dim sheet As Variant
      'переменная для проверки названий месяцев
      Dim numberMonth As Variant
      
      With Me
            If ThisWorkbook.Worksheets.Count < 2 Then
                  MsgBox "Добавьте листы для работы"
                  .CommandButton1.Enabled = False
                  .TextBox1.Enabled = False
                  .TextBox2.Enabled = False
                  Exit Sub
            ElseIf ThisWorkbook.Worksheets.Count < 13 Then
                  Select Case MsgBox("В данную книгу загружен не весь период. Выполнение общего подсчета за год будет некорректным. " & Chr(10) & Chr(10) & "Продолжить?", vbYesNo + vbExclamation + vbDefaultButton2, "Модуль специального подсчета")
                  Case vbYes
                        ReDim sheetNameArray(ThisWorkbook.Worksheets.Count - 2)
                  Case vbNo
                        .CommandButton1.Enabled = False
                        .TextBox1.Enabled = False
                        .TextBox2.Enabled = False
                        Exit Sub
                  End Select
            Else
                  ReDim sheetNameArray(ThisWorkbook.Worksheets.Count - 2)
            End If
      
            For Each sheet In ThisWorkbook.Worksheets
                  If sheet.Name <> "Программный лист" Then
                        On Error Resume Next
                        
                        numberMonth = month(DateValue("08/" & sheet.Name & "/1998"))
                        
                        If IsEmpty(numberMonth) Then
                              MsgBox "Переменуйте лист "" & sheet.Name & "" в словесно-действительное значение месяца, иначе работа будет невозможна", vbCritical, "Модуль специального подсчета"
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
End Sub
