Attribute VB_Name = "ModuleOfGlobalFunction"
Option Explicit
'Функция копирования данных с Excel в Word документы
Function copyDataToWord(wordDoc As Word.Document, data As String, numberLine As Integer, stringInWordDoc As Integer, Optional statusEnd As Variant, Optional note As Variant)
      With wordDoc.Tables(1)
            .cell(stringInWordDoc, 2).Range.Text = data
            .cell(stringInWordDoc, 1).Range.Text = numberLine
        
            'если переданы опциональные аргументы
            If Not IsMissing(statusEnd) Then
                  .cell(stringInWordDoc, 3).Range.Text = statusEnd
                  
            End If
            
            If Not IsMissing(note) Then
                  .cell(stringInWordDoc, 4).Range.Text = note
            End If
    End With
End Function
'функция отбора номера
Function getNumber(str As String) As String
      Dim number As String
      Dim i As Integer
    
      For i = 1 To Len(str)
            If IsNumeric(Mid(str, i, 1)) Then
                  number = number + Mid(str, i, 1)
            End If
      Next i
    
      'return
      getNumber = number
End Function
'функция отбора ФИО обрабатывавшего без пробелов
Function getFIONotGap(str As String) As String
      Dim number As String
      number = ""

      For i = 1 To Len(str)
            If Mid(str, i, 1) <> " " Then
                  number = number + Mid(str, i, 1)
            End If
      Next i
    
    'return
    getFIONotGap = number
End Function
'Функция проверки открытия word документа
Function fileWordDocIsOpen(fullDocName As String, firstOpen As Boolean) As Boolean
      Dim wordApp As Word.Application
      Dim wordDoc As Word.Document
      
      'return
      fileWordDocIsOpen = False
      
      'error handler
      On Error Resume Next
      
      'control
      Set wordApp = GetObject(, "Word.Application")
      If wordApp Is Nothing Then Exit Function
      
      For Each wordDoc In wordApp.Documents
            If wordDoc.FullName = fullDocName Then
                  If Not firstOpen Then
                        MsgBox "Закройте Word документ с именем: " & wordDoc.Name & ", удалите его и запустите программу снова", vbInformation, "Функция проверки открытия Word файлов"
                  End If
                  Set wordApp = Nothing
                  Set wordDoc = Nothing
                  fileWordDocIsOpen = True
                  Exit Function
            End If
      Next
End Function
'Функция проверки наличия созданного программного учетного блока на листе
Function findTextBoxInSheet(sheet As Worksheet) As Boolean
      'переменная фигуры
      Dim shapeProgramm As Shape
      
      findTextBoxInSheet = False
      
      For Each shapeProgramm In sheet.Shapes
            If shapeProgramm.Name = "programm figure" Then
                  findTextBoxInSheet = True
                  Exit Function
            End If
      Next
End Function
