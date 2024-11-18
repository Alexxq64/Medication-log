VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PillHandlingForm 
   Caption         =   "Схема приема"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   OleObjectBlob   =   "PillHandlingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PillHandlingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UpdateDateList(Optional ByVal dateValue As Variant = -1)
    Dim ws As Worksheet
    Dim firstDate As Date, lastDate As Date
    Dim currentDate As Date
    Dim dateList() As Date
    Dim comboBox As MSForms.comboBox
    Dim totalDays As Long
    Dim i As Long

    ' Устанавливаем ссылку на лист "Medication log"
    Set ws = ThisWorkbook.Worksheets("Medication log")
    
    ' Устанавливаем ссылку на ComboBox
    Set dateComboBox = Me.DateListComboBox
    
    ' Очищаем ComboBox перед обновлением
    dateComboBox.Clear
    
    ' Получаем первую и последнюю дату
    firstDate = ws.Cells(2, 1).Value
    lastDate = ws.Cells(ws.Cells(ws.Rows.count, 1).End(xlUp).Row, 1).Value
    
    ' Вычисляем количество дней между первой и последней датой
    totalDays = DateDiff("d", firstDate, lastDate) + 1 ' +1 чтобы включить последнюю дату

    ' Выделяем массив нужного размера
    ReDim dateList(1 To totalDays)

    ' Заполняем массив датами от первой до последней
    currentDate = firstDate
    For i = 1 To totalDays
        dateList(i) = currentDate
        currentDate = currentDate + 1            ' Переход к следующему дню
    Next i
    
    ' Добавляем массив дат в ComboBox
    dateComboBox.List = WorksheetFunction.Transpose(dateList)
    
    ' Если параметр dateValue не передан, то используем дату из ActiveCell
    '    If dateValue = -1 Then
    '        currentRowDate = ws.Cells(ActiveCell.Row, 1).Value
    '        If IsDate(currentRowDate) Then
    '            dateComboBox.Value = Format(currentRowDate, "dd-mm-yyyy")
    '        Else
    '            dateComboBox.Value = Format(Date, "dd-mm-yyyy")
    '        End If
    '    Else
    '        dateComboBox.Value = Format(dateValue, "dd-mm-yyyy")
    '    End If

    If dateValue = -1 Then
        currentRowDate = ws.Cells(ActiveCell.Row, 1).Value
        If IsDate(currentRowDate) Then
            dateValue = Format(currentRowDate, "dd-mm-yyyy")
        Else
            dateValue = Format(Date, "dd-mm-yyyy")
        End If
    End If
    dateComboBox.Value = Format(dateValue, "dd-mm-yyyy")

        
    Dim calculatedIndex As Long
    ' Рассчитываем индекс как разницу в днях между dateValue и первой датой
    calculatedIndex = DateDiff("d", dateList(LBound(dateList)), dateValue)

    ' Проверяем, что calculatedIndex находится в пределах массива
    If calculatedIndex >= 0 And calculatedIndex <= UBound(dateList) - LBound(dateList) Then
        dateComboBox.ListIndex = calculatedIndex
    End If
End Sub

Sub UpdatePillsList(Optional ByVal pillValue As Variant = "")
    Dim ws As Worksheet
    Dim cell As Range
    Dim pillComboBox As MSForms.comboBox
    Dim pillsCollection As Collection
    Dim pillsArray() As Variant
    Dim currentRowPill As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
    ' Указываем лист с данными
    Set ws = ThisWorkbook.Worksheets("Medication log")
    
    ' Указываем ComboBox для лекарств
    Set pillComboBox = Me.PillsListComboBox

    ' Очищаем ComboBox
    pillComboBox.Clear

    ' Создаем коллекцию для уникальных лекарств
    Set pillsCollection = New Collection
    
    ' Собираем уникальные лекарства в коллекцию
    On Error Resume Next
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            pillsCollection.Add cell.Value, CStr(cell.Value) ' Добавляем уникальные лекарства
        End If
    Next cell
    On Error GoTo 0

    ' Преобразуем коллекцию в массив
    ReDim pillsArray(1 To pillsCollection.count)
    For i = 1 To pillsCollection.count
        pillsArray(i) = pillsCollection(i)
    Next i

    ' Сортировка массива по алфавиту (метод пузырьковой сортировки)
    For i = LBound(pillsArray) To UBound(pillsArray) - 1
        For j = i + 1 To UBound(pillsArray)
            If pillsArray(i) > pillsArray(j) Then
                ' Меняем элементы местами
                temp = pillsArray(i)
                pillsArray(i) = pillsArray(j)
                pillsArray(j) = temp
            End If
        Next j
    Next i

    ' Устанавливаем массив как источник для ComboBox
    pillComboBox.List = Application.Transpose(pillsArray)

    ' Если параметр pillValue не передан, то используем лекарство из ActiveCell
    If pillValue = "" Then
        currentRowPill = ws.Cells(ActiveCell.Row, 2).Value
        If Not IsEmpty(currentRowPill) Then
            pillComboBox.Value = currentRowPill
        End If
    Else
        pillComboBox.Value = pillValue
    End If
End Sub




Private Sub DurationBox_AfterUpdate()
    Dim days As Integer

    On Error Resume Next ' Игнорируем ошибки ввода
    days = CInt(DurationBox.Value) ' Преобразуем текст в число

    ' Проверяем диапазон значений
    If days < DurationSpinButton.Min Or days > DurationSpinButton.Max Then
        MsgBox "Введите значение от " & DurationSpinButton.Min & " до " & DurationSpinButton.Max & ".", vbExclamation
        DurationBox.Value = DurationSpinButton.Value ' Возвращаем предыдущее значение
    Else
        DurationSpinButton.Value = days ' Синхронизация SpinButton
    End If

    On Error GoTo 0 ' Включаем обработку ошибок обратно
End Sub

Private Sub DurationSpinButton_Change()
    ' Синхронизация DurationBox с DurationSpinButton
    DurationBox.Value = DurationSpinButton.Value
End Sub

Private Sub chkEveryDay_Click()
    ' Выбор "Каждый день"
    SetCheckBoxState chkEveryDay
End Sub

Private Sub chkEveryOtherDay_Click()
    ' Выбор "Через день"
    SetCheckBoxState chkEveryOtherDay
End Sub

Private Sub chkCustomDays_Click()
    ' Выбор "Каждые несколько дней"
    SetCheckBoxState chkCustomDays
End Sub

' Общая процедура для управления состоянием чекбоксов
Private Sub SetCheckBoxState(selectedCheckBox As MSForms.CheckBox)
    ' Сбрасываем все чекбоксы
    chkEveryDay.Value = False
    chkEveryOtherDay.Value = False
    chkCustomDays.Value = False

    ' Активируем выбранный чекбокс
    selectedCheckBox.Value = True

    ' Включаем/выключаем поля на основе выбранного варианта
    If selectedCheckBox Is chkCustomDays Then
        spnCustomDays.Enabled = True
        txtCustomDays.Enabled = True
    Else
        spnCustomDays.Enabled = False
        txtCustomDays.Enabled = False
    End If
End Sub


Private Sub OKButton_Click()
    Unload Me
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Activate()
    
    UpdateDateList
    UpdatePillsList
End Sub

Private Sub UserForm_Initialize()
    ' Устанавливаем начальное значение для DurationBox и DurationSpinButton
    DurationBox.Value = 10
    DurationSpinButton.Min = 0
    DurationSpinButton.Max = 1000
    DurationSpinButton.Value = 10
    ' Устанавливаем начальный выбор на "Каждый день"
    chkEveryDay.Caption = "Каждый день"
    chkEveryOtherDay.Caption = "Через день"
    chkCustomDays.Caption = "Каждые несколько дней"

    chkEveryDay.Value = True
    chkEveryOtherDay.Value = False
    chkCustomDays.Value = False

    ' Отключаем элементы, связанные с "Каждые несколько дней"
    spnCustomDays.Enabled = False
    txtCustomDays.Enabled = False

End Sub
