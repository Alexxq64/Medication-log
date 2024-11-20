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
    Set dateComboBox = Me.cmbDate
    
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
    Set pillComboBox = Me.cmbName

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


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    ' Проверка обязательных полей
    If cmbName.Text = "" Then
        MsgBox "Введите название лекарства!", vbExclamation
        Exit Sub
    End If
    If IsDate(cmbDate.Value) = False Then
        MsgBox "Введите корректную дату начала!", vbExclamation
        Exit Sub
    End If

'    ' Считываем данные с формы
'    Dim pillName As String
'    Dim startDate As Date
    Dim duration As Integer
'    Dim dosageMorning As Double
'    Dim dosageAfternoon As Double
'    Dim dosageEvening As Double
'    Dim dosageNight As Double
    Dim repeateDays As Integer
    
    Dim record As SingleDoseRecord
    record.DateScheduled = CDate(cmbDate.Value)
    record.Medicine = cmbName.Text
    record.Dosage = ""
    record.Morning = FixDecimalSeparator(txtMorning.Value)
    record.Afternoon = FixDecimalSeparator(txtAfternoon.Value)
    record.Evening = FixDecimalSeparator(txtEvening.Value)
    record.Night = FixDecimalSeparator(txtNight.Value)
    record.InStock = True
    record.Class = ""
    record.Notes = ""


    

'    pillName = cmbName.Text
'    startDate = CDate(cmbDate.Value)
    duration = CInt(txtDuration.Value)
'
'    dosageMorning = FixDecimalSeparator(txtMorning.Value)
'    dosageAfternoon = FixDecimalSeparator(txtAfternoon.Value)
'    dosageEvening = FixDecimalSeparator(txtEvening.Value)
'    dosageNight = FixDecimalSeparator(txtNight.Value)
'
    ' Проверяем количество повторов (целое число)
    If IsNumeric(txtRepeateDays.Text) And CLng(txtRepeateDays.Text) >= 0 Then
        repeateDays = CLng(txtRepeateDays.Text)
    Else
        repeateDays = 0
    End If

    ' Вызываем AddNewSchedule
    MedicationLog.AddDoseSchedule record, duration, repeateDays
'    MedicationLog.AddNewSchedule pillName, startDate, Duration, dosageMorning, dosageAfternoon, dosageEvening, dosageNight, repeateDays

    ' Закрываем форму
    Unload Me
End Sub


Private Sub txtDuration_AfterUpdate()
    Dim days As Integer

    On Error Resume Next ' Игнорируем ошибки ввода
    days = CInt(txtDuration.Value) ' Преобразуем текст в число

    ' Проверяем диапазон значений
    If days < spnDuration.Min Or days > spnDuration.Max Then
        MsgBox "Введите значение от " & spnDuration.Min & " до " & spnDuration.Max & ".", vbExclamation
        txtDuration.Value = spnDuration.Value ' Возвращаем предыдущее значение
    Else
        spnDuration.Value = days ' Синхронизация SpinButton
    End If

    On Error GoTo 0 ' Включаем обработку ошибок обратно
End Sub

Private Sub spnDuration_Change()
    ' Синхронизация txtDuration с spnDuration
    txtDuration.Value = spnDuration.Value
End Sub

Private Sub chkEveryDay_Click()
    ' Выбор "Каждый день"
    SetCheckBoxState chkEveryDay
End Sub

Private Sub chkEveryOtherDay_Click()
    ' Выбор "Через день"
    SetCheckBoxState chkEveryOtherDay
End Sub

Private Sub chkRepeateDays_Click()
    ' Выбор "Каждые несколько дней"
    SetCheckBoxState chkRepeateDays
End Sub

' Общая процедура для управления состоянием чекбоксов
Private Sub SetCheckBoxState(selectedCheckBox As MSForms.CheckBox)
    ' Сбрасываем все чекбоксы
    chkEveryDay.Value = False
    chkEveryOtherDay.Value = False
    chkRepeateDays.Value = False

    ' Активируем выбранный чекбокс
    selectedCheckBox.Value = True

    ' Настройка TextBox и SpinButton в зависимости от выбранного чекбокса
    If selectedCheckBox Is chkEveryDay Then
        spnRepeateDays.Enabled = False
        txtRepeateDays.Enabled = False
        txtRepeateDays.Text = 1 ' Каждый день
    ElseIf selectedCheckBox Is chkEveryOtherDay Then
        spnRepeateDays.Enabled = False
        txtRepeateDays.Enabled = False
        txtRepeateDays.Text = 2 ' Через день
    ElseIf selectedCheckBox Is chkRepeateDays Then
        spnRepeateDays.Enabled = True
        txtRepeateDays.Enabled = True
        txtRepeateDays.Text = 3 ' Каждые несколько дней
    End If
End Sub


' Событие изменения SpinButton
Private Sub spnRepeateDays_Change()
    ' При изменении значения SpinButton обновляем TextBox
    txtRepeateDays.Text = spnRepeateDays.Value
End Sub

' Событие изменения TextBox
Private Sub txtRepeateDays_Change()
    Dim days As Long
    On Error Resume Next ' Игнорировать ошибки, если текст невозможно преобразовать в число

    ' Проверяем, что пользователь ввел число
    days = CLng(txtRepeateDays.Text)
    
    ' Ограничиваем значение диапазоном SpinButton
    If days < spnRepeateDays.Min Then days = spnRepeateDays.Min
    If days > spnRepeateDays.Max Then days = spnRepeateDays.Max

    ' Устанавливаем значение в SpinButton
    spnRepeateDays.Value = days

    On Error GoTo 0 ' Восстановить обработку ошибок
End Sub

Private Sub UserForm_Activate()
    
    UpdateDateList
    UpdatePillsList
End Sub

Private Sub UserForm_Initialize()
    ' Устанавливаем начальное значение для txtDuration и spnDuration
    txtDuration.Value = 10
    spnDuration.Min = 0
    spnDuration.Max = 1000
    spnDuration.Value = 10
    ' Устанавливаем начальный выбор на "Каждый день"
    chkEveryDay.Caption = "Каждый день"
    chkEveryOtherDay.Caption = "Через день"
    chkRepeateDays.Caption = "Каждые несколько дней"
    
    ' Устанавливаем начальные значения
    spnRepeateDays.Min = 1          ' Минимальное значение
    spnRepeateDays.Max = 31         ' Максимальное значение
    spnRepeateDays.Value = 1        ' Начальное значение

    txtRepeateDays.Text = spnRepeateDays.Value ' Синхронизация TextBox с SpinButton
    
    chkEveryDay.Value = True
    chkEveryOtherDay.Value = False
    chkRepeateDays.Value = False

    ' Отключаем элементы, связанные с "Каждые несколько дней"
    spnRepeateDays.Enabled = False
    txtRepeateDays.Enabled = False

End Sub
