VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PillHandlingForm 
   Caption         =   "����� ������"
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

    ' ������������� ������ �� ���� "Medication log"
    Set ws = ThisWorkbook.Worksheets("Medication log")
    
    ' ������������� ������ �� ComboBox
    Set dateComboBox = Me.cmbDate
    
    ' ������� ComboBox ����� �����������
    dateComboBox.Clear
    
    ' �������� ������ � ��������� ����
    firstDate = ws.Cells(2, 1).Value
    lastDate = ws.Cells(ws.Cells(ws.Rows.count, 1).End(xlUp).Row, 1).Value
    
    ' ��������� ���������� ���� ����� ������ � ��������� �����
    totalDays = DateDiff("d", firstDate, lastDate) + 1 ' +1 ����� �������� ��������� ����

    ' �������� ������ ������� �������
    ReDim dateList(1 To totalDays)

    ' ��������� ������ ������ �� ������ �� ���������
    currentDate = firstDate
    For i = 1 To totalDays
        dateList(i) = currentDate
        currentDate = currentDate + 1            ' ������� � ���������� ���
    Next i
    
    ' ��������� ������ ��� � ComboBox
    dateComboBox.List = WorksheetFunction.Transpose(dateList)
    
    ' ���� �������� dateValue �� �������, �� ���������� ���� �� ActiveCell
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
    ' ������������ ������ ��� ������� � ���� ����� dateValue � ������ �����
    calculatedIndex = DateDiff("d", dateList(LBound(dateList)), dateValue)

    ' ���������, ��� calculatedIndex ��������� � �������� �������
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
    
    ' ��������� ���� � �������
    Set ws = ThisWorkbook.Worksheets("Medication log")
    
    ' ��������� ComboBox ��� ��������
    Set pillComboBox = Me.cmbName

    ' ������� ComboBox
    pillComboBox.Clear

    ' ������� ��������� ��� ���������� ��������
    Set pillsCollection = New Collection
    
    ' �������� ���������� ��������� � ���������
    On Error Resume Next
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            pillsCollection.Add cell.Value, CStr(cell.Value) ' ��������� ���������� ���������
        End If
    Next cell
    On Error GoTo 0

    ' ����������� ��������� � ������
    ReDim pillsArray(1 To pillsCollection.count)
    For i = 1 To pillsCollection.count
        pillsArray(i) = pillsCollection(i)
    Next i

    ' ���������� ������� �� �������� (����� ����������� ����������)
    For i = LBound(pillsArray) To UBound(pillsArray) - 1
        For j = i + 1 To UBound(pillsArray)
            If pillsArray(i) > pillsArray(j) Then
                ' ������ �������� �������
                temp = pillsArray(i)
                pillsArray(i) = pillsArray(j)
                pillsArray(j) = temp
            End If
        Next j
    Next i

    ' ������������� ������ ��� �������� ��� ComboBox
    pillComboBox.List = Application.Transpose(pillsArray)

    ' ���� �������� pillValue �� �������, �� ���������� ��������� �� ActiveCell
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
    ' �������� ������������ �����
    If cmbName.Text = "" Then
        MsgBox "������� �������� ���������!", vbExclamation
        Exit Sub
    End If
    If IsDate(cmbDate.Value) = False Then
        MsgBox "������� ���������� ���� ������!", vbExclamation
        Exit Sub
    End If

'    ' ��������� ������ � �����
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
    ' ��������� ���������� �������� (����� �����)
    If IsNumeric(txtRepeateDays.Text) And CLng(txtRepeateDays.Text) >= 0 Then
        repeateDays = CLng(txtRepeateDays.Text)
    Else
        repeateDays = 0
    End If

    ' �������� AddNewSchedule
    MedicationLog.AddDoseSchedule record, duration, repeateDays
'    MedicationLog.AddNewSchedule pillName, startDate, Duration, dosageMorning, dosageAfternoon, dosageEvening, dosageNight, repeateDays

    ' ��������� �����
    Unload Me
End Sub


Private Sub txtDuration_AfterUpdate()
    Dim days As Integer

    On Error Resume Next ' ���������� ������ �����
    days = CInt(txtDuration.Value) ' ����������� ����� � �����

    ' ��������� �������� ��������
    If days < spnDuration.Min Or days > spnDuration.Max Then
        MsgBox "������� �������� �� " & spnDuration.Min & " �� " & spnDuration.Max & ".", vbExclamation
        txtDuration.Value = spnDuration.Value ' ���������� ���������� ��������
    Else
        spnDuration.Value = days ' ������������� SpinButton
    End If

    On Error GoTo 0 ' �������� ��������� ������ �������
End Sub

Private Sub spnDuration_Change()
    ' ������������� txtDuration � spnDuration
    txtDuration.Value = spnDuration.Value
End Sub

Private Sub chkEveryDay_Click()
    ' ����� "������ ����"
    SetCheckBoxState chkEveryDay
End Sub

Private Sub chkEveryOtherDay_Click()
    ' ����� "����� ����"
    SetCheckBoxState chkEveryOtherDay
End Sub

Private Sub chkRepeateDays_Click()
    ' ����� "������ ��������� ����"
    SetCheckBoxState chkRepeateDays
End Sub

' ����� ��������� ��� ���������� ���������� ���������
Private Sub SetCheckBoxState(selectedCheckBox As MSForms.CheckBox)
    ' ���������� ��� ��������
    chkEveryDay.Value = False
    chkEveryOtherDay.Value = False
    chkRepeateDays.Value = False

    ' ���������� ��������� �������
    selectedCheckBox.Value = True

    ' ��������� TextBox � SpinButton � ����������� �� ���������� ��������
    If selectedCheckBox Is chkEveryDay Then
        spnRepeateDays.Enabled = False
        txtRepeateDays.Enabled = False
        txtRepeateDays.Text = 1 ' ������ ����
    ElseIf selectedCheckBox Is chkEveryOtherDay Then
        spnRepeateDays.Enabled = False
        txtRepeateDays.Enabled = False
        txtRepeateDays.Text = 2 ' ����� ����
    ElseIf selectedCheckBox Is chkRepeateDays Then
        spnRepeateDays.Enabled = True
        txtRepeateDays.Enabled = True
        txtRepeateDays.Text = 3 ' ������ ��������� ����
    End If
End Sub


' ������� ��������� SpinButton
Private Sub spnRepeateDays_Change()
    ' ��� ��������� �������� SpinButton ��������� TextBox
    txtRepeateDays.Text = spnRepeateDays.Value
End Sub

' ������� ��������� TextBox
Private Sub txtRepeateDays_Change()
    Dim days As Long
    On Error Resume Next ' ������������ ������, ���� ����� ���������� ������������� � �����

    ' ���������, ��� ������������ ���� �����
    days = CLng(txtRepeateDays.Text)
    
    ' ������������ �������� ���������� SpinButton
    If days < spnRepeateDays.Min Then days = spnRepeateDays.Min
    If days > spnRepeateDays.Max Then days = spnRepeateDays.Max

    ' ������������� �������� � SpinButton
    spnRepeateDays.Value = days

    On Error GoTo 0 ' ������������ ��������� ������
End Sub

Private Sub UserForm_Activate()
    
    UpdateDateList
    UpdatePillsList
End Sub

Private Sub UserForm_Initialize()
    ' ������������� ��������� �������� ��� txtDuration � spnDuration
    txtDuration.Value = 10
    spnDuration.Min = 0
    spnDuration.Max = 1000
    spnDuration.Value = 10
    ' ������������� ��������� ����� �� "������ ����"
    chkEveryDay.Caption = "������ ����"
    chkEveryOtherDay.Caption = "����� ����"
    chkRepeateDays.Caption = "������ ��������� ����"
    
    ' ������������� ��������� ��������
    spnRepeateDays.Min = 1          ' ����������� ��������
    spnRepeateDays.Max = 31         ' ������������ ��������
    spnRepeateDays.Value = 1        ' ��������� ��������

    txtRepeateDays.Text = spnRepeateDays.Value ' ������������� TextBox � SpinButton
    
    chkEveryDay.Value = True
    chkEveryOtherDay.Value = False
    chkRepeateDays.Value = False

    ' ��������� ��������, ��������� � "������ ��������� ����"
    spnRepeateDays.Enabled = False
    txtRepeateDays.Enabled = False

End Sub
