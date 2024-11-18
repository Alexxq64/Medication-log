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
    Set dateComboBox = Me.DateListComboBox
    
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
    Set pillComboBox = Me.PillsListComboBox

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




Private Sub DurationBox_AfterUpdate()
    Dim days As Integer

    On Error Resume Next ' ���������� ������ �����
    days = CInt(DurationBox.Value) ' ����������� ����� � �����

    ' ��������� �������� ��������
    If days < DurationSpinButton.Min Or days > DurationSpinButton.Max Then
        MsgBox "������� �������� �� " & DurationSpinButton.Min & " �� " & DurationSpinButton.Max & ".", vbExclamation
        DurationBox.Value = DurationSpinButton.Value ' ���������� ���������� ��������
    Else
        DurationSpinButton.Value = days ' ������������� SpinButton
    End If

    On Error GoTo 0 ' �������� ��������� ������ �������
End Sub

Private Sub DurationSpinButton_Change()
    ' ������������� DurationBox � DurationSpinButton
    DurationBox.Value = DurationSpinButton.Value
End Sub

Private Sub chkEveryDay_Click()
    ' ����� "������ ����"
    SetCheckBoxState chkEveryDay
End Sub

Private Sub chkEveryOtherDay_Click()
    ' ����� "����� ����"
    SetCheckBoxState chkEveryOtherDay
End Sub

Private Sub chkCustomDays_Click()
    ' ����� "������ ��������� ����"
    SetCheckBoxState chkCustomDays
End Sub

' ����� ��������� ��� ���������� ���������� ���������
Private Sub SetCheckBoxState(selectedCheckBox As MSForms.CheckBox)
    ' ���������� ��� ��������
    chkEveryDay.Value = False
    chkEveryOtherDay.Value = False
    chkCustomDays.Value = False

    ' ���������� ��������� �������
    selectedCheckBox.Value = True

    ' ��������� TextBox � SpinButton � ����������� �� ���������� ��������
    If selectedCheckBox Is chkEveryDay Then
        spnCustomDays.Enabled = False
        txtCustomDays.Enabled = False
        txtCustomDays.Text = 1 ' ������ ����
    ElseIf selectedCheckBox Is chkEveryOtherDay Then
        spnCustomDays.Enabled = False
        txtCustomDays.Enabled = False
        txtCustomDays.Text = 2 ' ����� ����
    ElseIf selectedCheckBox Is chkCustomDays Then
        spnCustomDays.Enabled = True
        txtCustomDays.Enabled = True
        txtCustomDays.Text = spnCustomDays.Value ' ������ ��������� ����
    End If
End Sub

Private Sub OKButton_Click()
    AddNewSchedule "�������", #11/17/2024#, 3, 1, 1, 1, 1, 1
    Unload Me
End Sub

' ������� ��������� SpinButton
Private Sub spnCustomDays_Change()
    ' ��� ��������� �������� SpinButton ��������� TextBox
    txtCustomDays.Text = spnCustomDays.Value
End Sub

' ������� ��������� TextBox
Private Sub txtCustomDays_Change()
    Dim days As Long
    On Error Resume Next ' ������������ ������, ���� ����� ���������� ������������� � �����

    ' ���������, ��� ������������ ���� �����
    days = CLng(txtCustomDays.Text)
    
    ' ������������ �������� ���������� SpinButton
    If days < spnCustomDays.Min Then days = spnCustomDays.Min
    If days > spnCustomDays.Max Then days = spnCustomDays.Max

    ' ������������� �������� � SpinButton
    spnCustomDays.Value = days

    On Error GoTo 0 ' ������������ ��������� ������
End Sub

Private Sub UserForm_Activate()
    
    UpdateDateList
    UpdatePillsList
End Sub

Private Sub UserForm_Initialize()
    ' ������������� ��������� �������� ��� DurationBox � DurationSpinButton
    DurationBox.Value = 10
    DurationSpinButton.Min = 0
    DurationSpinButton.Max = 1000
    DurationSpinButton.Value = 10
    ' ������������� ��������� ����� �� "������ ����"
    chkEveryDay.Caption = "������ ����"
    chkEveryOtherDay.Caption = "����� ����"
    chkCustomDays.Caption = "������ ��������� ����"
    
    ' ������������� ��������� ��������
    spnCustomDays.Min = 3          ' ����������� ��������
    spnCustomDays.Max = 31         ' ������������ ��������
    spnCustomDays.Value = 3        ' ��������� ��������

    txtCustomDays.Text = spnCustomDays.Value ' ������������� TextBox � SpinButton
    
    chkEveryDay.Value = True
    chkEveryOtherDay.Value = False
    chkCustomDays.Value = False

    ' ��������� ��������, ��������� � "������ ��������� ����"
    spnCustomDays.Enabled = False
    txtCustomDays.Enabled = False

End Sub
