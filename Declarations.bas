Attribute VB_Name = "Declarations"
'@IgnoreModule MoveFieldCloserToUsage, UseMeaningfulName, EncapsulatePublicField
'@Folder("��������")
Option Explicit

Public MainLog As Range
Public TodayRowNumber As Range
Public Summary1 As Range
Public TotalColumns As Range
Public TotalMedicines As Range
Public TotalRows As Range
Public CFRulesQtty As Long

Public Type MedicationInfo
    Name As String ' �������� ���������
    Dosage As String ' ���������
    Morning As Double ' ���� �����
    Afternoon As Double ' ���� ����
    Evening As Double ' ���� �������
    Night As Double ' ���� �����
    duration As Integer ' ����������������� ����� (� ����)
    RepeatDays As Integer ' ������������� (������ ���� = 1, ����� ���� = 2 � �.�.)
    Class As String ' ����� ���������
    TrackStock As Boolean ' ����������� �������
    InFirstAidKit As Boolean ' ������ � ������ �������
End Type

Public Type SingleDoseRecord
    DateScheduled As Date ' ���� ������
    Medicine As String ' �������� ���������
    Dosage As String ' ��������� (�����, �������� "10 ��" ��� "1 ��������")
    Morning As Double ' ����
    Afternoon As Double ' ����
    Evening As Double ' �����
    Night As Double ' ����
    InStock As Boolean ' ���������, ����� �� ����������� �������
    Class As String '����� ���������
    Notes As String ' ����������
End Type



Public Sub MyCustomMacro()
    On Error GoTo ErrHandler
    MsgBox "Hello from the custom context menu!"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub

Function FixDecimalSeparator(inputText As String) As Double
    ' �������� ����� �� ����������� � ����������� �� ������
    inputText = Replace(inputText, ".", Application.DecimalSeparator)
    
    ' ���������, �������� �� ����� ������, � ������� ��������
    If IsNumeric(inputText) Then
        FixDecimalSeparator = CDbl(inputText)
    Else
        FixDecimalSeparator = 0 ' ���� ����� �� �������� ������, ������� 0
    End If
End Function

