Attribute VB_Name = "Declarations"
'@IgnoreModule MoveFieldCloserToUsage, UseMeaningfulName, EncapsulatePublicField
'@Folder("Таблетки")
Option Explicit

Public MainLog As Range
Public TodayRowNumber As Range
Public Summary1 As Range
Public TotalColumns As Range
Public TotalMedicines As Range
Public TotalRows As Range
Public CFRulesQtty As Long

Public Type MedicationInfo
    Name As String ' Название лекарства
    Dosage As String ' Дозировка
    Morning As Double ' Доза утром
    Afternoon As Double ' Доза днем
    Evening As Double ' Доза вечером
    Night As Double ' Доза ночью
    duration As Integer ' Продолжительность курса (в днях)
    RepeatDays As Integer ' Периодичность (каждый день = 1, через день = 2 и т.д.)
    Class As String ' Класс препарата
    TrackStock As Boolean ' Отслеживать остаток
    InFirstAidKit As Boolean ' Входит в состав аптечки
End Type

Public Type SingleDoseRecord
    DateScheduled As Date ' Дата приема
    Medicine As String ' Название лекарства
    Dosage As String ' Дозировка (текст, например "10 мг" или "1 таблетка")
    Morning As Double ' Утро
    Afternoon As Double ' День
    Evening As Double ' Вечер
    Night As Double ' Ночь
    InStock As Boolean ' Указывает, нужно ли отслеживать остаток
    Class As String 'Класс препарата
    Notes As String ' Примечания
End Type



Public Sub MyCustomMacro()
    On Error GoTo ErrHandler
    MsgBox "Hello from the custom context menu!"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub

Function FixDecimalSeparator(inputText As String) As Double
    ' Заменить точку на разделитель в зависимости от локали
    inputText = Replace(inputText, ".", Application.DecimalSeparator)
    
    ' Проверить, является ли текст числом, и вернуть значение
    If IsNumeric(inputText) Then
        FixDecimalSeparator = CDbl(inputText)
    Else
        FixDecimalSeparator = 0 ' Если текст не является числом, вернуть 0
    End If
End Function

