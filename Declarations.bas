Attribute VB_Name = "Declarations"
'@IgnoreModule MoveFieldCloserToUsage, UseMeaningfulName, EncapsulatePublicField
'@Folder("“аблетки")
Option Explicit

Public MainLog As Range
Public TodayRowNumber As Range
Public Summary1 As Range
Public TotalColumns As Range
Public TotalMedicines As Range
Public TotalRows As Range

Public Sub MyCustomMacro()
    On Error GoTo ErrHandler
    MsgBox "Hello from the custom context menu!"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub

Public Sub AddNewSchedule(Medicine As String, startDate As Date, _
                   duration As Integer, Morning As Integer, _
                   Afternoon As Integer, Evening As Integer, _
                   Night As Integer, skipDays As Integer)
                   
  
    toRow = TotalRows + 1
    For i = 0 To duration - 1
        CopyFromTo ActiveCell.Row, toRow + i, startDate + i * (skipDays + 1)
    Next i
                   
End Sub

Function FixDecimalSeparator(inputText As String) As Double
    ' «аменить точку на разделитель в зависимости от локали
    inputText = Replace(inputText, ".", Application.DecimalSeparator)
    
    ' ѕроверить, €вл€етс€ ли текст числом, и вернуть значение
    If IsNumeric(inputText) Then
        FixDecimalSeparator = CDbl(inputText)
    Else
        FixDecimalSeparator = 0 ' ≈сли текст не €вл€етс€ числом, вернуть 0
    End If
End Function

