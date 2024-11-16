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

Public Sub MyCustomMacro()
    On Error GoTo ErrHandler
    MsgBox "Hello from the custom context menu!"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub
