Attribute VB_Name = "Module1"
Option Explicit

Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' автозаполнение
'

'
    Range("I3942").Select
    Selection.AutoFill Destination:=Range("I3942:I3952")
    Range("I3942:I3952").Select
End Sub

'
'Sub CreateActiveCellName()
'    ThisWorkbook.Names.Add Name:="ActiveCell", RefersTo:="=INDIRECT(""RC"",FALSE)"
'End Sub

