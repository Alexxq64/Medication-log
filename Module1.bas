Attribute VB_Name = "Module1"
Option Explicit

Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��������������
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

