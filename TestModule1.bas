Attribute VB_Name = "TestModule1"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


Sub TestAddDoseSchedule()

    Dim record As SingleDoseRecord
    record.DateScheduled = Date
    record.Medicine = "Аспирин"
    record.Dosage = "500 мг"
    record.Morning = 1
    record.Afternoon = 0.5
    record.Evening = 0
    record.Night = 0
    record.InStock = False
    record.Class = "Анальгетик"
    record.Notes = "Примечание"

    MedicationLog.AddDoseSchedule record, 10, 1

End Sub


