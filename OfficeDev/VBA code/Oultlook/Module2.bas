Attribute VB_Name = "Module2"
Option Explicit

Dim rHandler As TestHandler

' #############
' Hook up the mail reply..
' #############
Public Sub HookTest()
    Set rHandler = New TestHandler
    Set rHandler.myExplorer = Application.ActiveExplorer
    Set rHandler.objReminders = Application.Reminders
    rHandler.ForceSelectionChange
    rHandler.ForceMailItemChange
    
End Sub

' #############
' Unhook the mail reply.
' #############
Public Sub UnhookTest()
    Set rHandler = Nothing
End Sub
