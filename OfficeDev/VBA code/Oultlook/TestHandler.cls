VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mItem As MailItem
Attribute mItem.VB_VarHelpID = -1
Public WithEvents myExplorer As Explorer
Attribute myExplorer.VB_VarHelpID = -1
Dim WithEvents colItems As Outlook.Items
Attribute colItems.VB_VarHelpID = -1
Public WithEvents objReminders As Outlook.Reminders
Attribute objReminders.VB_VarHelpID = -1

' #########
' destructor.
' #########
Private Sub Class_Terminate()
    If Not (mItem Is Nothing) Then
        Set mItem = Nothing
    End If

    If Not (myExplorer Is Nothing) Then
        Set myExplorer = Nothing
    End If
    
    If Not (objReminders Is Nothing) Then
        Set objReminders = Nothing
    End If
    
    If Not (colItems Is Nothing) Then
        Set colItems = Nothing
    End If
End Sub

Private Sub colItems_ItemChange(ByVal Item As Object)
 MsgBox "Change"
 Dim mySel As Selection

    Set mySel = myExplorer.Selection

    ' /* There is only one item in the browser window that is selected. */
    If mySel.Count = 1 Then
        Dim objItem As Object

        Set objItem = mySel.Item(1)

        ' /* The selected item is a mail. */
        If objItem.Class = olMail Then
            ' /* Dynamically register the event of a message */
            Set mItem = objItem
        End If
    End If

    Set mySel = Nothing
    Set objItem = Nothing
End Sub

Private Sub mItem_PropertyChange(ByVal Name As String)
  MsgBox "Here"
 'to do sth
End Sub

Private Sub mItem_Reply(ByVal Response As Object, _
                        Cancel As Boolean)
     
    Response.Subject = "[Modify message subject to the specified name]"
End Sub

' ##################################################
' Event: Occurs when you select another or more Microsoft Outlook items.
' ##################################################
Private Sub myExplorer_SelectionChange()
    Dim mySel As Selection

    Set mySel = myExplorer.Selection

    ' /* There is only one item in the browser window that is selected. */
    If mySel.Count = 1 Then
        Dim objItem As Object

        Set objItem = mySel.Item(1)

        ' /* The selected item is a mail. */
        If objItem.Class = olMail Then
            ' /* Dynamically register the event of a message */
            Set mItem = objItem
        End If
    End If

    Set mySel = Nothing
    Set objItem = Nothing
End Sub

' #####################################
' Forces the selected item in the browser window to change the event.
' #####################################
Public Sub ForceSelectionChange()
    Call myExplorer_SelectionChange
End Sub

Public Sub ForceMailItemChange()
    Call colItems_ItemChange
End Sub

Private Sub objReminders_ReminderFire(ByVal ReminderObject As Reminder)
ReminderObject.Item.Display
' to do send mail
End Sub
