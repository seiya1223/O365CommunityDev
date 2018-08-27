Attribute VB_Name = "Module1"
Sub Test()

Dim MyItem As Outlook.MailItem

With MyItem
            .To = EmailAddr
            .Subject = sSubject
            .SentOnBehalfOfName = "SoAndSo@sample.com"
            .HTMLBody = Msg
            .Importance = olImportanceHigh
            .FlagStatus = olFlagMarked
            .FlagRequest = "Follow up"
            .FlagDueBy = Range("F2").Value & " 10:08 AM"
End With

End Sub
