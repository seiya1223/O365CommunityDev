Attribute VB_Name = "Module1"
 
Sub test()
Dim lastrow As Integer
lastrow = 20

Dim src As Range, out As Range, wks As worksheet

Set src = Sheets("Sheet2").Cells(12, 8)
Set out = Sheets("ShippingOutput").Cells(1, 1)
'src.Copy out
'ThisWorkbook.Sheets("ShippingOutput").Range("AD2").Value = PackId
'ThisWorkbook.Sheets("ShippingOutput").Range("AD2:AD" & iLast).Value = cell.Value + 1

'With ThisWorkbook.Sheets("ShippingOutput").Range("A1")
    '.AutoFill Destination:=.Resize(lastrow - 1, 1), Type:=xlFillDefault
    
    Set SourceRange = Worksheets("Sheet2").Range("A1:A2")
    Set out = Worksheets("ShippingOutput").Range("A1:A2")
    SourceRange.Copy out
    Set fillRange = Worksheets("ShippingOutput").Range("A1:A20")
    out.AutoFill Destination:=fillRange
'End With
 
 'ThisWorkbook.Sheets("ShippingOutput").Range("A1").Select
   'Selection.FillDown
   
' Selection.AutoFill Destination:=ThisWorkbook.Sheets("ShippingOutput").Range("A1:B1").Resize("A1:A" & lastrow), Type:=xlFillDefault
 
End Sub


Sub Test1()
Dim lastrow As Integer
lastrow = 20

Dim src As Range, out As Range, SourceRange As Range
    
Set SourceRange = Worksheets("Sheet2").Range("A1").Offset(11, 7) 'Cells(12, 8)
MsgBox SourceRange.Value
Set out = Worksheets("ShippingOutput").Range("A1")
SourceRange.Copy out
MsgBox out.Value
Set fillRange = Worksheets("ShippingOutput").Range("A1:A" & lastrow)
out.AutoFill Destination:=fillRange
'
End Sub

Sub Test2()
Dim lastrow As Integer
Dim src As Range, out As Range, SourceRange As Range

lastrow = 20 ' You may recalculate this value.

Set SourceRange = Worksheets("Sheet2").Cells(12, 8) 'Cells(12, 8)

'MsgBox SourceRange.Value
Set out = Worksheets("ShippingOutput").Range("A2")
SourceRange.Copy out
 
Set fillRange = Worksheets("ShippingOutput").Range("A2:A" & lastrow)
out.AutoFill Destination:=fillRange
 
End Sub

Sub Test3()
'Worksheets("Sheet3").Cells(13, 1).Select
Dim AppOutlook As Object
Dim MailOutlook As Object
Dim Emailto, ccto, sendfrom As String

Set AppOutlook = CreateObject("Outlook.Application")
Set MailOutlook = AppOutlook.CreateItem(0)
 
Emailto = Worksheets("Sheet3").Cells(1, 1).Value
ccto = Worksheets("Sheet3").Cells(2, 1).Value
sendfrom = "v-seiyas@microsoft.com"

With MailOutlook
    .SentOnBehalfOfName = sendfrom
    .To = Emailto
    .CC = ccto
    .BCC = ""
    .Subject = "Test"
    .BodyFormat = olFormatHTML
    .HTMLBody = "body here"
    '.Display
    .Send

Application.Wait (Now + TimeValue("0:00:30"))

End With
End Sub

 
 
 
Sub Test4()
Dim all As Range
Set all = Range("A8:F18")

For Each c In all
  'i = i + 1
Next

Dim lastrow As Integer
Dim Source As workbook
MsgBox Rows.Count
lastrow = ThisWorkbook.Worksheets("Sheet4").Cells(Rows.Count, 2).End(xlUp).row + 1
MsgBox lastrow
Set Source = ThisWorkbook
lastrow = IIf(Len(ThisWorkbook.Worksheets("Sheet4").Range("F18")), ThisWorkbook.Worksheets("Sheet4").Range("F18").row, ThisWorkbook.Worksheets("Sheet4").Range("F18").End(xlUp).row)

  
MsgBox lastrow

 
ThisWorkbook.Worksheets("Sheet4").Range("B" & lastrow).Formula = _
    Source.Worksheets("SUMMARY DATA SHEET").Range("A8").Value
ThisWorkbook.Worksheets("Sheet4").Range("D" & lastrow).Formula = _
    Source.Worksheets("SUMMARY DATA SHEET").Range("D8").Value
ThisWorkbook.Worksheets("Sheet4").Range("A" & lastrow).Formula = _
    Source.Worksheets("SUMMARY DATA SHEET").Range("B4").Value
ThisWorkbook.Worksheets("Sheet4").Range("E" & lastrow).Formula = _
    Source.Worksheets("SUMMARY DATA SHEET").Range("F8").Value
ThisWorkbook.Worksheets("Sheet4").Range("C" & lastrow).Formula = _
    Source.Worksheets("SUMMARY DATA SHEET").Range("E8").Value
End Sub


Sub TestV2()
Dim rng As Range
Dim selectedRange As Range

Set selectedRange = Selection
    For Each rng In selectedRange.Cells
        If Application.CalculationState = xlDone Then
        'FireValidate
             MsgBox "OK"
        End If
    Next rng

End Sub

Sub TestV2BackUp()
Dim rng As Range
Dim selectedRange As Range
Dim arc As Range

Dim ws As worksheet
Set ws = ActiveSheet
'MsgBox ws.Rows.Count

Set arc = ws.Rows.Select
'MsgBox ws.Rows.Count
Set selectedRange = ws.UsedRange.Rows.Count
MsgBox ws.Rows.Count

For Each rng In selectedRange.Cells
If Application.CalculationState = xlDone Then
'FireValidate
MsgBox "OK"
End If
    Next rng
    End
End Sub


Sub TestOutlook()
    Dim olApp As Outlook.Application, olNs As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder, Item As Outlook.MailItem
    Dim eFolder As Outlook.Folder '~~> additional declaration
    Dim i As Long
    Dim x As Date, ws As worksheet '~~> declare WS variable instead
    Dim lrow As Long '~~> additional declaration
    Dim MessageInfo
    Dim Result
    Set ws = ActiveSheet '~~> or you can be more explicit using the next line
    'Set ws = Thisworkbook.Sheets("YourTargetSheet")
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    x = Date

    For Each eFolder In olNs.GetDefaultFolder(olFolderInbox).Folders
        'Debug.Print eFolder.Name
        Set olFolder = olNs.GetDefaultFolder(olFolderInbox).Folders(eFolder.Name)
        For i = olFolder.Items.Count To 1 Step -1
            If TypeOf olFolder.Items(i) Is MailItem Then
                Set Item = olFolder.Items(i)
                'MsgBox Item.Body
                'filter (Item)
                'If InStr(Item.Subject, "Test download") > 0 Then
                   ' MsgBox "Here"
                   '                     MessageInfo = "" & _
                    '        "Sender : " & Item.SenderEmailAddress & vbCrLf & _
                    '        "Sent : " & Item.SentOn & vbCrLf & _
                    '        "Received : " & Item.ReceivedTime & vbCrLf & _
                    '        "Subject : " & Item.Subject & vbCrLf & _
                    '        "Size : " & Item.Size & vbCrLf & _
                     '       "Message Body : " & vbCrLf & Item.Body
                     '   Result = MsgBox(MessageInfo, vbOKOnly, "New Message Received")
               ' End If
            End If
        Next i
        Set olFolder = Nothing
    Next eFolder
End Sub

Sub filter(Item As Outlook.MailItem)
    Dim ns As Outlook.Namespace
    Dim MailDest As Outlook.Folder
    Set ns = Application.GetNamespace("MAPI")
    Set Reg1 = CreateObject("VBScript.RegExp")
    Reg1.Global = True
    Reg1.Pattern = "(.*Test download.*)"
    If Reg1.test(Item.Subject) Then
        'Set MailDest = ns.Folders("Personal Folders").Folders("one").Folders("a")
        'Item.Move MailDest
        MsgBox Item.Body
    End If
End Sub


Sub TextTest()
Dim text As Range
text = Selection
  text.TextToColumns Destination:=text.Cells(1, -5), DataType:=xlDelimited _
            , TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
   Dim dTime As Date
Dim x As Integer

x = 1

For dTime = "3/01/2013 12:00:00 AM" To "3/02/2013 11:55:00 PM" Step "00:05"
    ' Sets the date in the cell of the first column
    Cells(x, 1).Value = Format(dTime, "mm/dd/yyyy")

    ' Sets the time in the cell of the second column
    Cells(x, 2).Value = Format(dTime, "hh:mm:ss d")
     
    x = x + 1
Next dTime
End Sub
 
Sub Macro2()
  
Dim dTime As Date
Dim x As Integer
Dim ws As worksheet
Set ws = ActiveSheet
Dim index As Integer
  
MsgBox ws.Rows.Count

x = 1

For index = 1 To ws.UsedRange.Rows.Count
    ' Sets the date in the cell of the first column
    Cells(x, 1).Value = Format(Cells(index, 3), "mm/dd/yyyy")

    ' Sets the time in the cell of the second column
    Cells(x, 2).Value = Format(Cells(index, 3), "hh:mm:ss t")
     
    x = x + 1
Next
End Sub


Sub Test8()
Dim a, b
Dim workbook As workbook
Dim sourceSheet As worksheet
Dim worksheet As worksheet

  Set sourceSheet = Worksheets("Sheet4")
  Set workbook = Application.Workbooks.Open("D:\test.xlsx")
  Set worksheet = workbook.Worksheets("Sheet2")
  a = sourceSheet.Cells(Rows.Count, 1).End(xlUp).row
  
    For i = 2 To a
        If sourceSheet.Cells(i, 3).Value = "KSR" Then
            'sourceSheet.Rows(i).Copy
             sourceSheet.Cells(i, 3).Copy
             
            worksheet.Activate

            b = worksheet.Cells(Rows.Count, 1).End(xlUp).row
            
            worksheet.Cells(b + 1, 1).Select
            
            MsgBox "A" & (b + 1)
            
            worksheet.Paste
             
            sourceSheet.Activate
        End If
    Next

    Application.CutCopyMode = False

    ThisWorkbook.Worksheets("Sheet4").Cells(1, 1).Select

End Sub
