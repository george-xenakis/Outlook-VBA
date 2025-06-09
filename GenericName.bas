Attribute VB_Name = "GenericName"
Public Sub Asattachmentname(Item As Outlook.MailItem)
Dim objAtt As Outlook.Attachment
Dim saveFolder As String
saveFolder = "\\cru-file-01\Reports\Analyst_Test\Test"
     For Each objAtt In Item.Attachments
            timerec = Format(Item.ReceivedTime, "yyyy-mm-dd-hh-mm-ss")
          objAtt.SaveAsFile saveFolder & "/" & timerec & objAtt.DisplayName
          Set objAtt = Nothing
     Next
End Sub
