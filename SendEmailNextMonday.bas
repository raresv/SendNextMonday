Sub SendEmailNextMonday()

 Dim myinspector As Outlook.Inspector
 Dim myItem As Outlook.MailItem
 
 Dim mondayDate As Date
 Dim todayDate As Date
 
 todayDate = DateValue(Now())

 mondayDate = DateAdd("d", 8 - (Weekday(todayDate, vbMonday)), todayDate)
 mondayDate = DateAdd("h", 8, mondayDate)
 
 Set myinspector = Application.ActiveInspector

 Set myItem = myinspector.CurrentItem
 
 myItem.DeferredDeliveryTime = mondayDate
  
 myItem.Send

End Sub
