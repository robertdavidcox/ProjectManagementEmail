Option Explicit

Sub SendEmails()

Dim d As Range
Dim action As String
Dim property As String

Dim i As Integer
Dim j As Integer
Dim olApp As Object
Dim objMail As Object

Set olApp = CreateObject("Outlook.application")
i = 0

Call deleteOutbox

For Each d In Range("dates")

    For j = 0 To Range("num_properties")

        action = Range("first_date").Offset(i, j + 1).Value
        property = Range("first_property").Offset(0, j)
        
        If Not (action = "") And Range("today") <= d Then

            Set objMail = olApp.CreateItem(0)
            With objMail
                .Display
                .To = Range("email")
                .Subject = property & " - " & d & " - " & action
                .HTMLBody = ""
                .DeferredDeliveryTime = Range("first_date").Offset(i, 0)
                .Send
            End With
            
        End If
        
    Next j
    
    i = i + 1
    
Next

End Sub


Sub deleteOutbox()
Application.DisplayAlerts = False
Application.Calculate
Dim datecheck As Date
Dim emailBody As String
Dim olApp As Object
Dim myNameSpace As Object
Dim myFolder As Object
Dim myItem As Object

Set olApp = CreateObject("Outlook.Application")
Set myNameSpace = olApp.GetNamespace("MAPI")
Set myFolder = myNameSpace.GetDefaultFolder(4)

For Each myItem In myFolder.Items
    myItem.Delete

Next

End Sub