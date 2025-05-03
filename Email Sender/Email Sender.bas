Attribute VB_Name = "Module1"
Sub Email_Sender()

    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim currentDate As Date
    Dim bodyText As String
    Dim tableRows As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fixedExtension As String
    Dim shiftType As String
    Dim subjectLine As String
    Dim toRecipients As String
    Dim ccRecipients As String
    Dim rowColor As String
    Dim extValue As String

    ' Set fixed extension
    fixedExtension = "0000"
    
    ' Set current date
    currentDate = Date
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Main")
    
    ' Get shift type from cell D4
    shiftType = ws.Cells(4, 4).Value
    
    ' Build subject line
    subjectLine = "ABC operations team contacts || " & Format(currentDate, "d/m/yyyy") & " " & shiftType & " Shift"
    
    ' Build table rows
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    tableRows = ""
    
    For i = 4 To lastRow
        If ws.Cells(i, 2).Value <> "" And ws.Cells(i, 3).Value <> "" Then
            rowColor = IIf(i Mod 2 = 0, "#FFFFFF", "#F2F2F2")
            extValue = IIf(i = 4, fixedExtension, "") ' Show extension only in the first row

            tableRows = tableRows & "<tr style='background-color:" & rowColor & ";'>" & _
                "<td style='border:1px solid #BFBFBF; padding:6px; width:2.45in; text-align:center; vertical-align:middle;'>" & ws.Cells(i, 2).Value & "</td>" & _
                "<td style='border:1px solid #BFBFBF; padding:6px; width:2.45in; text-align:center; vertical-align:middle;'>" & ws.Cells(i, 3).Value & "</td>" & _
                "<td style='border:1px solid #BFBFBF; padding:6px; width:1.45in; text-align:center; vertical-align:middle;'>" & extValue & "</td>" & _
                "</tr>"
        End If
    Next i
    
    ' Collect TO and CC recipients
    toRecipients = ""
    ccRecipients = ""
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    For i = 4 To lastRow
        If ws.Cells(i, 5).Value <> "" Then
            toRecipients = toRecipients & ws.Cells(i, 5).Value & ";"
        End If
        If ws.Cells(i, 6).Value <> "" Then
            ccRecipients = ccRecipients & ws.Cells(i, 6).Value & ";"
        End If
    Next i

    ' Remove trailing semicolons
    If Right(toRecipients, 1) = ";" Then
        toRecipients = Left(toRecipients, Len(toRecipients) - 1)
    End If
    If Right(ccRecipients, 1) = ";" Then
        ccRecipients = Left(ccRecipients, Len(ccRecipients) - 1)
    End If
    
    ' Build HTML body
    bodyText = "<html><body style='font-family:Calibri, sans-serif; font-size:11pt; color:#000;'>" & _
        "<p style='color:#2F5597;'>Dear All,<br>&nbsp;&nbsp;&nbsp;&nbsp;Kindly find the <span style='color:#C00000; font-weight:bold;'>ABC operations Team</span> contact number below:</p>" & _
        "<div style='display:flex; justify-content:center;'>" & _
        "<table style='border-collapse:collapse;'>" & _
        "<tr style='background-color:#D9E1F2; font-weight:bold;'>" & _
        "<th style='border:1px solid #BFBFBF; padding:6px; width:2.45in; text-align:center; vertical-align:middle;'>Name</th>" & _
        "<th style='border:1px solid #BFBFBF; padding:6px; width:2.45in; text-align:center; vertical-align:middle;'>Mobile number</th>" & _
        "<th style='border:1px solid #BFBFBF; padding:6px; width:1.45in; text-align:center; vertical-align:middle;'>Extension</th>" & _
        "</tr>" & tableRows & "</table></div>" & _
        "<p>Best Regards,<br>ABC Opertions Team</p>" & _
        "</body></html>"

    ' Create and display the email
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)

    With outlookMail
        .To = toRecipients
        .CC = ccRecipients
        .Subject = subjectLine
        .HTMLBody = bodyText
        .Display ' Use .Send to send automatically
    End With

End Sub


