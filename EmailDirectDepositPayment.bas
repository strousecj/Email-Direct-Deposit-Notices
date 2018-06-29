'A function to create and format direct deposit email notices from Excel
'and email them out to vendors.
Sub emailDDPayment()
    Dim OutApp As Object
    Dim OutMail As Object, tempOutMail As Object
    Dim cell As Range ', tempCell As String
    Dim htmlString As String
    Dim needToDisplay As Boolean
    Dim depositAmount As String
    Dim depositTotal As Currency
    Dim companyName As String

    Application.ScreenUpdating = False
    needToDisplay = False
    Set OutApp = CreateObject("Outlook.Application")
    depositTotal = 0

    If IsEmpty(Range("I2").Value) = True Then
        MsgBox "Don't forget to list list which company is doing the depositing!"
    
    Else
        On Error GoTo cleanup
        For Each cell In Columns("G").Cells.SpecialCells(xlCellTypeConstants)
            If cell.Value Like "?*@?*.?*" Then
                depositAmount = FormatCurrency(Cells(cell.Row, "D").Value, 2)
                depositTotal = depositTotal + depositAmount
                companyName = Cells(2, "I")
                Set OutMail = OutApp.CreateItem(0)
                    Set tempOutMail = OutApp.CreateItem(0)
                    tempOutMail.HTMLBody = "<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""background-color:#FFFFFF;border-style:hidden;"">" & _
                                "<tbody>" & _
                                    "<tr> <td>Vendor #</td> <td>Vendor Name</td> <td>Payment Date</td>" & _
                                    "<td>Invoice #</td> <td>Deposit Amount</td> <td>Bank Acct #</td> </tr>" & vbNewLine & _
                                    "<tr> <td>" & Cells(cell.Row, "A").Value & "</td> <td>" & Cells(cell.Row, "C") & _
                                    "</td> <td>" & Cells(cell.Row, "B").Value & "</td> <td>" & Cells(cell.Row, "H").Value & _
                                    "</td> <td>" & depositAmount & "</td> <td>" & Cells(cell.Row, "F").Value & "</td> </tr>" & _
                                "</tbody>" & _
                            "</table>"
                On Error Resume Next
                If cell.Value = Cells(cell.Row + 1, 7).Value Then
                    htmlString = htmlString & vbNewLine & tempOutMail.HTMLBody
                    With tempOutMail
                        .To = cell.Value
                        .Subject = "Direct Deposit Payment"
                        .HTMLBody = "<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""background-color:#FFFFFF;border-style:hidden;"">" & _
                                "<tbody>" & _
                                    "<tr> <td>Vendor #</td> <td>Vendor Name</td> <td>Payment Date</td>" & _
                                    "<td>Invoice #</td> <td>Deposit Amount</td> <td>Bank Acct #</td> </tr>" & vbNewLine & _
                                    "<tr> <td>" & Cells(cell.Row, "A").Value & "</td> <td>" & Cells(cell.Row, "C") & _
                                    "</td> <td>" & Cells(cell.Row, "B").Value & "</td> <td>" & Cells(cell.Row, "H").Value & _
                                    "</td> <td>" & depositAmount & "</td> <td>" & Cells(cell.Row, "F").Value & "</td> </tr>" & _
                                "</tbody>" & _
                            "</table>"
                    End With
                    On Error GoTo 0
                    Set OutMail = Nothing
                    needToDisplay = True
                Else
                    If needToDisplay = False Then
                        With OutMail
                            .To = cell.Value
                            .Subject = "Direct Deposit Payment"
                            'Edit the email message below here! ******************************************************************************************
                            .HTMLBody = "Hello," & "<br> <br>" & _
                                        "Below you will find the details on the payment made to you this week by " & companyName & ". " & _
                                        "This will be deposited into your bank account on Friday, 06/29/18." & "<br> <br>" & _
                                        "<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""background-color:#FFFFFF;border-style:hidden;"">" & _
                                            "<tbody>" & _
                                                "<tr> <td>Vendor #</td> <td>Vendor Name</td> <td>Payment Date</td>" & _
                                                "<td>Invoice #</td> <td>Deposit Amount</td> <td>Bank Acct #</td> </tr>" & vbNewLine & _
                                                "<tr> <td>" & Cells(cell.Row, "A").Value & "</td> <td>" & Cells(cell.Row, "C") & _
                                                "</td> <td>" & Cells(cell.Row, "B").Value & "</td> <td>" & Cells(cell.Row, "H").Value & _
                                                "</td> <td>" & depositAmount & "</td> <td>" & Cells(cell.Row, "F").Value & "</td> </tr>" & _
                                            "</tbody>" & _
                                        "</table>" & "<br>" & "Deposit Total: " & FormatCurrency(depositTotal)
                            .Display 'Or use Display to show email before sending
                        End With
                        depositTotal = 0
                    Else
                        htmlString = htmlString & vbNewLine & tempOutMail.HTMLBody
                        With OutMail
                            .To = cell.Value
                            .Subject = "Direct Deposit Payment"
                            'Edit the email message below here! ******************************************************************************************
                            .HTMLBody = "Hello," & "<br>" & "<br>" & _
                                    "Below you will find the details on the payment made to you this week by " & companyName & ". " & _
                                    "This will be deposited into your bank account on Friday, 06/29/18." & "<br> <br>" & _
                                    htmlString & "<br>" & "Deposit Total: " & FormatCurrency(depositTotal)
                            .Display
                        End With
                        needToDisplay = False
                        htmlString = ""
                        depositTotal = 0
                    End If
                End If
            End If
            'tempCell = cell.Value
        Next cell
    End If
    
cleanup:
    Set OutApp = Nothing
    Application.ScreenUpdating = True
End Sub
