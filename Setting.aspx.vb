Imports System.Data
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail

Partial Class Setting
    Inherits Page

    Dim settingClass As New SettingClass
    Dim myConn As String = ConfigurationManager.ConnectionStrings("DefaultConnection").ConnectionString

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If String.IsNullOrEmpty(Request.QueryString("action")) Then
            Exit Sub
        End If

        Dim thisAction As String = Request.QueryString("action").ToString()
        If thisAction = "createsales" Then
            CreateSales()
        End If
        If thisAction = "refreshsales" Then
            RefreshSales()
        End If
        If thisAction = "downloadboe" Then
            UpdateDownloadBOE()
        End If
        If thisAction = "deleteorderactioncontext" Then
            DeleteOrderActionContext()
        End If
        If thisAction = "deletenullsession" Then
            DeleteNullSession()
        End If
        If thisAction = "resetproformaorder" Then
            ResetProformaOrder()
        End If
        If thisAction = "shipment" Then
            If String.IsNullOrEmpty(Request.QueryString("OrdID")) Then
                Exit Sub
            End If
            If String.IsNullOrEmpty(Request.QueryString("status")) Then
                Exit Sub
            End If
            Dim id As String = Request.QueryString("OrdID").ToString()
            Dim status As String = Request.QueryString("Status").ToString()
            Dim shipmentNumber As String = Request.QueryString("ShipmentNo").ToString()
            Dim containerNumber As String = Request.QueryString("ContainerNo").ToString()
            Dim courier As String = Request.QueryString("Courier").ToString()
            Dim invoiceNumber As String = Request.QueryString("InvoiceNo").ToString()

            Dim shipDateStr As String = Request.QueryString("ShipDate")
            Dim shipDate As DateTime

            If String.IsNullOrEmpty(shipDateStr) OrElse Not DateTime.TryParse(shipDateStr, shipDate) Then
                Exit Sub
            End If

            UpdateShipment(id, status, shipmentNumber, shipDate, containerNumber, courier, invoiceNumber)
        End If
    End Sub

    Protected Sub CreateSales()
        Try
            Dim companyData As DataTable = settingClass.GetDataTable("SELECT Id FROM Companys WHERE Status='Active' ORDER BY Id ASC")
            If companyData.Rows.Count > 0 Then
                For i As Integer = 0 To companyData.Rows.Count - 1
                    Dim companyId As String = companyData.Rows(i)("Id").ToString()

                    Dim salesData As Integer = settingClass.GetItemData_Integer("SELECT COUNT(*) FROM Sales WHERE SummaryDate=GETDATE() AND CompanyId='" & companyId & "'")
                    If salesData = 0 Then
                        Dim thisId As String = settingClass.CreateId("SELECT TOP 1 Id FROM Sales ORDER BY Id DESC")

                        Using thisConn As New SqlConnection(myConn)
                            Using thisCmd As SqlCommand = New SqlCommand("INSERT INTO Sales(Id, CompanyId, SummaryDate, TotalCostPrice, TotalSellingPrice, TotalPaidAmount) VALUES(@Id, @CompanyId, GETDATE(), 0, 0, 0)", thisConn)
                                thisCmd.Parameters.AddWithValue("@Id", thisId)
                                thisCmd.Parameters.AddWithValue("@CompanyId", companyId)
                                thisConn.Open()
                                thisCmd.ExecuteNonQuery()
                            End Using
                        End Using
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub RefreshSales()
        Try
            Dim dataCompany As DataTable = settingClass.GetDataTable("SELECT Id FROM Companys WHERE Status='Active' OR Status='Inactive'")
            If dataCompany.Rows.Count > 0 Then
                For i As Integer = 0 To dataCompany.Rows.Count - 1
                    Dim companyId As String = dataCompany.Rows(i)("Id").ToString()
                    settingClass.RefreshSalesData(companyId)
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub UpdateFactory()
        Try
            Dim orderData As DataTable = settingClass.GetDataTable("SELECT Id FROM OrderHeaders WHERE Active=1 ORDER BY Id ASC")
            If orderData.Rows.Count > 0 Then
                For i As Integer = 0 To orderData.Rows.Count - 1
                    Dim headerId As String = orderData.Rows(i)("Id").ToString()
                    settingClass.UpdateOrderFactory(headerId)
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub UpdateDownloadBOE()
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand("sp_OrderHeaders_Update_BOE", thisConn)
                    thisCmd.CommandType = CommandType.StoredProcedure
                    thisConn.Open()
                    thisCmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub DeleteOrderActionContext()
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As SqlCommand = New SqlCommand("DELETE FROM OrderActionContext", thisConn)
                    thisConn.Open()
                    thisCmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub DeleteNullSession()
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As SqlCommand = New SqlCommand("DELETE FROM Sessions WHERE LoginId IS NULL", thisConn)
                    thisConn.Open()
                    thisCmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ResetProformaOrder()
        Try
            Dim thisData As DataTable = settingClass.GetDataTable("SELECT Id FROM OrderHeaders WHERE Status='Proforma Sent' AND DueDate=CAST(GETDATE() AS DATE)")
            If thisData.Rows.Count > 0 Then
                For i As Integer = 0 To thisData.Rows.Count - 1
                    Dim thisId As String = thisData.Rows(i)("Id").ToString()

                    Using thisConn As New SqlConnection(myConn)
                        Using thisCmd As SqlCommand = New SqlCommand("UPDATE OrderHeaders SET Status='Pending Payment' WHERE Id=@Id", thisConn)
                            thisCmd.Parameters.AddWithValue("@Id", thisId)
                            thisConn.Open()
                            thisCmd.ExecuteNonQuery()
                        End Using
                    End Using

                    Dim dataLog As Object() = {"OrderHeaders", thisId, 2, "Pending Payment Order"}
                    settingClass.Logs(dataLog)

                    ResetProformaOrder(thisId)
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ResetProformaOrder(headerId As String)
        Try
            If String.IsNullOrEmpty(headerId) Then Exit Sub

            Dim orderData As DataRow = settingClass.GetDataRow("SELECT OrderHeaders.*, Customers.Name AS CustomerName, Customers.CompanyId AS CompanyId, Customers.Operator AS Operator, OrderHeaders.InvoiceNumber AS InvoiceNumber FROM OrderHeaders LEFT JOIN Customers ON OrderHeaders.CustomerId=Customers.Id WHERE OrderHeaders.Id='" & headerId & "'")
            If orderData Is Nothing Then Exit Sub

            Dim customerId As String = orderData("CustomerId").ToString()
            Dim orderId As String = orderData("OrderId").ToString()
            Dim orderNumber As String = orderData("OrderNumber").ToString()
            Dim orderName As String = orderData("OrderName").ToString()
            Dim invoiceNumber As String = orderData("InvoiceNumber").ToString()

            Dim customerName As String = orderData("CustomerName").ToString()

            Dim companyId As String = orderData("CompanyId").ToString()
            Dim companyName As String = settingClass.GetItemData("SELECT Name FROM Companys WHERE Id='" & companyId & "'")
            If companyId = "3" Then companyName = "PT Bumi Indah Global"

            Dim mailData As DataRow = settingClass.GetDataRow("SELECT * FROM Mailings WHERE CompanyId='" & companyId & "' AND Name='Reset Proforma Order' AND Active=1")
            If mailData Is Nothing Then Exit Sub

            Dim mailServer As String = mailData("Server").ToString()
            Dim mailHost As String = mailData("Host").ToString()
            Dim mailPort As Integer = mailData("Port")

            Dim mailAccount As String = mailData("Account").ToString()
            Dim mailPassword As String = mailData("Password").ToString()
            Dim mailAlias As String = mailData("Alias").ToString()
            Dim mailSubject As String = mailData("Subject").ToString()

            Dim mailTo As String = mailData("To").ToString()
            Dim mailCc As String = mailData("Cc").ToString()
            Dim mailBcc As String = mailData("Bcc").ToString()

            Dim mailNetworkCredentials As Boolean = mailData("NetworkCredentials")
            Dim mailDefaultCredentials As Boolean = mailData("DefaultCredentials")
            Dim mailEnableSSL As Boolean = mailData("EnableSSL")

            Dim customerMail As DataTable = settingClass.GetDataTable("SELECT Email FROM CustomerContacts CROSS APPLY STRING_SPLIT(Tags, ',') AS thisArray WHERE CustomerId='" & customerId & "' AND thisArray.VALUE='Confirming'")
            If customerMail.Rows.Count = 0 Then Exit Sub

            Dim mailBody As String = String.Empty

            mailBody = "<span style='font-family: Cambria; font-size: 16px;'>"
            mailBody &= "<i>- THIS IS AN AUTOMATED EMAIL. KINDLY DO NOT REPLY WITHOUT COPYING OUR TEAM. -</i>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "Dear Valued Customer,"
            mailBody &= "<br /><br />"
            mailBody &= "This order has been moved to <b>Pending Payment</b>.</u></b>."
            mailBody &= "<br /><br />"
            mailBody &= "Please note that if payment is received after a price change has taken effect, the order will be processed using our <b>latest pricing</b>."
            mailBody &= "<br /><br />"
            mailBody &= "If our pricing changes before your payment is received, the order status will be <b>automatically updated to Unsubmitted</b>."
            mailBody &= "<br />"
            mailBody &= "The order will be repriced according to our <b>current price list</b>, and you will be <b>required to review and resubmit the order</b> before processing can continue."
            mailBody &= "<br /><br />"
            mailBody &= "Thank you for your understanding."
            mailBody &= "</span>"

            mailBody &= "<br /><br /><br />"

            mailBody &= "<span style='font-family: Cambria; font-size:16px;'>Kind Regards,</span>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "<span style='font-family: Cambria; font-size:16px; font-weight: bold;'>" & companyName.ToUpper() & "</span>"

            Dim myMail As New MailMessage()

            Dim subject As String = String.Format("{0} - {1} - {2} - Due Date Order # {3}", customerName, orderNumber, orderName, orderId)

            myMail.Subject = subject
            myMail.From = New MailAddress(mailServer, mailAlias)

            If customerMail.Rows.Count > 0 Then
                For i As Integer = 0 To customerMail.Rows.Count - 1
                    Dim thisEmail As String = customerMail.Rows(i)("Email").ToString()
                    myMail.To.Add(thisEmail)
                Next
            End If

            If Not String.IsNullOrEmpty(mailCc) Then
                For Each thisMail In mailCc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.CC.Add(thisMail.Trim())
                Next
            End If

            If companyId = "2" Then
                Dim operatorEmail As String = settingClass.GetItemData("SELECT ISNULL(STRING_AGG(Logins.Email, ';'), '') FROM Customers OUTER APPLY STRING_SPLIT(Customers.Operator, ',') operatorArray LEFT JOIN Logins ON Logins.Id = TRY_CAST(operatorArray.value AS INT) WHERE Customers.Id='" & customerId & "';")

                If Not String.IsNullOrEmpty(operatorEmail) Then
                    Dim emailList() As String = operatorEmail.Split(";"c)

                    For Each email As String In emailList
                        If Not String.IsNullOrWhiteSpace(email) Then
                            myMail.CC.Add(email.Trim())
                        End If
                    Next
                End If
            End If

            If Not String.IsNullOrEmpty(mailBcc) Then
                For Each thisMail In mailBcc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.Bcc.Add(thisMail.Trim())
                Next
            End If

            myMail.IsBodyHtml = True
            myMail.Body = mailBody
            Dim smtpClient As New SmtpClient()
            smtpClient.Host = mailHost
            smtpClient.Port = mailPort
            smtpClient.EnableSsl = mailEnableSSL
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
            smtpClient.Timeout = 120000

            If mailNetworkCredentials Then
                smtpClient.UseDefaultCredentials = False
                smtpClient.Credentials = New NetworkCredential(mailAccount, mailPassword)
            Else
                smtpClient.UseDefaultCredentials = mailDefaultCredentials
            End If

            smtpClient.Send(myMail)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub UpdateShipment(id As String, status As String, shipNumber As String, shipDate As Date, conNumber As String, courier As String, invNumber As String)
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As SqlCommand = New SqlCommand("UPDATE OrderHeaders SET Status=@Status, ShipmentNumber=@ShipmentNumber, ShipmentDate=@ShipmentDate, ContainerNumber=@ContainerNumber, Courier=@Courier, InvoiceNumber=@InvoiceNumber WHERE Id=@Id", thisConn)
                    thisCmd.Parameters.AddWithValue("@Id", id)
                    thisCmd.Parameters.AddWithValue("@ShipmentNumber", shipNumber)
                    thisCmd.Parameters.AddWithValue("@ShipmentDate", shipDate)
                    thisCmd.Parameters.AddWithValue("@ContainerNumber", conNumber)
                    thisCmd.Parameters.AddWithValue("@Courier", courier)
                    thisCmd.Parameters.AddWithValue("@Status", status)
                    thisCmd.Parameters.AddWithValue("@InvoiceNumber", invNumber)
                    thisConn.Open()
                    thisCmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
