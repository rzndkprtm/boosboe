Imports System.Data
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail

Public Class SettingClass
    Dim myConn As String = ConfigurationManager.ConnectionStrings("DefaultConnection").ConnectionString

    Public Function GetDataRow(thisString As String) As DataRow
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        Dim thisTable As New DataTable()
                        thisAdapter.Fill(thisTable)
                        If thisTable.Rows.Count > 0 Then
                            Return thisTable.Rows(0)
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetDataRowSP(spName As String, params As List(Of SqlParameter)) As DataRow
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(spName, thisConn)
                    thisCmd.CommandType = CommandType.StoredProcedure
                    thisCmd.Parameters.AddRange(params.ToArray())
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        Dim thisTable As New DataTable()
                        thisAdapter.Fill(thisTable)
                        If thisTable.Rows.Count > 0 Then
                            Return thisTable.Rows(0)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

    Public Function GetDataTable(thisString As String) As DataTable
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        Dim thisTable As New DataTable()
                        thisAdapter.Fill(thisTable)
                        Return thisTable
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetDataTableSP(spName As String, params As List(Of SqlParameter)) As DataTable
        Dim thisTable As New DataTable()
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(spName, thisConn)
                    thisCmd.CommandType = CommandType.StoredProcedure
                    If params IsNot Nothing Then
                        If params.Count > 0 Then
                            thisCmd.Parameters.AddRange(params.ToArray())
                        End If
                    End If
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        thisAdapter.Fill(thisTable)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            thisTable = New DataTable()
        End Try
        Return thisTable
    End Function

    Public Function GetItemData(thisString As String) As String
        Dim result As String = String.Empty
        Try
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult = thisCmd.ExecuteReader
                        While rdResult.Read
                            result = rdResult.Item(0).ToString()
                        End While
                    End Using
                End Using
                thisConn.Close()
            End Using
        Catch ex As Exception
            result = String.Empty
        End Try
        Return result
    End Function

    Public Function GetItemData_Integer(thisString As String) As Integer
        Dim result As Integer = 0
        Try
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult = thisCmd.ExecuteReader
                        While rdResult.Read
                            result = rdResult.Item(0)
                        End While
                    End Using
                End Using
                thisConn.Close()
            End Using
        Catch ex As Exception
            result = 0
        End Try
        Return result
    End Function

    Public Function GetItemData_Decimal(thisString As String) As Decimal
        Dim result As Double = 0D
        Try
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult = thisCmd.ExecuteReader
                        While rdResult.Read
                            result = rdResult.Item(0)
                        End While
                    End Using
                End Using
                thisConn.Close()
            End Using
        Catch ex As Exception
            result = 0D
        End Try
        Return result
    End Function

    Public Function GetItemData_Boolean(thisString As String) As Boolean
        Dim result As Boolean = False
        Try
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult = thisCmd.ExecuteReader
                        While rdResult.Read
                            result = rdResult.Item(0)
                        End While
                    End Using
                End Using
                thisConn.Close()
            End Using
        Catch ex As Exception
            result = False
        End Try
        Return result
    End Function

    Public Function CreateId(thisString As String) As String
        Dim result As String = String.Empty
        Try
            Dim id As Integer = 0
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult As SqlDataReader = thisCmd.ExecuteReader()
                        If rdResult.Read() Then
                            Integer.TryParse(rdResult(0).ToString(), id)
                        End If
                    End Using
                End Using
            End Using
            result = (id + 1).ToString()
        Catch ex As Exception
            result = String.Empty
        End Try
        Return result
    End Function

    Public Sub Logs(data As Object())
        Try
            If data.Length = 4 Then
                Dim type As String = Convert.ToString(data(0))
                Dim dataId As String = Convert.ToString(data(1))
                Dim loginId As String = Convert.ToString(data(2))
                Dim description As String = Convert.ToString(data(3))

                Using thisConn As SqlConnection = New SqlConnection(myConn)
                    Using thisCmd As SqlCommand = New SqlCommand("INSERT INTO Logs VALUES (NEWID(), @Type, @DataId, @ActionBy, GETDATE(), @Description)", thisConn)
                        thisCmd.Parameters.AddWithValue("@Type", type)
                        thisCmd.Parameters.AddWithValue("@DataId", If(String.IsNullOrEmpty(dataId), CType(DBNull.Value, Object), dataId))
                        thisCmd.Parameters.AddWithValue("@ActionBy", loginId)
                        thisCmd.Parameters.AddWithValue("@Description", description)
                        thisConn.Open()
                        thisCmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub RefreshSalesData(companyId As String)
        Try
            If Not String.IsNullOrEmpty(companyId) Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand("sp_Sales_Refresh", thisConn)
                        thisCmd.CommandType = CommandType.StoredProcedure
                        thisCmd.Parameters.AddWithValue("@CompanyId", companyId)
                        thisConn.Open()
                        thisCmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub UpdateOrderFactory(headerId As String)
        Try
            Try
                Dim factoryList As New List(Of String)

                Dim detailData As DataTable = GetDataTable("SELECT OrderDetails.*, Products.Name AS ProductName, Designs.Name AS DesignName, Blinds.Name AS BlindName FROM OrderDetails LEFT JOIN Products ON OrderDetails.ProductId=Products.Id LEFT JOIN Designs ON Products.DesignId=Designs.Id LEFT JOIN Blinds ON Products.DesignId=Blinds.Id WHERE OrderDetails.HeaderId='" & headerId & "' AND OrderDetails.Active=1")
                For iDetail As Integer = 0 To detailData.Rows.Count - 1
                    Dim designName As String = detailData.Rows(iDetail)("DesignName").ToString()
                    Dim blindName As String = detailData.Rows(iDetail)("BlindName").ToString()

                    Dim isBig As Boolean = False
                    Dim isAustralia As Boolean = False
                    Dim isTaiwan As Boolean = False
                    Dim isChina As Boolean = False

                    If designName = "Aluminium Blind" OrElse designName = "Design Shades" OrElse designName = "Linea Valance" OrElse designName = "Panel Glide" OrElse designName = "Pelmet" OrElse designName = "Roman Blind" OrElse designName = "Soft Roman" OrElse designName = "Privacy Venetian" OrElse designName = "Venetian Blind" OrElse designName = "Vertical" OrElse designName = "Roller Blind" OrElse designName = "Sample" OrElse designName = "Skyline Shutter Express" OrElse designName = "Outdoor" OrElse designName = "Saphora Drape" OrElse designName = "Roller Horizon" OrElse designName = "Venetian Part" Then
                        isBig = True
                    End If

                    If designName = "Cellular Shades" Then
                        Dim fabricColourId As String = detailData.Rows(iDetail)("FabricColourId").ToString()
                        Dim fabricFactory As String = GetItemData("SELECT Factory FROM FabricColours WHERE Id='" & fabricColourId & "'")

                        If fabricFactory = "Express" Then isBig = True
                        If fabricFactory = "Regular" Then isTaiwan = True
                    End If

                    If designName = "Curtain" Then
                        Dim fabricColourId As String = detailData.Rows(iDetail)("FabricColourId").ToString()
                        Dim fabricColourIdB As String = detailData.Rows(iDetail)("FabricColourIdB").ToString()

                        Dim trackType As String = detailData.Rows(iDetail)("TrackType").ToString()
                        Dim trackTypeB As String = detailData.Rows(iDetail)("TrackTypeB").ToString()

                        Dim fabricFactory As String = String.Empty
                        Dim fabricFactoryB As String = String.Empty

                        If fabricColourId <> "" Then
                            fabricFactory = GetItemData("SELECT Factory FROM FabricColours WHERE Id='" & fabricColourId & "'")
                        End If

                        If fabricColourIdB <> "" Then
                            fabricFactoryB = GetItemData("SELECT Factory FROM FabricColours WHERE Id='" & fabricColourIdB & "'")
                        End If

                        If fabricFactory = "Express" Then
                            isBig = True
                        ElseIf fabricFactory = "Regular" Then
                            isAustralia = True
                        End If

                        If fabricFactoryB = "Express" Then
                            isBig = True
                        ElseIf fabricFactoryB = "Regular" Then
                            isAustralia = True
                        End If

                        If trackType = "Standard Track" Then
                            isAustralia = True
                        ElseIf trackType = "Motorised Track" Then
                            isBig = True
                        End If

                        If trackTypeB = "Standard Track" Then
                            isAustralia = True
                        ElseIf trackTypeB = "Motorised Track" Then
                            isBig = True
                        End If
                    End If

                    If designName = "Skyline Shutter Ocean" OrElse designName = "Evolve Shutter Ocean" Then
                        isChina = True
                    End If

                    If designName = "Door" OrElse designName = "Window" Then
                        Dim frameColour As String = detailData.Rows(iDetail)("FrameColour").ToString()

                        If frameColour.Contains("Regular") Then isAustralia = True
                        If frameColour.Contains("Express") Then isBig = True
                    End If

                    If isBig AndAlso Not factoryList.Contains("BIG") Then
                        factoryList.Add("BIG")
                    End If

                    If isTaiwan AndAlso Not factoryList.Contains("TAIWAN") Then
                        factoryList.Add("TAIWAN")
                    End If

                    If isAustralia AndAlso Not factoryList.Contains("AUS") Then
                        factoryList.Add("AUS")
                    End If

                    If isChina AndAlso Not factoryList.Contains("CHINA") Then
                        factoryList.Add("CHINA")
                    End If
                Next

                Dim factory As String = String.Join(", ", factoryList)

                Using thisConn As SqlConnection = New SqlConnection(myConn)
                    Using thisCmd As SqlCommand = New SqlCommand("UPDATE OrderHeaders SET OrderFactory=@OrderFactory WHERE Id=@Id", thisConn)
                        thisCmd.Parameters.AddWithValue("@Id", headerId)
                        thisCmd.Parameters.AddWithValue("@OrderFactory", factory)
                        thisConn.Open()
                        thisCmd.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
            End Try
        Catch ex As Exception

        End Try
    End Sub


    ' MAILING
    Public Sub JournalSummary(salesId As String)
        Try
            If String.IsNullOrEmpty(salesId) Then Exit Sub

            Dim loginData As DataRow = GetDataRow("SELECT * FROM Logins WHERE Id='" & salesId & "'")
            If loginData Is Nothing Then Exit Sub

            Dim mailData As DataRow = GetDataRow("SELECT * FROM Mailings WHERE CompanyId='2' AND Name='Journal Summary' AND Active=1")
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

            Dim params As New List(Of SqlParameter) From {
            New SqlParameter("@LoginId", salesId)
        }
            Dim journalData As DataTable = GetDataTableSP("sp_Journals_Reminder", params)
            If journalData.Rows.Count = 0 Then Exit Sub

            Dim salesName As String = GetItemData("SELECT FullName FROM Logins WHERE Id='" & salesId & "'")
            Dim salesMail As String = GetItemData("SELECT Email FROM Logins WHERE Id='" & salesId & "'")
            If String.IsNullOrEmpty(salesMail) Then Exit Sub

            Dim mailBody As String = String.Empty

            Dim today As Date = Date.Today
            Dim startDate As Date = today.AddDays(-(CInt(today.DayOfWeek) + 6) Mod 7)
            Dim endDate As Date = startDate.AddDays(4)

            Dim period As String = startDate.ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture) & " - " & endDate.ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture)

            mailBody &= "<div style='font-family:Cambria; font-size:16px;'>"
            mailBody &= "<p>Hi <b>" & salesName & "</b>,</p>"
            mailBody &= "<p>Please find below a summary of the journals recorded this week for the period <b>" & period & "</b>.</p>"
            mailBody &= "</div>"
            mailBody &= "<br />"

            mailBody &= "<table style='border-collapse:collapse; width:60%; font-family:Cambria; font-size:16px;'>"
            mailBody &= "<tr style='background-color:#00B0F0; color:black; font-weight:bold; text-align:center;'>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>No</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Journal Date</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Customer</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Total Customer</th>"
            mailBody &= "</tr>"

            Dim no As Integer = 1
            For i As Integer = 0 To journalData.Rows.Count - 1
                mailBody &= "<tr>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & no & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & Convert.ToDateTime(journalData.Rows(i)("JournalDate")).ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture) & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("CustomerName").ToString() & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("Total").ToString() & "</td>"
                mailBody &= "</tr>"
                no += 1
            Next
            mailBody &= "</table>"

            mailBody &= "<br /><br /><br />"

            mailBody &= "<span style='font-family: Cambria; font-size:16px;'>Kind Regards,</span>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "<span style='font-family: Cambria; font-size:16px; font-weight: bold;'>Reza Andika Pratama</span>"

            Dim myMail As New MailMessage()
            myMail.Subject = String.Format("JOURNAL SUMMARY")
            myMail.From = New MailAddress(mailServer, mailAlias)

            myMail.To.Add(salesMail)

            If Not String.IsNullOrEmpty(mailCc) Then
                For Each thisMail In mailCc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.CC.Add(thisMail.Trim())
                Next
            End If

            If Not String.IsNullOrEmpty(mailBcc) Then
                For Each thisMail In mailBcc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.Bcc.Add(thisMail.Trim())
                Next
            End If

            myMail.Body = mailBody
            myMail.IsBodyHtml = True
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

    Public Sub JournalReport()
        Try
            Dim mailData As DataRow = GetDataRow("SELECT * FROM Mailings WHERE CompanyId='2' AND Name='Journal Report' AND Active=1")
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

            Dim journalData As DataTable = GetDataTableSP("sp_Journals_Report", Nothing)
            If journalData.Rows.Count = 0 Then Exit Sub

            Dim mailBody As String = String.Empty

            Dim today As Date = Date.Today
            Dim startDate As Date = today.AddDays(-(CInt(today.DayOfWeek) + 6) Mod 7)
            Dim endDate As Date = startDate.AddDays(4)

            Dim period As String = startDate.ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture) & " - " & endDate.ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture)

            mailBody &= "<div style='font-family:Cambria; font-size:16px;'>"
            mailBody &= "<p>Hi <b>Matt McCamey</b>,</p>"
            mailBody &= "<p>Please find below a summary of the journals recorded this week for the period <b>" & period & "</b>.</p>"
            mailBody &= "</div>"
            mailBody &= "<br />"

            mailBody &= "<table style='border-collapse:collapse; width:60%; font-family:Cambria; font-size:16px;'>"
            mailBody &= "<tr style='background-color:#00B0F0; color:black; font-weight:bold; text-align:center;'>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>No</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Journal Date</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Sales Name</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Customer</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Total Customer</th>"
            mailBody &= "</tr>"

            Dim no As Integer = 1
            For i As Integer = 0 To journalData.Rows.Count - 1
                mailBody &= "<tr>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & no & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & Convert.ToDateTime(journalData.Rows(i)("JournalDate")).ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture) & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("FullName").ToString() & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("CustomerName").ToString() & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("Total").ToString() & "</td>"
                mailBody &= "</tr>"
                no += 1
            Next
            mailBody &= "</table>"

            mailBody &= "<br /><br /><br />"

            mailBody &= "<span style='font-family: Cambria; font-size:16px;'>Kind Regards,</span>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "<span style='font-family: Cambria; font-size:16px; font-weight: bold;'>Reza Andika Pratama</span>"

            Dim myMail As New MailMessage()
            myMail.Subject = String.Format("JOURNAL REPORT")
            myMail.From = New MailAddress(mailServer, mailAlias)

            myMail.To.Add(mailTo)

            If Not String.IsNullOrEmpty(mailCc) Then
                For Each thisMail In mailCc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.CC.Add(thisMail.Trim())
                Next
            End If

            If Not String.IsNullOrEmpty(mailBcc) Then
                For Each thisMail In mailBcc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.Bcc.Add(thisMail.Trim())
                Next
            End If

            myMail.Body = mailBody
            myMail.IsBodyHtml = True
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

    Public Sub NewTicket(ticketId As String)
        Try
            If String.IsNullOrEmpty(ticketId) Then Exit Sub

            Dim ticketData As DataRow = GetDataRow("SELECT * FROM Tickets WHERE Id='" & ticketId & "'")
            If ticketData Is Nothing Then Exit Sub

            Dim ticketNo As String = ticketData("TicketNo").ToString()
            Dim ticketType As String = ticketData("Type").ToString()
            Dim ticketSubject As String = ticketData("Subject").ToString()

            Dim ticketBy As String = ticketData("LoginId").ToString()
            Dim createdName As String = GetItemData("SELECT FullName FROM Logins WHERE Id='" & ticketBy & "'")
            Dim customerName As String = GetItemData("SELECT Customers.Name AS CustomerName FROM Logins LEFT JOIN Customers ON Logins.CustomerId=Customers.Id WHERE Logins.Id='" & ticketBy & "'")
            Dim companyId As String = GetItemData("SELECT Customers.CompanyId FROM Logins LEFT JOIN Customers ON Logins.CustomerId=Customers.Id WHERE Logins.Id='" & ticketBy & "'")
            Dim companyName As String = GetItemData("SELECT Name FROM Companys WHERE Id='" & companyId & "'")

            Dim mailName As String = String.Empty
            If ticketType = "General" Then mailName = "Ticket General"
            If ticketType = "Product" Then mailName = "Ticket Product"
            If ticketType = "Pricing" Then mailName = "Ticket Pricing"
            If ticketType = "Shipment" Then mailName = "Ticket Shipment"

            If String.IsNullOrEmpty(mailName) Then Exit Sub

            Dim mailData As DataRow = GetDataRow("SELECT * FROM Mailings WHERE CompanyId='" & companyId & "' AND Name='" & mailName & "' AND Active=1")
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

            Dim mailBody As String = String.Empty

            mailBody = "<span style='font-family: Cambria; font-size: 16px;'>"
            mailBody &= "Hi Team,"
            mailBody &= "<br /><br />"
            mailBody &= "A new ticket has been assigned to you with the following details:"
            mailBody &= "</span>"

            mailBody &= "<br /><br />"

            mailBody &= "<table cellpadding='3' cellspacing='0' style='font-family:Cambria; font-size: 16px;'>"
            mailBody &= "<tr><td valign='top'>Company</td><td valign='top'>:</td><td valign='top'><b>" & companyName & "</b></td></tr>"
            mailBody &= "<tr><td valign='top'>Customer Name</td><td valign='top'>:</td><td valign='top'><b>" & customerName & "</b></td></tr>"
            mailBody &= "<tr><td valign='top'>Created By</td><td valign='top'>:</td><td valign='top'><b>" & createdName & "</b></td></tr>"
            mailBody &= "<tr><td valign='top'>Ticket No</td><td valign='top'>:</td><td valign='top'><b>" & ticketNo & "</b></td></tr>"
            mailBody &= "<tr><td valign='top'>Type</td><td valign='top'>:</td><td valign='top'><b>" & ticketType & "</b></td></tr>"
            mailBody &= "<tr><td valign='top'>Subject</td><td valign='top'>:</td><td valign='top'><b>" & ticketSubject & "</b></td></tr>"
            mailBody &= "</table>"

            mailBody &= "<br /><br />"

            mailBody &= "<span style='font-family: Cambria; font-size: 16px;'>Please check it on the online ordering portal and take the necessary action as soon as possible.</span>"
            mailBody &= "<br /><br />"
            mailBody &= "<span style='font-family: Cambria; font-size: 14px; font-style: italic;'>* Please use the <u>Ticket No</u> as the keyword when searching on the Ticket page in BOOS.</span>"

            mailBody &= "<br /><br /><br />"

            mailBody &= "<span style='font-family: Cambria; font-size:16px;'>Kind Regards,</span>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "<span style='font-family: Cambria; font-size:16px; font-weight: bold;'>Reza Andika Pratama</span><span style='font-family: Cambria; font-size:16px;'> | Developer</span>"
            mailBody &= "<br />"
            mailBody &= "<span style='font-family: Cambria; font-size:16px;'>reza@bigblinds.co.id</span>"
            mailBody &= "<br /><br /><br />"

            Dim myMail As New MailMessage()
            myMail.Subject = String.Format("New Ticket : {0} #{1}", ticketNo, ticketType)
            myMail.From = New MailAddress(mailServer, mailAlias)

            If Not String.IsNullOrEmpty(mailTo) Then
                For Each thisMail In mailTo.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.To.Add(thisMail.Trim())
                Next
            End If

            If Not String.IsNullOrEmpty(mailCc) Then
                For Each thisMail In mailCc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.CC.Add(thisMail.Trim())
                Next
            End If

            If Not String.IsNullOrEmpty(mailBcc) Then
                For Each thisMail In mailBcc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.Bcc.Add(thisMail.Trim())
                Next
            End If

            myMail.Body = mailBody
            myMail.IsBodyHtml = True
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
End Class
