Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports iTextSharp.text
Imports iTextSharp.text.pdf

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

    Public Function JournalSummary_PDF(salesId As String, startDate As Date, endDate As Date) As Byte()
        Try
            Using ms As New MemoryStream()
                Dim salesName As String = "ALL SALES"

                If Not String.IsNullOrWhiteSpace(salesId) Then
                    salesName = GetItemData("SELECT FullName FROM Logins WHERE Id='" & salesId.Replace("'", "''") & "'")
                    If String.IsNullOrWhiteSpace(salesName) Then salesName = "ALL SALES"
                End If

                Dim pageTitle As String = "JOURNAL SUMMARY"
                Dim doc As New Document(PageSize.A4.Rotate(), 25, 25, 125, 55)
                Dim writer As PdfWriter = PdfWriter.GetInstance(doc, ms)

                Dim pageEvent As New JournalEvents() With {
                    .PageTitle = pageTitle,
                    .PageSales = salesName,
                    .PageStartDate = startDate,
                    .PageEndDate = endDate
                }

                writer.PageEvent = pageEvent
                doc.Open()

                Dim table As New PdfPTable(4)
                table.WidthPercentage = 100
                table.SetWidths(New Single() {0.14F, 0.15F, 0.33F, 0.38F})
                table.HeaderRows = 1
                table.SplitLate = False
                table.KeepTogether = False
                table.SpacingBefore = 5
                table.SpacingAfter = 5

                Dim headerFont As New Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, BaseColor.WHITE)
                Dim bodyFont As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, New BaseColor(51, 65, 85))
                Dim emptyFont As New Font(Font.FontFamily.HELVETICA, 9, Font.ITALIC, New BaseColor(107, 114, 128))
                Dim headerBackground As New BaseColor(36, 52, 71)
                Dim alternateBackground As New BaseColor(248, 250, 252)
                Dim borderColor As New BaseColor(226, 232, 240)

                Dim headers As String() = {"JOURNAL DATE", "ACTIVITY", "CUSTOMER NAME", "DESCRIPTION"}

                For Each header As String In headers
                    Dim cell As New PdfPCell(New Phrase(header, headerFont))
                    cell.BackgroundColor = headerBackground
                    cell.HorizontalAlignment = Element.ALIGN_CENTER
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE
                    cell.PaddingTop = 9
                    cell.PaddingBottom = 9
                    cell.PaddingLeft = 5
                    cell.PaddingRight = 5
                    cell.BorderColor = borderColor
                    cell.BorderWidth = 0.5F
                    cell.Border = Rectangle.BOX
                    table.AddCell(cell)
                Next

                Dim params As New List(Of SqlParameter) From {
                    New SqlParameter("@LoginId", If(String.IsNullOrWhiteSpace(salesId), CType(DBNull.Value, Object), salesId))
                }
                Dim thisData As DataTable = GetDataTableSP("sp_Journals_Summary", params)

                If thisData.Rows.Count > 0 Then
                    For i As Integer = 0 To thisData.Rows.Count - 1
                        Dim row As DataRow = thisData.Rows(i)

                        Dim createdDate As String = Convert.ToDateTime(row("CreatedDate")).ToString("dd MMM yyyy")
                        Dim activity As String = If(IsDBNull(row("Activity")), "", row("Activity").ToString())
                        Dim customerName As String = If(IsDBNull(row("CustomerName")), "", row("CustomerName").ToString())
                        Dim description As String = If(IsDBNull(row("Description")), "", row("Description").ToString())

                        Dim rowBackground As BaseColor = If(i Mod 2 = 0, BaseColor.WHITE, alternateBackground)
                        Dim values As String() = {createdDate, activity, customerName, description}

                        For col As Integer = 0 To values.Length - 1
                            Dim cell As New PdfPCell(New Phrase(If(String.IsNullOrWhiteSpace(values(col)), "-", values(col)), bodyFont))
                            cell.BackgroundColor = rowBackground
                            cell.VerticalAlignment = Element.ALIGN_MIDDLE
                            cell.PaddingTop = 7
                            cell.PaddingBottom = 7
                            cell.PaddingLeft = 6
                            cell.PaddingRight = 6
                            cell.BorderColor = borderColor
                            cell.BorderWidth = 0.5F
                            cell.Border = Rectangle.BOX
                            If col = 0 OrElse col = 1 Then
                                cell.HorizontalAlignment = Element.ALIGN_CENTER
                            Else
                                cell.HorizontalAlignment = Element.ALIGN_LEFT
                            End If

                            table.AddCell(cell)
                        Next
                    Next
                Else
                    Dim emptyCell As New PdfPCell(New Phrase("No journal records found for the selected period.", emptyFont))
                    emptyCell.Colspan = 5
                    emptyCell.HorizontalAlignment = Element.ALIGN_CENTER
                    emptyCell.VerticalAlignment = Element.ALIGN_MIDDLE
                    emptyCell.PaddingTop = 18
                    emptyCell.PaddingBottom = 18
                    emptyCell.BackgroundColor = alternateBackground
                    emptyCell.BorderColor = borderColor
                    emptyCell.BorderWidth = 0.5F
                    table.AddCell(emptyCell)
                End If

                doc.Add(table)
                doc.Close()

                Return ms.ToArray()
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("JournalSummary_PDF Error: " & ex.ToString())
        Return New Byte() {}
        End Try
    End Function

    Public Function JournalReport_PDF(startDate As Date, endDate As Date) As Byte()
        Try
            Using ms As New MemoryStream()
                Dim salesName As String = "ALL SALES"

                Dim pageTitle As String = "JOURNAL REPORT"
                Dim doc As New Document(PageSize.A4.Rotate(), 25, 25, 125, 55)
                Dim writer As PdfWriter = PdfWriter.GetInstance(doc, ms)

                Dim pageEvent As New JournalEvents() With {
                    .PageTitle = pageTitle,
                    .PageSales = salesName,
                    .PageStartDate = startDate,
                    .PageEndDate = endDate
                }

                writer.PageEvent = pageEvent
                doc.Open()

                Dim table As New PdfPTable(5)
                table.WidthPercentage = 100
                table.SetWidths(New Single() {0.2F, 0.14F, 0.15F, 0.23F, 0.28F})
                table.HeaderRows = 1
                table.SplitLate = False
                table.KeepTogether = False
                table.SpacingBefore = 5
                table.SpacingAfter = 5

                Dim headerFont As New Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, BaseColor.WHITE)
                Dim bodyFont As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, New BaseColor(51, 65, 85))
                Dim emptyFont As New Font(Font.FontFamily.HELVETICA, 9, Font.ITALIC, New BaseColor(107, 114, 128))
                Dim headerBackground As New BaseColor(36, 52, 71)
                Dim alternateBackground As New BaseColor(248, 250, 252)
                Dim borderColor As New BaseColor(226, 232, 240)

                Dim headers As String() = {"SALES NAME", "JOURNAL DATE", "ACTIVITY", "CUSTOMER NAME", "DESCRIPTION"}

                For Each header As String In headers
                    Dim cell As New PdfPCell(New Phrase(header, headerFont))
                    cell.BackgroundColor = headerBackground
                    cell.HorizontalAlignment = Element.ALIGN_CENTER
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE
                    cell.PaddingTop = 9
                    cell.PaddingBottom = 9
                    cell.PaddingLeft = 5
                    cell.PaddingRight = 5
                    cell.BorderColor = borderColor
                    cell.BorderWidth = 0.5F
                    cell.Border = Rectangle.BOX
                    table.AddCell(cell)
                Next

                Dim thisData As DataTable = GetDataTableSP("sp_Journals_Report", Nothing)

                If thisData.Rows.Count > 0 Then
                    For i As Integer = 0 To thisData.Rows.Count - 1
                        Dim row As DataRow = thisData.Rows(i)

                        Dim createdBy As String = If(IsDBNull(row("CreatedByFull")), "", row("CreatedByFull").ToString())
                        Dim createdDate As String = Convert.ToDateTime(row("CreatedDate")).ToString("dd MMM yyyy")
                        Dim activity As String = If(IsDBNull(row("Activity")), "", row("Activity").ToString())
                        Dim customerName As String = If(IsDBNull(row("CustomerName")), "", row("CustomerName").ToString())
                        Dim description As String = If(IsDBNull(row("Description")), "", row("Description").ToString())

                        Dim rowBackground As BaseColor = If(i Mod 2 = 0, BaseColor.WHITE, alternateBackground)
                        Dim values As String() = {createdBy, createdDate, activity, customerName, description}

                        For col As Integer = 0 To values.Length - 1
                            Dim cell As New PdfPCell(New Phrase(If(String.IsNullOrWhiteSpace(values(col)), "-", values(col)), bodyFont))
                            cell.BackgroundColor = rowBackground
                            cell.VerticalAlignment = Element.ALIGN_MIDDLE
                            cell.PaddingTop = 7
                            cell.PaddingBottom = 7
                            cell.PaddingLeft = 6
                            cell.PaddingRight = 6
                            cell.BorderColor = borderColor
                            cell.BorderWidth = 0.5F
                            cell.Border = Rectangle.BOX

                            If col = 0 OrElse col = 1 OrElse col = 2 Then
                                cell.HorizontalAlignment = Element.ALIGN_CENTER
                            Else
                                cell.HorizontalAlignment = Element.ALIGN_LEFT
                            End If

                            table.AddCell(cell)
                        Next
                    Next
                Else
                    Dim emptyCell As New PdfPCell(New Phrase("No journal records found for the selected period.", emptyFont))
                    emptyCell.Colspan = 5
                    emptyCell.HorizontalAlignment = Element.ALIGN_CENTER
                    emptyCell.VerticalAlignment = Element.ALIGN_MIDDLE
                    emptyCell.PaddingTop = 18
                    emptyCell.PaddingBottom = 18
                    emptyCell.BackgroundColor = alternateBackground
                    emptyCell.BorderColor = borderColor
                    emptyCell.BorderWidth = 0.5F
                    table.AddCell(emptyCell)
                End If

                doc.Add(table)
                doc.Close()

                Return ms.ToArray()
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("JournalSummary_PDF Error: " & ex.ToString())
            Return New Byte() {}
        End Try
    End Function


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
            Dim journalData As DataTable = GetDataTableSP("sp_Journals_Summary_Group", params)
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

            mailBody &= "<table style='border-collapse:collapse; width:100%; font-family:Cambria; font-size:16px;'>"
            mailBody &= "<tr style='background-color:#00B0F0; color:black; font-weight:bold; text-align:center;'>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>No</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Journal Date</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Activity</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Customer</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Total Customer</th>"
            mailBody &= "</tr>"

            Dim no As Integer = 1
            For i As Integer = 0 To journalData.Rows.Count - 1
                mailBody &= "<tr>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & no & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & Convert.ToDateTime(journalData.Rows(i)("JournalDate")).ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture) & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("Activity").ToString() & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px;'>" & journalData.Rows(i)("CustomerName").ToString() & "</td>"
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

            myMail.IsBodyHtml = True
            myMail.Body = mailBody

            Dim pdfBytes As Byte() = JournalSummary_PDF(salesId, startDate, endDate)

            Dim ms As New MemoryStream(pdfBytes)
            Dim attach As New Attachment(ms, "JOURNAL SUMMARY - " & salesName.ToUpper() & ".pdf", "application/pdf")
            myMail.Attachments.Add(attach)

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

            Dim journalData As DataTable = GetDataTableSP("sp_Journals_Report_Group", Nothing)
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

            mailBody &= "<table style='border-collapse:collapse; width:100%; font-family:Cambria; font-size:16px;'>"
            mailBody &= "<tr style='background-color:#00B0F0; color:black; font-weight:bold; text-align:center;'>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>No</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Journal Date</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Sales Name</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Activity</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Customer</th>"
            mailBody &= "<th style='border:1px solid #000; padding:6px;'>Total Customer</th>"
            mailBody &= "</tr>"

            Dim no As Integer = 1
            For i As Integer = 0 To journalData.Rows.Count - 1
                mailBody &= "<tr>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & no & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & Convert.ToDateTime(journalData.Rows(i)("JournalDate")).ToString("dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture) & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("FullName").ToString() & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px; text-align:center;'>" & journalData.Rows(i)("Activity").ToString() & "</td>"
                mailBody &= "<td style='border:1px solid #000; padding:6px;'>" & journalData.Rows(i)("CustomerName").ToString() & "</td>"
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

            myMail.IsBodyHtml = True
            myMail.Body = mailBody

            Dim pdfBytes As Byte() = JournalReport_PDF(startDate, endDate)

            Dim ms As New MemoryStream(pdfBytes)
            Dim attach As New Attachment(ms, "JOURNAL REPORT.pdf", "application/pdf")
            myMail.Attachments.Add(attach)

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

Public Class JournalEvents
    Inherits PdfPageEventHelper

    Public PageTitle As String
    Public PageSales As String
    Public PageStartDate As Date?
    Public PageEndDate As Date?

    Public Overrides Sub OnEndPage(writer As PdfWriter, document As Document)
        Dim cb As PdfContentByte = writer.DirectContent

        Dim titleFont As New Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, New BaseColor(36, 52, 71))
        Dim labelFont As New Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, New BaseColor(107, 114, 128))
        Dim valueFont As New Font(Font.FontFamily.HELVETICA, 9, Font.BOLD, New BaseColor(31, 41, 55))
        Dim smallFont As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, New BaseColor(107, 114, 128))
        Dim footerFont As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, New BaseColor(107, 114, 128))

        Dim leftMargin As Single = document.LeftMargin
        Dim rightMargin As Single = document.PageSize.Width - document.RightMargin
        Dim contentWidth As Single = rightMargin - leftMargin

        ' HEADER TITLE
        Dim titleTable As New PdfPTable(2)
        titleTable.TotalWidth = contentWidth
        titleTable.LockedWidth = True
        titleTable.SetWidths(New Single() {0.68F, 0.32F})

        Dim titleCell As New PdfPCell(New Phrase(If(String.IsNullOrWhiteSpace(PageTitle), "REPORT", PageTitle), titleFont))
        titleCell.Border = Rectangle.NO_BORDER
        titleCell.HorizontalAlignment = Element.ALIGN_LEFT
        titleCell.VerticalAlignment = Element.ALIGN_MIDDLE
        titleCell.PaddingBottom = 5
        titleTable.AddCell(titleCell)

        Dim generatedCell As New PdfPCell(New Phrase("GENERATED ON" & vbLf & DateTime.Now.ToString("dd MMM yyyy HH:mm"), smallFont))
        generatedCell.Border = Rectangle.NO_BORDER
        generatedCell.HorizontalAlignment = Element.ALIGN_RIGHT
        generatedCell.VerticalAlignment = Element.ALIGN_MIDDLE
        generatedCell.PaddingBottom = 5
        titleTable.AddCell(generatedCell)

        titleTable.WriteSelectedRows(0, -1, leftMargin, document.PageSize.Height - 22, cb)

        ' HEADER DIVIDER
        Dim lineY As Single = document.PageSize.Height - 57

        cb.SaveState()
        cb.SetColorStroke(New BaseColor(37, 99, 235))
        cb.SetLineWidth(2.0F)
        cb.MoveTo(leftMargin, lineY)
        cb.LineTo(rightMargin, lineY)
        cb.Stroke()
        cb.RestoreState()

        ' COMPANY AND REPORT PERIOD
        Dim infoTable As New PdfPTable(2)
        infoTable.TotalWidth = contentWidth
        infoTable.LockedWidth = True
        infoTable.SetWidths(New Single() {0.5F, 0.5F})

        Dim companyCell As New PdfPCell()
        companyCell.Border = Rectangle.NO_BORDER
        companyCell.HorizontalAlignment = Element.ALIGN_LEFT
        companyCell.VerticalAlignment = Element.ALIGN_TOP
        companyCell.PaddingTop = 9
        companyCell.PaddingBottom = 5
        companyCell.PaddingLeft = 0
        companyCell.PaddingRight = 10
        companyCell.AddElement(New Paragraph("SALES NAME", labelFont))
        companyCell.AddElement(New Paragraph(If(String.IsNullOrWhiteSpace(PageSales), "ALL COMPANIES", PageSales), valueFont))
        infoTable.AddCell(companyCell)

        Dim periodText As String = "All Dates"

        If PageStartDate.HasValue AndAlso PageEndDate.HasValue Then
            periodText = PageStartDate.Value.ToString("dd MMM yyyy") & " - " & PageEndDate.Value.ToString("dd MMM yyyy")
        ElseIf PageStartDate.HasValue Then
            periodText = "From " & PageStartDate.Value.ToString("dd MMM yyyy")
        ElseIf PageEndDate.HasValue Then
            periodText = "Until " & PageEndDate.Value.ToString("dd MMM yyyy")
        End If

        Dim periodCell As New PdfPCell()
        periodCell.Border = Rectangle.NO_BORDER
        periodCell.HorizontalAlignment = Element.ALIGN_LEFT
        periodCell.VerticalAlignment = Element.ALIGN_TOP
        periodCell.PaddingTop = 9
        periodCell.PaddingBottom = 5
        periodCell.PaddingLeft = 0
        periodCell.PaddingRight = 10
        periodCell.AddElement(New Paragraph("JOURNAL PERIOD", labelFont))
        periodCell.AddElement(New Paragraph(periodText, valueFont))
        infoTable.AddCell(periodCell)

        infoTable.WriteSelectedRows(0, -1, leftMargin, lineY - 3, cb)

        ' FOOTER DIVIDER
        Dim footerY As Single = document.BottomMargin - 5

        cb.SaveState()
        cb.SetColorStroke(New BaseColor(226, 232, 240))
        cb.SetLineWidth(0.7F)
        cb.MoveTo(leftMargin, footerY + 15)
        cb.LineTo(rightMargin, footerY + 15)
        cb.Stroke()
        cb.RestoreState()

        ' FOOTER
        Dim footerTable As New PdfPTable(2)
        footerTable.TotalWidth = contentWidth
        footerTable.LockedWidth = True
        footerTable.SetWidths(New Single() {0.75F, 0.25F})

        Dim leftFooterCell As New PdfPCell(New Phrase("CONFIDENTIAL  |  All information within this report is private and confidential.", footerFont))
        leftFooterCell.Border = Rectangle.NO_BORDER
        leftFooterCell.HorizontalAlignment = Element.ALIGN_LEFT
        leftFooterCell.VerticalAlignment = Element.ALIGN_MIDDLE
        leftFooterCell.PaddingTop = 5
        footerTable.AddCell(leftFooterCell)

        Dim rightFooterCell As New PdfPCell(New Phrase("Page " & writer.PageNumber, New Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, New BaseColor(36, 52, 71))))
        rightFooterCell.Border = Rectangle.NO_BORDER
        rightFooterCell.HorizontalAlignment = Element.ALIGN_RIGHT
        rightFooterCell.VerticalAlignment = Element.ALIGN_MIDDLE
        rightFooterCell.PaddingTop = 5
        footerTable.AddCell(rightFooterCell)

        footerTable.WriteSelectedRows(0, -1, leftMargin, footerY + 10, cb)
    End Sub
End Class