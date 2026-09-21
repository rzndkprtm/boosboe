Imports System.Data
Imports System.Data.SqlClient

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
End Class
