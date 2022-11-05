 Private Sub SetInitialRow()
        Dim dt As New DataTable()
        dt.Columns.Add(New DataColumn("RowNumber", GetType(String)))
        Dim dr As DataRow = dt.NewRow()
        dt.Rows.Add(dr)
        ViewState("CurrentTable") = dt
        Gridview1.DataSource = dt
        Gridview1.DataBind()
        Gridview1.Rows(0).Visible = False
    End Sub

    Protected Sub ButtonAdd_Click(ByVal sender As Object, ByVal e As EventArgs)
        AddNewRowToGrid()
    End Sub


    Private Sub AddNewRowToGrid()

        If ViewState("CurrentTable") IsNot Nothing Then
            Dim dtCurrentTable As DataTable = CType(ViewState("CurrentTable"), DataTable)

            Dim ddl As DropDownList = CType(Gridview1.FooterRow.FindControl("DropDownList1"), DropDownList)
            Dim drCurrentRow As DataRow = dtCurrentTable.NewRow()
            drCurrentRow("RowNumber") = ddl.SelectedItem.Text

            dtCurrentTable.Rows.Add(drCurrentRow)
            ViewState("CurrentTable") = dtCurrentTable
            Gridview1.DataSource = dtCurrentTable
            Gridview1.DataBind()
            Gridview1.Rows(0).Visible = False
        End If
        SetPreviousData()
    End Sub

    Private Sub SetPreviousData()
        Dim rowIndex As Integer = 0

        If ViewState("CurrentTable") IsNot Nothing Then
            Dim dt As DataTable = CType(ViewState("CurrentTable"), DataTable)

            If dt.Rows.Count > 0 Then

                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim label As Label = CType(Gridview1.Rows(rowIndex).Cells(1).FindControl("Label1"), Label)

                    label.Text = dt.Rows(i)("RowNumber").ToString()

                    rowIndex += 1

                Next
            End If
        End If
    End Sub

    Protected Sub LinkButton1_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim lb As LinkButton = CType(sender, LinkButton)
        Dim gvRow As GridViewRow = CType(lb.NamingContainer, GridViewRow)
        Dim rowID As Integer = gvRow.RowIndex

        If ViewState("CurrentTable") IsNot Nothing Then
            Dim dt As DataTable = CType(ViewState("CurrentTable"), DataTable)

            If dt.Rows.Count > 1 Then

                If gvRow.RowIndex < dt.Rows.Count Then
                    dt.Rows.Remove(dt.Rows(rowID))
                End If
            End If

            ViewState("CurrentTable") = dt
            Gridview1.DataSource = dt
            Gridview1.DataBind()
            Gridview1.Rows(0).Visible = False
        End If

        SetPreviousData()
    End Sub
