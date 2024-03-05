```
Protected Sub ExportToExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExportButton.Click
    Response.Clear()
    Response.Buffer = True
    Response.AddHeader("content-disposition", "attachment;filename=GridViewExport.xlsx")
    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    Using memStream As New MemoryStream()
        Using document As SpreadsheetDocument = SpreadsheetDocument.Create(memStream, SpreadsheetDocumentType.Workbook)
            Dim workbookPart As WorkbookPart = document.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()

            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            Dim sheetData As New SheetData()
            worksheetPart.Worksheet = New Worksheet(sheetData)

            Dim sheets As Sheets = workbookPart.Workbook.AppendChild(New Sheets())
            Dim sheet As New Sheet() With {.Id = workbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "Sheet1"}
            sheets.Append(sheet)

            ' Ajouter les en-têtes de colonnes à la feuille Excel
            Dim headerRow As New Row()
            For Each column As DataColumn In YourDataTable.Columns
                Dim cell As New Cell(New InlineString(New Text(column.ColumnName)))
                cell.StyleIndex = 1 ' Mettre en gras
                headerRow.AppendChild(cell)
            Next
            sheetData.AppendChild(headerRow)

            ' Ajouter les données de la DataTable à la feuille Excel
            For Each rowItem As DataRow In YourDataTable.Rows
                Dim row As New Row()
                For Each column As DataColumn In YourDataTable.Columns
                    Dim cell As New Cell(New InlineString(New Text(rowItem(column).ToString())))
                    cell.StyleIndex = 1 ' Mettre en gras
                    row.AppendChild(cell)
                Next
                sheetData.AppendChild(row)
            Next

            ' Créer un nouveau style pour mettre en gras
            Dim boldStyleIndex As UInt32Value = CreateBoldCellStyle(workbookPart)

            ' Enregistrement du classeur
            workbookPart.Workbook.Save()
        End Using

        memStream.Seek(0, SeekOrigin.Begin)
        memStream.CopyTo(Response.OutputStream)
        Response.Flush()
        Response.End()
    End Using
End Sub

Private Function CreateBoldCellStyle(ByVal workbookPart As WorkbookPart) As UInt32Value
    Dim stylesheet As Stylesheet = workbookPart.WorkbookStylesPart.Stylesheet

    ' Créer un nouveau style pour mettre en gras
    Dim boldFont As New Font With {.Bold = New Bold}
    Dim boldFontId As UInt32Value = AddFontToStylesheet(stylesheet, boldFont)

    Dim cellFormats As CellFormats = stylesheet.CellFormats
    Dim boldStyleIndex As UInt32Value = AppendCellFormat(cellFormats, boldFontId)

    Return boldStyleIndex
End Function

Private Function AddFontToStylesheet(ByVal stylesheet As Stylesheet, ByVal font As Font) As UInt32Value
    Dim fonts As Fonts = If(stylesheet.Fonts, New Fonts())
    fonts.Append(font)
    stylesheet.Fonts = fonts
    Return Convert.ToUInt32(fonts.Count - 1)
End Function

Private Function AppendCellFormat(ByVal cellFormats As CellFormats, ByVal fontId As UInt32Value) As UInt32Value
    Dim cellFormat As New CellFormat With {.FontId = fontId}
    cellFormats.Append(cellFormat)
    Return Convert.ToUInt32(cellFormats.Count - 1)
End Function

```
Section 2

```
Protected Sub ExportToExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExportButton.Click
    Response.Clear()
    Response.Buffer = True
    Response.AddHeader("content-disposition", "attachment;filename=GridViewExport.xlsx")
    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    Using memStream As New MemoryStream()
        Using document As SpreadsheetDocument = SpreadsheetDocument.Create(memStream, SpreadsheetDocumentType.Workbook)
            Dim workbookPart As WorkbookPart = document.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()

            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            Dim sheetData As New SheetData()
            worksheetPart.Worksheet = New Worksheet(sheetData)

            Dim sheets As Sheets = workbookPart.Workbook.AppendChild(New Sheets())
            Dim sheet As New Sheet() With {.Id = workbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "Sheet1"}
            sheets.Append(sheet)

            ' Ajouter les en-têtes de colonnes à la feuille Excel
            Dim headerRow As New Row()
            For Each column As DataColumn In YourDataTable.Columns
                Dim cell As New Cell(New InlineString(New Text(column.ColumnName)))
                cell.StyleIndex = 1 ' Mettre en gras
                headerRow.AppendChild(cell)
            Next
            sheetData.AppendChild(headerRow)

            ' Ajouter les données de la DataTable à la feuille Excel
            For Each rowItem As DataRow In YourDataTable.Rows
                Dim row As New Row()
                For Each column As DataColumn In YourDataTable.Columns
                    Dim cell As New Cell(New InlineString(New Text(rowItem(column).ToString())))
                    row.AppendChild(cell)
                Next
                sheetData.AppendChild(row)
            Next

            ' Créer un nouveau style pour mettre en gras
            Dim boldStyleIndex As UInt32Value = CreateBoldCellStyle(workbookPart)

            ' Mettre en gras la plage spécifique (de A1 à AQ1)
            Dim boldCellsRange As String = "A1:AQ1"
            Dim boldRangeParts As String() = boldCellsRange.Split(":"c)
            Dim boldStartCell As String = boldRangeParts(0)
            Dim boldEndCell As String = boldRangeParts(1)
            ApplyCellStyleToRange(worksheetPart, boldStartCell, boldEndCell, boldStyleIndex)

            ' Enregistrement du classeur
            workbookPart.Workbook.Save()
        End Using

        memStream.Seek(0, SeekOrigin.Begin)
        memStream.CopyTo(Response.OutputStream)
        Response.Flush()
        Response.End()
    End Using
End Sub

Private Sub ApplyCellStyleToRange(ByVal worksheetPart As WorksheetPart, ByVal startCellReference As String, ByVal endCellReference As String, ByVal styleIndex As UInt32Value)
    Dim worksheet As Worksheet = worksheetPart.Worksheet
    Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
    Dim startCell As Cell = GetOrCreateCell(sheetData, startCellReference)
    Dim endCell As Cell = GetOrCreateCell(sheetData, endCellReference)

    ' Appliquer le style à chaque cellule dans la plage
    Dim startRowIndex As Integer = GetRowIndex(startCellReference)
    Dim endRowIndex As Integer = GetRowIndex(endCellReference)

    For rowIndex As Integer = startRowIndex To endRowIndex
        Dim row As Row = GetOrCreateRow(sheetData, rowIndex)

        For Each cell As Cell In row.Elements(Of Cell)()
            cell.StyleIndex = styleIndex
        Next
    Next
End Sub

Private Function GetOrCreateCell(ByVal sheetData As SheetData, ByVal cellReference As String) As Cell
    Dim column As String = Regex.Replace(cellReference, "[0-9]", "")
    Dim row As String = Regex.Replace(cellReference, "[A-Za-z]", "")
    Dim targetRow As Row = GetOrCreateRow(sheetData, Integer.Parse(row))

    For Each targetCell As Cell In targetRow.Elements(Of Cell)()
        If String.Equals(targetCell.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase) Then
            Return targetCell
        End If
    Next

    Dim newCell As New Cell() With {.CellReference = cellReference}
    targetRow.InsertAt(newCell, GetCellIndex(targetRow, column))
    Return newCell
End Function

Private Function GetOrCreateRow(ByVal sheetData As SheetData, ByVal rowIndex As Integer) As Row
    Dim targetRow As Row = sheetData.Elements(Of Row).FirstOrDefault(Function(r) r.RowIndex.Value = rowIndex)

    If targetRow Is Nothing Then
        targetRow = New Row With {.RowIndex = rowIndex}
        Dim previousRow As Row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value < rowIndex).LastOrDefault()

        If previousRow IsNot Nothing Then
            sheetData.InsertAfter(targetRow, previousRow)
        Else
            sheetData.PrependChild(targetRow)
        End If
    End If

    Return targetRow
End Function

Private Function GetRowIndex(ByVal cellReference As String) As Integer
    Dim rowString As String = Regex.Replace(cellReference, "[A-Za-z]", "")
    Return Integer.Parse(rowString)
End Function

Private Function GetCellIndex(ByVal row As Row, ByVal columnReference As String) As Integer
    Dim index As Integer = 0

    For Each cell As Cell In row.Elements(Of Cell)()
        If String.Compare(cell.CellReference.Value, columnReference, StringComparison.OrdinalIgnoreCase) > 0 Then
            Exit For
        End If

        index += 1
    Next

    Return index
End Function

Private Sub ApplyCellStyleToRange(ByVal worksheetPart As WorksheetPart, ByVal startCellReference As String, ByVal endCellReference As String, ByVal styleIndex As UInt32Value)
    Dim worksheet As Worksheet = worksheetPart.Worksheet
    Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
    Dim startCell As Cell = GetOrCreateCell(sheetData, startCellReference)
    Dim endCell As Cell = GetOrCreateCell(sheetData, endCellReference)

    ' Appliquer le style à chaque cellule dans la plage
    Dim startRowIndex As Integer = GetRowIndex(startCellReference)
    Dim endRowIndex As Integer = GetRowIndex(endCellReference)

    For rowIndex As Integer = startRowIndex To endRowIndex
        Dim row As Row = GetOrCreateRow(sheetData, rowIndex)

        For Each cell As Cell In row.Elements(Of Cell)()
            cell.StyleIndex = styleIndex
        Next
    Next
End Sub

Private Function GetOrCreateCell(ByVal sheetData As SheetData, ByVal cellReference As String) As Cell
    Dim column As String = Regex.Replace(cellReference, "[0-9]", "")
    Dim row As String = Regex.Replace(cellReference, "[A-Za-z]", "")
    Dim targetRow As Row = GetOrCreateRow(sheetData, Integer.Parse(row))

    For Each targetCell As Cell In targetRow.Elements(Of Cell)()
        If String.Equals(targetCell.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase) Then
            Return targetCell
        End If
    Next

    Dim newCell As New Cell() With {.CellReference = cellReference}
    targetRow.InsertAt(newCell, GetCellIndex(targetRow, column))
    Return newCell
End Function

Private Function GetOrCreateRow(ByVal sheetData As SheetData, ByVal rowIndex As Integer) As Row
    Dim targetRow As Row = sheetData.Elements(Of Row).FirstOrDefault(Function(r) r.RowIndex.Value = rowIndex)

    If targetRow Is Nothing Then
        targetRow = New Row With {.RowIndex = rowIndex}
        Dim previousRow As Row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value < rowIndex).LastOrDefault()

        If previousRow IsNot Nothing Then
            sheetData.InsertAfter(targetRow, previousRow)
        Else
            sheetData.PrependChild(targetRow)
        End If
    End If

    Return targetRow
End Function

Private Function GetRowIndex(ByVal cellReference As String) As Integer
    Dim rowString As String = Regex.Replace(cellReference, "[A-Za-z]", "")
    Return Integer.Parse(rowString)
End Function

Private Function GetCellIndex(ByVal row As Row, ByVal columnReference As String) As Integer
    Dim index As Integer = 0

    For Each cell As Cell In row.Elements(Of Cell)()
        If String.Compare(cell.CellReference.Value, columnReference, StringComparison.OrdinalIgnoreCase) > 0 Then
            Exit For
        End If

        index += 1
    Next

    Return index
End Function
```
Section 3

```
  Protected Sub ExportToExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExportButton.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=GridViewExport.xlsx")
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        Using memStream As New MemoryStream()
            Using document As SpreadsheetDocument = SpreadsheetDocument.Create(memStream, SpreadsheetDocumentType.Workbook)
                Dim workbookPart As WorkbookPart = document.AddWorkbookPart()
                workbookPart.Workbook = New Workbook()

                Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                Dim sheetData As New SheetData()
                worksheetPart.Worksheet = New Worksheet(sheetData)

                Dim sheets As Sheets = workbookPart.Workbook.AppendChild(New Sheets())
                Dim sheet As New Sheet() With {.Id = workbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "Sheet1"}
                sheets.Append(sheet)

                ' Ajouter les en-têtes de colonnes à la feuille Excel
                Dim headerRow As New Row()
                For Each column As DataColumn In YourDataTable.Columns
                    Dim cell As New Cell(New InlineString(New Text(column.ColumnName)))
                    headerRow.AppendChild(cell)
                Next
                sheetData.AppendChild(headerRow)

                ' Ajouter les données de la DataTable à la feuille Excel
                For Each rowItem As DataRow In YourDataTable.Rows
                    Dim row As New Row()
                    For Each column As DataColumn In YourDataTable.Columns
                        Dim cell As New Cell(New InlineString(New Text(rowItem(column).ToString())))
                        row.AppendChild(cell)
                    Next
                    sheetData.AppendChild(row)
                Next

                ' Créer un nouveau style pour mettre en gras
                Dim boldStyleIndex As UInt32Value = CreateBoldCellStyle(workbookPart)

                ' Appliquer le style en gras à la plage spécifiée (A1:AQ1)
                ApplyStyleToRange(worksheetPart.Worksheet, "A1:AQ1", boldStyleIndex)

                ' Enregistrement du classeur
                workbookPart.Workbook.Save()
            End Using

            memStream.Seek(0, SeekOrigin.Begin)
            memStream.CopyTo(Response.OutputStream)
            Response.Flush()
            Response.End()
        End Using
    End Sub

    Private Function CreateBoldCellStyle(ByVal workbookPart As WorkbookPart) As UInt32Value
        Dim stylesheet As Stylesheet = workbookPart.WorkbookStylesPart.Stylesheet

        ' Créer un nouveau style pour mettre en gras
        Dim boldFont As New Font With {.Bold = New Bold}
        Dim boldFontId As UInt32Value = AddFontToStylesheet(stylesheet, boldFont)

        Dim cellFormats As CellFormats = stylesheet.CellFormats
        Dim boldStyleIndex As UInt32Value = AppendCellFormat(cellFormats, boldFontId)

        Return boldStyleIndex
    End Function

    Private Function AddFontToStylesheet(ByVal stylesheet As Stylesheet, ByVal font As Font) As UInt32Value
        Dim fonts As Fonts = If(stylesheet.Fonts, New Fonts())
        fonts.Append(font)
        stylesheet.Fonts = fonts
        Return Convert.ToUInt32(fonts.Count - 1)
    End Function

    Private Function AppendCellFormat(ByVal cellFormats As CellFormats, ByVal fontId As UInt32Value) As UInt32Value
        Dim cellFormat As New CellFormat With {.FontId = fontId}
        cellFormats.Append(cellFormat)
        Return Convert.ToUInt32(cellFormats.Count - 1)
    End Function

    Private Sub ApplyStyleToRange(ByVal worksheet As Worksheet, ByVal range As String, ByVal styleIndex As UInt32Value)
        Dim cells As IEnumerable(Of Cell) = worksheet.Descendants(Of Cell)().Where(Function(c) IsInRange(c.CellReference.Value, range))

        For Each cell In cells
            cell.StyleIndex = styleIndex
        Next
    End Sub

    Private Function IsInRange(ByVal cellReference As String, ByVal range As String) As Boolean
        Dim parts As String() = range.Split(":"c)
        Dim startCell As String = parts(0)
        Dim endCell As String = parts(1)

        Dim startCol As String = Regex.Match(startCell, "[A-Za-z]+").Value
        Dim startRow As Integer = Integer.Parse(Regex.Match(startCell, "\d+").Value)

        Dim endCol As String = Regex.Match(endCell, "[A-Za-z]+").Value
        Dim endRow As Integer = Integer.Parse(Regex.Match(endCell, "\d+").Value)

        Dim cellCol As String = Regex.Match(cellReference, "[A-Za-z]+").Value
        Dim cellRow As Integer = Integer.Parse(Regex.Match(cellReference, "\d+").Value)

        Return cellCol >= startCol AndAlso cellCol <= endCol AndAlso cellRow >= startRow AndAlso cellRow <= endRow
    End Function
```
