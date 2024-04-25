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

Section SQL
```

WITH Priorities AS (
    SELECT *,
           CASE
               WHEN id_enregistrement > 0 AND id_modification = 0 THEN 1
               WHEN id_enregistrement > 0 AND id_modification > 0 THEN 2
               WHEN id_enregistrement = 0 AND id_modification > 0 THEN 3
               ELSE 4
           END AS priority
    FROM annuaire
)

SELECT *
FROM (
    SELECT *,
           ROW_NUMBER() OVER(PARTITION BY employee_id ORDER BY priority, id_enregistrement DESC, id_modification DESC) AS rn
    FROM Priorities
) AS ranked
WHERE rn = 1;


SELECT a.*
FROM annuaire a
INNER JOIN (
    SELECT id_employe,
           MAX(id_enregistrement) AS max_enregistrement
    FROM annuaire
    GROUP BY id_employe
) max_enr ON a.id_employe = max_enr.id_employe
INNER JOIN (
    SELECT id_employe, 
           id_enregistrement, 
           MAX(id_modification) AS max_modification
    FROM annuaire
    WHERE id_enregistrement = (
        SELECT MAX(id_enregistrement)
        FROM annuaire
        WHERE id_employe = annuaire.id_employe
    )
    GROUP BY id_employe, id_enregistrement
) max_mod ON a.id_employe = max_mod.id_employe 
           AND a.id_enregistrement = max_mod.id_enregistrement 
           AND a.id_modification = max_mod.max_modification
WHERE a.id_enregistrement = max_enr.max_enregistrement
ORDER BY a.id_employe;

SELECT a.*
FROM annuaire a
INNER JOIN (
    SELECT id_employe, 
           MAX(id_enregistrement) AS max_enregistrement
    FROM annuaire
    GROUP BY id_employe
) AS max_enr ON a.id_employe = max_enr.id_employe AND a.id_enregistrement = max_enr.max_enregistrement
LEFT JOIN (
    SELECT id_employe, 
           id_enregistrement, 
           MAX(id_modification) AS max_modification
    FROM annuaire
    GROUP BY id_employe, id_enregistrement
) AS max_mod ON a.id_employe = max_mod.id_employe 
               AND a.id_enregistrement = max_mod.id_enregistrement 
               AND a.id_modification = max_mod.max_modification
WHERE a.id_enregistrement > 0 OR (a.id_enregistrement = 0 AND a.id_modification = (
    SELECT MAX(id_modification)
    FROM annuaire
    WHERE id_employe = a.id_employe AND id_enregistrement = 0
))
ORDER BY a.id_employe;


```

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
section 4 

```
' Obtenir la feuille de calcul à partir du document Excel
Dim worksheet As Worksheet = spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements(Of Sheet).First().WorksheetPart.Worksheet

' Obtenir la première ligne de la feuille de calcul
Dim firstRow As Row = worksheet.Elements(Of SheetData).First().Elements(Of Row).First()

' Parcourir les cellules de la première ligne
For cell As Cell In firstRow.Elements(Of Cell)()
    ' Vérifier si la cellule fait partie de la plage A1:AQ1
    If cell.CellReference.Value.StartsWith("A") AndAlso cell.CellReference.Value.EndsWith("1") Then
        ' Mettre en gras la police de la cellule
        Dim font As Font = New Font() With {
            .Bold = True
        }
        Dim runProperties As RunProperties = New RunProperties() With {
            .Font = font
        }
        Dim text As Text = New Text() With {
            .Text = cell.CellValue.Text
        }
        cell.CellValue.Remove()
        cell.Append(New InlineString() With {
            .Append(runProperties),
            .Append(text)
        })
    End If
Next
```
