

    Private Function CreateOrGetFillStyleIndex(workbookPart As WorkbookPart, fillColor As String) As UInt32Value
    Dim fills = workbookPart.WorkbookStylesPart.Stylesheet.Fills
    Dim fillCount = fills.Count()

    ' Vérifiez si le fill existe déjà
    For i As Integer = 0 To fillCount - 1
        Dim fill = fills.ChildElements(i)
        If TypeOf fill Is Fill Then
            Dim patternFill = DirectCast(fill, Fill).PatternFill
            If patternFill IsNot Nothing AndAlso
                patternFill.ForegroundColor IsNot Nothing AndAlso
                patternFill.ForegroundColor.Rgb.Value.Equals(fillColor, StringComparison.OrdinalIgnoreCase) Then
                Return i
            End If
        End If
    Next

    ' Ajoute un nouveau fill
    Dim newFill = New Fill(
        New PatternFill(
            New ForegroundColor() With {.Rgb = New HexBinaryValue(fillColor)} With {.PatternType = PatternValues.Solid}
        )
    )
    fills.AppendChild(newFill)
    fills.Count = fills.Count() + 1

    ' Sauvegardez les modifications
    workbookPart.WorkbookStylesPart.Stylesheet.Save()

    Return fillCount
End Function


' Définir un style pour la plage de cellules
Private Sub ApplyStyleToRange(workbookPart As WorkbookPart, startColumn As String, endColumn As String, startRow As Integer, endRow As Integer, styleIndex As UInt32Value)
    Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.FirstOrDefault()
    Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData).FirstOrDefault()

    For rowIndex As Integer = startRow To endRow
        Dim row As Row = sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex.Value = rowIndex).FirstOrDefault()
        If row Is Nothing Then
            row = New Row() With {.RowIndex = CType(rowIndex, UInt32)}
            sheetData.AppendChild(row)
        End If

        For colIndex As Integer = ColumnIndex(startColumn) To ColumnIndex(endColumn)
            Dim cellReference As String = GetCellReference(colIndex, rowIndex)
            Dim cell As Cell = row.Elements(Of Cell).FirstOrDefault(Function(c) c.CellReference.Value = cellReference)
            If cell Is Nothing Then
                cell = New Cell() With {.CellReference = cellReference}
                row.AppendChild(cell)
            End If
            cell.StyleIndex = styleIndex
        Next
    Next

    workbookPart.WorkbookStylesPart.Stylesheet.Save()
End Sub

' Convertir le nom de colonne en index de colonne (A -> 1, B -> 2, ...)
Private Function ColumnIndex(columnName As String) As Integer
    Dim index As Integer = 0
    For Each c As Char In columnName
        index *= 26
        index += Asc(c.ToUpper()) - Asc("A") + 1
    Next
    Return index
End Function

' Obtenir la référence de cellule en fonction de l'index de colonne et de ligne (ex : 1, 1 -> A1)
Private Function GetCellReference(columnIndex As Integer, rowIndex As Integer) As String
    Dim columnName As String = ""
    While columnIndex > 0
        Dim remainder As Integer = (columnIndex - 1) Mod 26
        columnName = Convert.ToChar(65 + remainder) & columnName
        columnIndex = (columnIndex - remainder) \ 26
    End While
    Return columnName & rowIndex.ToString()
End Function

Dim workbookPart As WorkbookPart = document.WorkbookPart
Dim styleIndex As UInt32Value = CreateOrGetFillStyleIndex(workbookPart, "FFFF0000") ' Style avec fond rouge
ApplyStyleToRange(workbookPart, "A", "AQ", 1, 10, styleIndex) ' Appliquer le style à la plage de cellules de A1 à AQ10

