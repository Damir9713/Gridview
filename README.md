

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
