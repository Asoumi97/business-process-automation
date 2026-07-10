Public Function ValidateDate(value As String) As Boolean

    ValidateDate = IsDate(value)

End Function




Public Sub SaveRecord(ws As Worksheet, row As Long)

    ws.Cells(row, 2).Value = txtName.Value

End Sub



Public Sub RefreshTrendChart()

    Worksheets("Dashboard").ChartObjects("TrendChart").Chart.Refresh

End Sub



Public Function FormatValue(value As Variant) As Variant

    If value = "" Then
        FormatValue = "-"
    Else
        FormatValue = value
    End If

End Function



Public Sub ClearTextBoxes(frm As Object)

    Dim ctrl As Control

    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.Value = ""
        End If
    Next ctrl

End Sub
