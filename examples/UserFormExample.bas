
Private Sub UserForm_Initialize()

    ' Initialize dropdown lists
    cmbTestType.List = Array( _
        "Release", _
        "Stability", _
        "Transfer", _
        "Special Test")

    cmbSamplePosition.List = Array( _
        "Beginning", _
        "Middle", _
        "End")

End Sub



Private Function ValidateDate(inputDate As String) As Boolean

    If IsDate(inputDate) Then
        ValidateDate = True
    Else
        MsgBox "Please enter a valid date."
        ValidateDate = False
    End If

End Function




Private Sub SaveValue(ws As Worksheet, row As Long, col As Long, value As Variant)

    If value = "" Then
        ws.Cells(row, col).Value = "-"
    Else
        ws.Cells(row, col).Value = value
    End If

End Sub




Sub UpdateTrendChart()

    With Worksheets("Dashboard").ChartObjects("TrendChart").Chart
        .Refresh
    End With

End Sub




Private Sub ClearFields()

    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.Value = ""
        End If
    Next ctrl

End Sub
