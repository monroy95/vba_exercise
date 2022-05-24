Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Sheets("DB").Visible = False

    If Target.Column = 3 Then
        If Target.Row = 5 Then
            Beep
            Cells(Target.Row, Target.Column).Offset(0, 1).Select
        End If
    End If
End Sub
