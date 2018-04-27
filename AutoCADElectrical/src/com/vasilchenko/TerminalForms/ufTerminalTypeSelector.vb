Public Class ufTerminalTypeSelector
    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.rbtnSignalisation.Checked = False
        Me.rbtnMeasurement.Checked = False
        Me.rbtnControl.Checked = False
        Me.rbtnPower.Checked = False
        Me.Dispose()
    End Sub

    Private Sub rbtnControl_Click(sender As Object, e As EventArgs) Handles rbtnControl.Click
        Me.rbtnSignalisation.Checked = False
        Me.rbtnMeasurement.Checked = False
        Me.rbtnPower.Checked = False
    End Sub

    Private Sub rbtnMeasurement_Click(sender As Object, e As EventArgs) Handles rbtnMeasurement.Click
        Me.rbtnSignalisation.Checked = False
        Me.rbtnControl.Checked = False
        Me.rbtnPower.Checked = False
    End Sub


    Private Sub rbtnPower_Click(sender As Object, e As EventArgs) Handles rbtnPower.Click
        Me.rbtnSignalisation.Checked = False
        Me.rbtnMeasurement.Checked = False
        Me.rbtnControl.Checked = False
    End Sub

    Private Sub rbtnSignalisation_Click(sender As Object, e As EventArgs) Handles rbtnSignalisation.Click
        Me.rbtnMeasurement.Checked = False
        Me.rbtnControl.Checked = False
        Me.rbtnPower.Checked = False
    End Sub
End Class