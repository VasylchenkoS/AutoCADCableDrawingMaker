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

    Private Sub rbtnControl_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnControl.CheckedChanged
        Me.rbtnSignalisation.Checked = False
        Me.rbtnMeasurement.Checked = False
        Me.rbtnPower.Checked = False
    End Sub

    Private Sub rbtnMeasurement_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnMeasurement.CheckedChanged
        Me.rbtnSignalisation.Checked = False
        Me.rbtnControl.Checked = False
        Me.rbtnPower.Checked = False
    End Sub

    Private Sub rbtnPower_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnPower.CheckedChanged
        Me.rbtnSignalisation.Checked = False
        Me.rbtnMeasurement.Checked = False
        Me.rbtnControl.Checked = False
    End Sub

    Private Sub rbtnSignalisation_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnSignalisation.CheckedChanged
        Me.rbtnMeasurement.Checked = False
        Me.rbtnControl.Checked = False
        Me.rbtnPower.Checked = False
    End Sub
End Class