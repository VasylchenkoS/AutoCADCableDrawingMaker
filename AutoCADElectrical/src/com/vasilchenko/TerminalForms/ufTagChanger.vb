Public Class ufTagChanger
    Private Sub ufTagChanger_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbReplace.Text = ""
        tbSearch.Text = ""
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.Hide 
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        tbReplace.Text = ""
        tbSearch.Text = ""
        Me.Close 
    End Sub
End Class