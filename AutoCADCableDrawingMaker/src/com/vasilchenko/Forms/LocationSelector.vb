Imports AutoCADCableDrawingMaker.com.vasilchenko
Imports AutoCADCableDrawingMaker.com.vasilchenko.Modules

Public Class LocationSelector
    Private Sub LocationSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim objLocationList As List(Of String) = Database.DataAccessObject.GetAllLocations()
        objLocationList.Sort(Function(x, y) AdditionalFunctions.GetLastNumericFromString(x).CompareTo(AdditionalFunctions.GetLastNumericFromString(y)))

        If IsNothing(objLocationList) Then
            Throw New ArgumentNullException("Проект пуст или не указаны местоположения")
            Me.Close()
        Else
            cbxLocation.DataSource = objLocationList
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If Not cbxLocation.Text.Equals("") Then
            Me.Hide()
            Modules.CableClassMapping.CreateTerminalBlock(cbxLocation.Text)
            Me.Close()
        Else
            MsgBox("Введите все данные", MsgBoxStyle.Critical, "DataError")
        End If
    End Sub
End Class