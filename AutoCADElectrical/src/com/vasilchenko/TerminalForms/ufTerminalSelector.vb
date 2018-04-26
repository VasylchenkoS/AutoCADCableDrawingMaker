Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.TerminalModules

Public Class ufTerminalSelector

    Private Sub ufTerminalSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim objLocationList As ArrayList = DBDataAccessObject.GetAllLocations()

        cbxTerminal.Enabled = False
        cbxOrientation.Enabled = False
        cbxDuctSide.Enabled = False

        If IsNothing(objLocationList) Then
            MsgBox("Проект пуст иил не указаны местоположения", MsgBoxStyle.Critical, "Warning")
            Me.Close()
        End If
        cbxLocation.DataSource = objLocationList
        cbxOrientation.DataSource = New EnumDescriptorCollection(Of OrientationEnum)
        cbxDuctSide.DataSource = New EnumDescriptorCollection(Of SideEnum)
    End Sub

    Private Sub cbxLocation_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbxLocation.SelectedValueChanged
        cbxTerminal.DataSource = Nothing
        cbxOrientation.Enabled = False
        cbxDuctSide.Enabled = False

        If cbxLocation.Text <> "" Then
            cbxTerminal.DataSource = DBDataAccessObject.GetAllTagstripInLocation(cbxLocation.Text)
            cbxTerminal.Enabled = True
        Else : cbxTerminal.Enabled = False
        End If
    End Sub

    Private Sub cbxTerminal_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbxTerminal.SelectedValueChanged
        If cbxTerminal.Text <> "" Then
            cbxOrientation.Enabled = True
            cbxDuctSide.Enabled = True
        Else
            cbxOrientation.Enabled = False
            cbxDuctSide.Enabled = False
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If cbxLocation.Text <> "" Or cbxTerminal.Text <> "" Or cbxOrientation.Text <> "" Or cbxDuctSide.Text <> "" Then
            Me.Hide()
            TerminalClassMapping.CreateTerminalBlock(cbxLocation.Text,
                                    cbxTerminal.Text,
                                    EnumFunctions.GetEnumFromDescriptionAttribute(Of OrientationEnum)(cbxOrientation.Text),
                                    EnumFunctions.GetEnumFromDescriptionAttribute(Of SideEnum)(cbxDuctSide.Text))
            Me.Close()
        Else
            MsgBox("Введите все данные", MsgBoxStyle.Critical, "DataError")
        End If
    End Sub

End Class
