Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.TerminalModules
    Module TerminalClassMapping

        Public Sub CreateTerminalBlock(strLocation As String, strTagstrip As String,
                                    eOrientation As TerminalOrientationEnum, eDucktSide As DuctSideEnum)
            Dim objTermsArray As ArrayList
            Dim objMappingTermsCollection As ArrayList

            objTermsArray = GetAllTermsInLocation(strLocation, strTagstrip)
            objMappingTermsCollection = FillTerminalData(objTermsArray, strTagstrip, eDucktSide)

            DrawTerminalBlock.DrawTerminalBlock(objMappingTermsCollection.Item(0), New Autodesk.AutoCAD.Geometry.Point3d,
                                                 eOrientation, eDucktSide)

            Dim ufTTS As New ufTerminalTypeSelector
            With ufTTS
                ufTTS.ShowDialog()
                If .rbtnSignalisation.Checked Then
                    'AddPhoenixAccForSignalisation objMappingTermsCollection, eDucktSide
                ElseIf .rbtnMeasurement.Checked Then
                ElseIf .rbtnControl.Checked Then
                ElseIf .rbtnPower.Checked Then
                End If
            End With

        End Sub

        Private Function FillTerminalData(objInputList As ArrayList, strTagstrip As String, eDucktSide As DuctSideEnum) As ArrayList
            Dim objResultArray As New ArrayList

            For Each strTempItem As String In objInputList
                objResultArray.Add(FillTermTypeData(strTagstrip, strTempItem))
            Next
            For Each objTempItem As TerminalClass In objResultArray
                FillTerminalBlockPath(objTempItem)
                FillTerminalConnectionsData(objTempItem, eDucktSide)
            Next

            Return objResultArray

        End Function

    End Module
End Namespace
