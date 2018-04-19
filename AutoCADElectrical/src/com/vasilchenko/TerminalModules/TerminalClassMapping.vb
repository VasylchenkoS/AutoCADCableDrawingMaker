Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko.TerminalModules
    Module TerminalClassMapping

        Public Sub CreateTerminalBlock(strLocation As String, strTagstrip As String,
                                    eOrientation As TerminalOrientationEnum, eDucktSide As DuctSideEnum)

            Dim objTermsArray As List(Of String)
            Dim objMappingTermsCollection As List(Of TerminalAccessoriesClass)
            Dim dblScale As Double = 0.6

            objTermsArray = GetAllTermsInLocation(strLocation, strTagstrip)
            objMappingTermsCollection = FillTerminalData(objTermsArray, strTagstrip, eDucktSide)

            objMappingTermsCollection.Sort(Function(x As TerminalClass, y As TerminalClass)
                                               Return x.TERM.CompareTo(y.TERM)
                                           End Function)

            Dim ufTTS As New ufTerminalTypeSelector
            With ufTTS
                ufTTS.ShowDialog()
                If .rbtnSignalisation.Checked Then
                    AddPhoenixAccForSignalisation(objMappingTermsCollection, eDucktSide)
                ElseIf .rbtnMeasurement.Checked Then
                ElseIf .rbtnControl.Checked Then
                ElseIf .rbtnPower.Checked Then
                End If
            End With

            'Тут нужно добавить логику установки перемычек

            Dim acPromtPntOpt As New PromptPointOptions("Выберите точку вставки")
            Dim acPromtPntResult As PromptPointResult = Application.DocumentManager.MdiActiveDocument.Editor.GetPoint(acPromtPntOpt)

            If acPromtPntResult.Status <> PromptStatus.OK Then
                MsgBox("fail to get the insert point")
                Exit Sub
            End If

            Dim acInsertPt As Point3d = acPromtPntResult.Value

            CheckLayers()

            For Each objCurItem As Object In objMappingTermsCollection
                If TypeOf (objCurItem) Is TerminalClass Then
                    DrawTerminalBlock.DrawTerminalBlock(objCurItem, acInsertPt, dblScale)
                ElseIf TypeOf (objCurItem) Is TerminalAccessoriesClass Then
                    DrawTerminalBlock.DrawAccesoriesBlock(objCurItem, acInsertPt, dblScale)
                End If
                acInsertPt = New Point3d(acInsertPt.X, acInsertPt.Y - (objCurItem.HEIGHT * dblScale), acInsertPt.Z)
            Next

        End Sub

        Private Function FillTerminalData(objInputList As List(Of String), strTagstrip As String, eDucktSide As DuctSideEnum) As List(Of TerminalAccessoriesClass)
            Dim objResultArray As New List(Of TerminalAccessoriesClass)

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
