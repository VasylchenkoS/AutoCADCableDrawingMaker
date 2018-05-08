Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices

Namespace com.vasilchenko.TerminalModules
    Module TerminalClassMapping

        Public Sub CreateTerminalBlock(strLocation As String, strTagstrip As String,
                                    eOrientation As OrientationEnum, eDucktSide As SideEnum)

            Dim objTermsList As List(Of String)
            Dim dblScale As Double = 0.6
            Dim dblRotation As Double
            Dim objTerminalStripList As New TerminalStripClass

            objTermsList = DBDataAccessObject.GetAllTermsInLocation(strLocation, strTagstrip)
            objTerminalStripList.TerminalList = FillTerminalData(objTermsList, strTagstrip, eDucktSide)

            objTerminalStripList.TerminalList.Sort(Function(x As TerminalClass, y As TerminalClass)
                                                       Return x.TERM.CompareTo(y.TERM)
                                                   End Function)

            AddPhoenixJumpers.FillStripByJumpers(objTerminalStripList)

            Dim ufTTS As New ufTerminalTypeSelector
            With ufTTS
                ufTTS.ShowDialog()
                If .rbtnSignalisation.Checked Then
                    AddPhoenixTerminalAccessories.AddPhoenixAccForSignalisation(objTerminalStripList.TerminalList)
                ElseIf .rbtnMeasurement.Checked Then
                    AddPhoenixTerminalAccessories.AddPhoenixAccForMeasurement(objTerminalStripList.TerminalList)
                ElseIf .rbtnControl.Checked Then
                    AddPhoenixTerminalAccessories.AddPhoenixAccForControl(objTerminalStripList.TerminalList)
                ElseIf .rbtnPower.Checked Then
                Else
                    Exit Sub
                End If
            End With

            Dim acPromtPntOpt As New PromptPointOptions("Выберите точку вставки")
            Dim acPromtPntResult As PromptPointResult = Application.DocumentManager.MdiActiveDocument.Editor.GetPoint(acPromtPntOpt)

            If acPromtPntResult.Status <> PromptStatus.OK Then
                MsgBox("fail to get the insert point")
                Exit Sub
            End If

            Dim acInsertPt As Point3d = acPromtPntResult.Value

            DrawLayerChecker.CheckLayers()

            Dim objJumper As JumperClass

            'Добавить логику прорисовки джамперов!
            Dim acDatabase As Database = HostApplicationServices.WorkingDatabase()
            Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
            Try
                For Each objCurItem As Object In objTerminalStripList.TerminalList
                    If TypeOf (objCurItem) Is TerminalClass Then
                        If objCurItem.CAT = "UT-2,5-MT" And eDucktSide = SideEnum.Left Then
                            dblRotation = 180.0
                        Else dblRotation = 0.0
                        End If
                        DrawBlocks.DrawTerminalBlock(acDatabase, acTransaction, objCurItem, acInsertPt, dblScale, dblRotation)
                        objJumper = objTerminalStripList.JumperList.Find(Function(x) x.StartTermNum = objCurItem.TERM)
                        If objJumper IsNot Nothing Then
                            Dim acJumperInsertPt As Point3d = CalculateJumperInsertPoint(dblScale, acInsertPt, objJumper, objCurItem, dblRotation)
                            DrawBlocks.DrawJumpers(acDatabase, acTransaction, objJumper, acJumperInsertPt, dblScale)
                        End If
                    ElseIf TypeOf (objCurItem) Is TerminalAccessoriesClass Then
                        DrawBlocks.DrawAccesoriesBlock(acDatabase, acTransaction, objCurItem, acInsertPt, dblScale)
                    End If
                    acInsertPt = New Point3d(acInsertPt.X, acInsertPt.Y - (objCurItem.HEIGHT * dblScale), acInsertPt.Z)
                Next
                acTransaction.Commit()
            Catch
                MsgBox(Err.Description)
            Finally
                acTransaction.Dispose()
            End Try

        End Sub

        Private Function CalculateJumperInsertPoint(dblScale As Double, acInsertPt As Point3d, objJumper As JumperClass, objCurItem As Object, dblRotation As Double) As Point3d
            Dim acJumperInsertPt As Point3d
            If objCurItem.CAT = "UT-2,5" Then
                If objJumper.Side = SideEnum.Left Then
                    acJumperInsertPt = New Point3d(acInsertPt.X - (2.5 * dblScale), acInsertPt.Y - ((objCurItem.HEIGHT / 2) * dblScale), acInsertPt.Z)
                Else
                    acJumperInsertPt = New Point3d(acInsertPt.X + (2.5 * dblScale), acInsertPt.Y - ((objCurItem.HEIGHT / 2) * dblScale), acInsertPt.Z)
                End If
            ElseIf objCurItem.CAT = "UT-2,5-MT" Then
                If dblRotation = 0.0 Then
                    If objJumper.Side = SideEnum.Left Then
                        acJumperInsertPt = New Point3d(acInsertPt.X + (2.5 * dblScale), acInsertPt.Y - ((objCurItem.HEIGHT / 2) * dblScale), acInsertPt.Z)
                    Else
                        acJumperInsertPt = New Point3d(acInsertPt.X + (7.5 * dblScale), acInsertPt.Y - ((objCurItem.HEIGHT / 2) * dblScale), acInsertPt.Z)
                    End If
                Else
                    If objJumper.Side = SideEnum.Left Then
                        acJumperInsertPt = New Point3d(acInsertPt.X - (7.5 * dblScale), acInsertPt.Y - ((objCurItem.HEIGHT / 2) * dblScale), acInsertPt.Z)
                    Else
                        acJumperInsertPt = New Point3d(acInsertPt.X - (2.5 * dblScale), acInsertPt.Y - ((objCurItem.HEIGHT / 2) * dblScale), acInsertPt.Z)
                    End If
                End If
            End If
            Return acJumperInsertPt
        End Function

        Private Function FillTerminalData(objInputList As List(Of String), strTagstrip As String, eDucktSide As SideEnum) As List(Of TerminalAccessoriesClass)
            Dim objResultArray As New List(Of TerminalAccessoriesClass)

            For Each strTempItem As String In objInputList
                objResultArray.Add(DBDataAccessObject.FillTermTypeData(strTagstrip, strTempItem))
            Next
            DBDataAccessObject.FillTerminalBlockPath(objResultArray)
            DBDataAccessObject.FillTerminalConnectionsData(objResultArray, eDucktSide)

            'For Each objTempItem As TerminalClass In objResultArray
            '    DBDataAccessObject.FillTerminalBlockPath(objTempItem)
            '    DBDataAccessObject.FillTerminalConnectionsData(objTempItem, eDucktSide)
            'Next

            Return objResultArray

        End Function

    End Module
End Namespace
