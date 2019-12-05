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

            Const dblScale As Double = 1
            Const dblRotation As Double = 0
            Dim objTerminalStripList As New TerminalStripClass

            Dim objTermsList As List(Of Short) = DataAccessObject.GetAllTermsInLocation(strLocation, strTagstrip)
            objTerminalStripList.TerminalList = FillTerminalData(objTermsList, strTagstrip, eDucktSide, strLocation)

            objTerminalStripList.TerminalList.Sort(Function(x As TerminalClass, y As TerminalClass)
                Return x.MainTermNumber.CompareTo(y.MainTermNumber)
            End Function)

            AddJumpers.FillStripByJumpers(objTerminalStripList)

            Dim ufTerminalTypeSelector As New ufTerminalTypeSelector
            With ufTerminalTypeSelector
                ufTerminalTypeSelector.ShowDialog()
                If .rbtnSignalisation.Checked Then
                    AddTerminalAccessories.AddAccessoriesForSignalisation(objTerminalStripList.TerminalList)
                ElseIf .rbtnMeasurement.Checked Then
                    AddTerminalAccessories.AddAccessoriesForMeasurement(objTerminalStripList.TerminalList)
                ElseIf .rbtnControl.Checked Then
                    AddTerminalAccessories.AddAccessoriesForControl(objTerminalStripList.TerminalList)
                ElseIf .rbtnPower.Checked Then
                    AddTerminalAccessories.AddAccessoriesForPower(objTerminalStripList.TerminalList)
                ElseIf .rbtnMetering.Checked Then
                    AddTerminalAccessories.AddAccessoriesForMetering(objTerminalStripList.TerminalList)
                Else
                    Exit Sub
                End If
            End With

            Dim acPromptPntOpt As New PromptPointOptions("Выберите точку вставки")
            Dim acPromptPntResult As PromptPointResult = Application.DocumentManager.MdiActiveDocument.Editor.GetPoint(acPromptPntOpt)

            If acPromptPntResult.Status <> PromptStatus.OK Then
                MsgBox("fail to get the insert point")
                Exit Sub
            End If

            Dim acInsertPt As Point3d = acPromptPntResult.Value

            DrawLayerChecker.CheckLayers()

            Dim objJumper As JumperClass

            Dim acDatabase As Database = HostApplicationServices.WorkingDatabase()
            Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
                For Each objCurItem As TerminalClass In objTerminalStripList.TerminalList
                    If not objCurItem Is nothing Then
                        If objCurItem.MainTermNumber <> 0
                            DrawBlocks.DrawTerminalBlock(acDatabase, acTransaction, objCurItem, acInsertPt, dblScale, dblRotation)

                            objJumper = objTerminalStripList.JumperList.Find(Function(x) x.StartTermNum = objCurItem.BottomLevelTerminal.TerminalNumber)
                            If objJumper IsNot Nothing Then
                                Dim acJumperInsertPt As Point3d = CalculateJumperInsertPoint(dblScale, acInsertPt, objJumper, objCurItem)
                                DrawBlocks.DrawJumpers(acDatabase, acTransaction, objJumper, acJumperInsertPt, dblScale)
                            End If
                        Else
                            DrawBlocks.DrawAccessorisesBlock(acDatabase, acTransaction, objCurItem, acInsertPt, dblScale)
                        End If
                        acInsertPt = New Point3d(acInsertPt.X, acInsertPt.Y - (objCurItem.Height*dblScale), acInsertPt.Z)
                    end if
                Next
                acTransaction.Commit()
            End Using
        End Sub

        Private Function FillTerminalData(objInputList As IReadOnlyCollection(Of Short), strTagstrip As String, eDucktSide As SideEnum, strLocation As string) As List(Of TerminalClass)
            Dim objResultArray As List(Of TerminalClass)
            objResultArray = (From shtTermNum In objInputList
                Select DataAccessObject.FillTermTypeData(strTagstrip, shtTermNum, strLocation)
                ).Cast (Of TerminalClass)().ToList()
            DataAccessObject.FillTerminalBlockPath(objResultArray)
            DataAccessObject.FillTerminalConnectionsData(objResultArray, eDucktSide)
            UnionTerminals(objResultArray)

            Return objResultArray
        End Function

        Private Sub UnionTerminals(ByRef objInputArray As List(Of TerminalClass))
            Dim objResultArray As List(Of TerminalClass)
            Dim objTopTermList = objInputArray.FindAll(Function(x) Not IsNothing(x.TopLevelTerminal)).ToList()
            If objTopTermList.Count <> 0 Then
                objResultArray = objInputArray.FindAll(Function(x) IsNothing(x.TopLevelTerminal)).ToList()
                For Each curTerminal As TerminalClass In objResultArray
                    curTerminal.TopLevelTerminal = objTopTermList.Find(Function(y)
                        Return y.LinkTerm = curTerminal.LinkTerm 
                    End Function).TopLevelTerminal
                Next
                objInputArray = objResultArray
            End If
        End Sub

        Private Function CalculateJumperInsertPoint(dblScale As Double, acInsertPt As Point3d, objJumper As JumperClass, objCurItem As TerminalClass) As Point3d
            Dim acJumperInsertPt As Point3d
            If objCurItem.Catalog.StartsWith("UT ") OrElse objCurItem.Catalog.StartsWith("ST ") And Not (objCurItem.Catalog.Contains("-MT")) Then
                If objJumper.Side = SideEnum.Left Then
                    acJumperInsertPt = New Point3d(acInsertPt.X - (2.5*dblScale), acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                Else
                    acJumperInsertPt = New Point3d(acInsertPt.X + (2.5*dblScale), acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                End If
            ElseIf objCurItem.Catalog.Equals("UT 2,5-MT") Or objCurItem.Catalog.Equals("UT 4-MT") Or objCurItem.Catalog.Equals("UT 4-MTD-DIO/R-L") Then
                If objJumper.Side = SideEnum.Left Then
                    acJumperInsertPt = New Point3d(acInsertPt.X + (2.5*dblScale), acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                Else
                    acJumperInsertPt = New Point3d(acInsertPt.X + (7.5*dblScale), acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                End If
            ElseIf objCurItem.Catalog.StartsWith("UTTB 2,5") OrElse objCurItem.Catalog.StartsWith("STTB 2,5") Then
                If objJumper.Side = SideEnum.Left Then
                    acJumperInsertPt = New Point3d(acInsertPt.X + 15, acInsertPt.Y - (6.425*dblScale), acInsertPt.Z)
                Else
                    acJumperInsertPt = New Point3d(acInsertPt.X + 20, acInsertPt.Y - (6.425*dblScale), acInsertPt.Z)
                End If
            ElseIf objCurItem.Catalog = "URTK 6" Then
                If objJumper.Jumper.Catalog.StartsWith("SB") Then
                    acJumperInsertPt = New Point3d(acInsertPt.X - 25, acInsertPt.Y - (4.075*dblScale), acInsertPt.Z)
                ElseIf objJumper.Jumper.Catalog.StartsWith("FBRI") Then
                    acJumperInsertPt = New Point3d(acInsertPt.X + 15, acInsertPt.Y - (4.075*dblScale), acInsertPt.Z)
                End If
                
            ElseIf objCurItem.Catalog.StartsWith("AVK 2,5 R") Then
                If objJumper.Side = SideEnum.Left Then
                    acJumperInsertPt = New Point3d(acInsertPt.X - 10, acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                Else
                    acJumperInsertPt = New Point3d(acInsertPt.X - 5, acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                End If
            ElseIf objCurItem.Catalog.StartsWith("AVK 4 A") or objCurItem.Catalog.StartsWith("AVK 2,5 A") Then
                acJumperInsertPt = New Point3d(acInsertPt.X + 10, acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
            ElseIf objCurItem.Catalog.StartsWith("AVK 6") Then
                acJumperInsertPt = New Point3d(acInsertPt.X, acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
            ElseIf objCurItem.Catalog.StartsWith("WGO 4") Then
                If objJumper.Jumper.Catalog.StartsWith("UK") Then
                    acJumperInsertPt = New Point3d(acInsertPt.X + 15, acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                ElseIf objJumper.Jumper.Catalog.StartsWith("IZUK") Then
                    acJumperInsertPt = New Point3d(acInsertPt.X - 25, acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                End If

            Else If objCurItem.Catalog.StartsWith("WDU 2.5") Then
                If objJumper.Side = SideEnum.Left Then
                    acJumperInsertPt = New Point3d(acInsertPt.X - (8.5*dblScale), acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                Else
                    acJumperInsertPt = New Point3d(acInsertPt.X + (5*dblScale), acInsertPt.Y - ((objCurItem.Height/2)*dblScale), acInsertPt.Z)
                End If
            End If

            Return acJumperInsertPt
        End Function
    End Module
End Namespace