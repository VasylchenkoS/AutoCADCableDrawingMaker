Imports AutoCADCableDrawingMaker.com.vasilchenko.Classes
Imports AutoCADCableDrawingMaker.com.vasilchenko.Database
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.Modules
    Module CableClassMapping
        Friend Sub CreateTerminalBlock(strLocation As String)
            Dim acCableList As SortedList(Of Integer, CableClass) = DataAccessObject.FillCableData(strLocation)
            Dim acInsertionPoint As Point3d
            For Each acCable In acCableList.Values
                acInsertionPoint = GetBasePoint(acCable, acInsertionPoint)
                acCable.InsertionPoint = acInsertionPoint
                DrawCable(acCable)
            Next
        End Sub
        Private Function GetBasePoint(acCable As CableClass, acBasePoint As Point3d) As Point3d
            If acBasePoint = New Point3d Then
                If acCable.Orientation = Enums.OrientationEnum.Horisontal Then
                    If acCable.TerminalConnections(0).CatalogName.Equals("UT 2,5") Then
                        acBasePoint = New Point3d(acCable.TerminalConnections(0).ConnectionPoint.X - 25, acCable.TerminalConnections(0).ConnectionPoint.Y - 45 + 0.44, acCable.TerminalConnections(0).ConnectionPoint.Z)
                    ElseIf acCable.TerminalConnections(0).CatalogName.Equals("UT 2,5-MT") Then
                        acBasePoint = New Point3d(acCable.TerminalConnections(0).ConnectionPoint.X - 25, acCable.TerminalConnections(0).ConnectionPoint.Y - 45 - 2.34, acCable.TerminalConnections(0).ConnectionPoint.Z)
                    Else
                        acBasePoint = New Point3d(acCable.TerminalConnections(0).ConnectionPoint.X - 25, acCable.TerminalConnections(0).ConnectionPoint.Y - 45, acCable.TerminalConnections(0).ConnectionPoint.Z)
                    End If
                Else
                        acBasePoint = New Point3d(acCable.TerminalConnections(0).ConnectionPoint.X - 75, acCable.TerminalConnections(0).ConnectionPoint.Y + 55, acCable.TerminalConnections(0).ConnectionPoint.Z)
                End If
            Else
                acBasePoint = New Point3d(acBasePoint.X, acBasePoint.Y - 5, acBasePoint.Z)
            End If
            Return acBasePoint
        End Function
        Private Sub DrawCable(acCable As CableClass)
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor
            Dim acInsObjectId As ObjectId = Nothing


            Using acTransaction = acDatabase.TransactionManager.StartTransaction()
                Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(acCable.BlockPath)
                Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)

                If acBlockTable.Has(strBlkName) Then
                    Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                    acInsObjectId = acCurrBlkTblRcd.Id
                Else
                    Try
                        Dim acNewDbDwg As New Autodesk.AutoCAD.DatabaseServices.Database(False, True)
                        acNewDbDwg.ReadDwgFile(acCable.BlockPath, FileOpenMode.OpenTryForReadShare, True, "")
                        acInsObjectId = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                        acNewDbDwg.Dispose()
                    Catch ex As Exception
                        Throw New ArgumentNullException(String.Format("Не найден графический образ с именем {0}", strBlkName))
                    End Try
                End If

                Using acBlkRef As New BlockReference(acCable.InsertionPoint, acInsObjectId)
                    acBlkRef.Layer = "PSYMS"
                    acBlkRef.ScaleFactors = New Scale3d(1.0)

                    Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    acBlockTableRecord.AppendEntity(acBlkRef)
                    acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                    Dim acBlockTableAttrRec As BlockTableRecord = acTransaction.GetObject(acInsObjectId, OpenMode.ForRead)
                    Dim acAttrObjectId As ObjectId

                    For Each acAttrObjectId In acBlockTableAttrRec
                        Dim acAttrEntity As Entity = acTransaction.GetObject(acAttrObjectId, OpenMode.ForRead)
                        Dim acAttrDefinition = TryCast(acAttrEntity, AttributeDefinition)
                        If (acAttrDefinition IsNot Nothing) Then
                            Using acAttrReference As New AttributeReference()
                                Dim strTermdesc As String = ""
                                acAttrReference.SetAttributeFromBlock(acAttrDefinition, acBlkRef.BlockTransform)
                                Select Case acAttrReference.Tag
                                    Case "P_TAG1"
                                        acAttrReference.TextString = acCable.Tag
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PTAG"
                                    Case "INST"
                                        acAttrReference.TextString = acCable.Instance
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PLOC"
                                    Case "LOC"
                                        acAttrReference.TextString = acCable.Location
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PLOC"
                                    Case "MFG"
                                        acAttrReference.TextString = acCable.Manufacture
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PMFG"
                                    Case "CAT"
                                        acAttrReference.TextString = acCable.CatalogName
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PCAT"
                                    Case "DESC1"
                                        acAttrReference.TextString = acCable.Desc1
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PDESC"
                                    Case "DESC2"
                                        acAttrReference.TextString = acCable.Desc2
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PDESC"
                                    Case "DESC3"
                                        acAttrReference.TextString = acCable.Desc3
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PDESC"
                                    Case "WDBLKNAM"
                                        acAttrReference.TextString = acCable.Family
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "0"
                                    Case "CNT"
                                        acAttrReference.TextString = acCable.Count
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "0"
                                    Case Else
                                        If acAttrReference.Tag.StartsWith("WIRENO") Then
                                            Dim acConnection = acCable.TerminalConnections(AdditionalFunctions.GetLastNumericFromString(acAttrReference.Tag) - 1)
                                            acAttrReference.TextString = acConnection.Wireno
                                            acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                            acAttrReference.Layer = "PWIRE"
                                        End If
                                End Select
                                acBlkRef.AttributeCollection.AppendAttribute(acAttrReference)
                                acTransaction.AddNewlyCreatedDBObject(acAttrReference, True)
                            End Using
                        End If
                    Next

                    If acBlkRef.IsDynamicBlock Then
                        For Each acDynProp As DynamicBlockReferenceProperty In acBlkRef.DynamicBlockReferencePropertyCollection
                            Dim acConnection As TerminalConnectionClass
                            If acDynProp.PropertyName.StartsWith("CAB") Then
                                acConnection = acCable.TerminalConnections(AdditionalFunctions.GetLastNumericFromString(acDynProp.PropertyName) - 1)
                                acDynProp.Value = Math.Abs(acConnection.ConnectionPoint.X - acCable.InsertionPoint.X)
                            ElseIf acDynProp.PropertyName.StartsWith("WIRE") Then
                                acConnection = acCable.TerminalConnections(AdditionalFunctions.GetLastNumericFromString(acDynProp.PropertyName) - 1)
                                acDynProp.Value = Math.Abs(acConnection.ConnectionPoint.Y - acCable.InsertionPoint.Y)
                            End If
                        Next
                    End If
                End Using
                acTransaction.Commit()
            End Using
        End Sub
    End Module
End Namespace

