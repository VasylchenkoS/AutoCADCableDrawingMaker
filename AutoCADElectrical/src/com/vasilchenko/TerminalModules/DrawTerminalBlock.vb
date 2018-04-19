Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.TerminalModules

    Module DrawTerminalBlock
        Public Sub DrawTerminalBlock(ByRef objInputTerminal As TerminalClass, acInsertPt As Point3d, dblScale As Double)

            Dim acDatabase As Database = HostApplicationServices.WorkingDatabase()
            Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(objInputTerminal.BLOCK)

            Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
            Try
                Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
                Dim acInsObjectID As ObjectId

                If acBlockTable.Has(strBlkName) Then
                    Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                    acInsObjectID = acCurrBlkTblRcd.Id
                Else
                    Dim acNewDbDwg As New Database(False, True)
                    acNewDbDwg.ReadDwgFile(objInputTerminal.BLOCK, FileOpenMode.OpenTryForReadShare, True, "")
                    acInsObjectID = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                    acNewDbDwg.Dispose()
                End If
                Using acBlkRef As New BlockReference(acInsertPt, acInsObjectID)
                    acBlkRef.Layer = "PSYMS"
                    acBlkRef.ScaleFactors = New Scale3d(dblScale)

                    Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    acBlockTableRecord.AppendEntity(acBlkRef)
                    acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                    ' modify the attribute of the block reference
                    Dim acBlockTableAttrbRec As BlockTableRecord = acTransaction.GetObject(acInsObjectID, OpenMode.ForRead)
                    Dim acAttrbObjectID As ObjectId
                    For Each acAttrbObjectID In acBlockTableAttrbRec
                        Dim acAttrbEntity As Entity = acTransaction.GetObject(acAttrbObjectID, OpenMode.ForRead)
                        If TypeOf acAttrbEntity Is AttributeDefinition Then
                            Dim acAttrbDefinition As AttributeDefinition = CType(acAttrbEntity, AttributeDefinition)
                            Dim acAttrbReference As New AttributeReference()
                            Dim strTermdesc As String = ""
                            acAttrbReference.SetAttributeFromBlock(acAttrbDefinition, acBlkRef.BlockTransform)
                            Select Case acAttrbReference.Tag
                                Case "P_TAGSTRIP"
                                    acAttrbReference.TextString = objInputTerminal.P_TAGSTRIP
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PTAG"
                                Case "INST"
                                    acAttrbReference.TextString = objInputTerminal.INST
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PLOC"
                                Case "LOC"
                                    acAttrbReference.TextString = objInputTerminal.LOC
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PLOC"
                                Case "MFG"
                                    acAttrbReference.TextString = objInputTerminal.MFG
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PMFG"
                                Case "CAT"
                                    acAttrbReference.TextString = objInputTerminal.CAT
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PCAT"
                                Case "TERM"
                                    acAttrbReference.TextString = objInputTerminal.TERM
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                    acAttrbReference.Layer = "PTERM"
                                Case "WIRENOL"
                                    If objInputTerminal.WiresLeftList.Count <> 0 Then acAttrbReference.TextString = objInputTerminal.WiresLeftList.Item(0).Wireno
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PWIRE"
                                Case "WIRENOR"
                                    If objInputTerminal.WiresRigthList.Count <> 0 Then acAttrbReference.TextString = objInputTerminal.WiresRigthList.Item(0).Wireno
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PWIRE"
                                Case "TERMDESCL"
                                    If objInputTerminal.WiresLeftList.Count <> 0 And CableInList(objInputTerminal.WiresLeftList) = -1 Then
                                        For lngA = 0 To objInputTerminal.WiresLeftList.Count - 1
                                            strTermdesc = strTermdesc & objInputTerminal.WiresLeftList.Item(lngA).TERMDESC & ", "
                                        Next
                                        strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
                                    End If
                                    acAttrbReference.TextString = strTermdesc
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                    acAttrbReference.Layer = "PDESC"
                                Case "TERMDESCR"
                                    If objInputTerminal.WiresRigthList.Count <> 0 And CableInList(objInputTerminal.WiresRigthList) = -1 Then
                                        For lngA = 0 To objInputTerminal.WiresRigthList.Count - 1
                                            strTermdesc = strTermdesc & objInputTerminal.WiresRigthList.Item(lngA).TERMDESC & ", "
                                        Next
                                        strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
                                    End If
                                    acAttrbReference.TextString = strTermdesc
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                    acAttrbReference.Layer = "PDESC"
                                Case "WIDTH"
                                    objInputTerminal.WIDTH = CDbl(acAttrbReference.TextString)
                                Case "HEIGHT"
                                    objInputTerminal.HEIGHT = CDbl(acAttrbReference.TextString)
                            End Select

                            acBlkRef.AttributeCollection.AppendAttribute(acAttrbReference)
                            acTransaction.AddNewlyCreatedDBObject(acAttrbReference, True)
                        End If
                    Next
                End Using

                DrawLines(acTransaction, objInputTerminal, acDatabase, acInsertPt, dblScale)

                acTransaction.Commit()
            Catch
                MsgBox(Err.Description)
            Finally
                acTransaction.Dispose()
            End Try
        End Sub

        Friend Sub DrawAccesoriesBlock(objInputTerminal As TerminalAccessoriesClass, acInsertPt As Point3d, dblScale As Double)

            Dim acDatabase As Database = HostApplicationServices.WorkingDatabase()
            Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(objInputTerminal.BLOCK)

            Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
            Try
                Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
                Dim acInsObjectID As ObjectId

                If acBlockTable.Has(strBlkName) Then
                    Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                    acInsObjectID = acCurrBlkTblRcd.Id
                Else
                    Dim acNewDbDwg As New Database(False, True)
                    acNewDbDwg.ReadDwgFile(objInputTerminal.BLOCK, FileOpenMode.OpenTryForReadShare, True, "")
                    acInsObjectID = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                    acNewDbDwg.Dispose()
                End If
                Using acBlkRef As New BlockReference(acInsertPt, acInsObjectID)
                    acBlkRef.Layer = "PSYMS"
                    acBlkRef.ScaleFactors = New Scale3d(dblScale)

                    Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    acBlockTableRecord.AppendEntity(acBlkRef)
                    acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                    ' modify the attribute of the block reference
                    Dim acBlockTableAttrbRec As BlockTableRecord = acTransaction.GetObject(acInsObjectID, OpenMode.ForRead)
                    Dim acAttrbObjectID As ObjectId
                    For Each acAttrbObjectID In acBlockTableAttrbRec
                        Dim acAttrbEntity As Entity = acTransaction.GetObject(acAttrbObjectID, OpenMode.ForRead)
                        If TypeOf acAttrbEntity Is AttributeDefinition Then
                            Dim acAttrbDefinition As AttributeDefinition = CType(acAttrbEntity, AttributeDefinition)
                            Dim acAttrbReference As New AttributeReference()
                            Dim strTermdesc As String = ""
                            acAttrbReference.SetAttributeFromBlock(acAttrbDefinition, acBlkRef.BlockTransform)
                            Select Case acAttrbReference.Tag
                                Case "P_TAGSTRIP"
                                    acAttrbReference.TextString = objInputTerminal.P_TAGSTRIP
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PTAG"
                                Case "INST"
                                    acAttrbReference.TextString = objInputTerminal.INST
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PLOC"
                                Case "LOC"
                                    acAttrbReference.TextString = objInputTerminal.LOC
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PLOC"
                                Case "MFG"
                                    acAttrbReference.TextString = objInputTerminal.MFG
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PMFG"
                                Case "CAT"
                                    acAttrbReference.TextString = objInputTerminal.CAT
                                    acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                    acAttrbReference.Layer = "PCAT"
                                Case "HEIGHT"
                                    objInputTerminal.HEIGHT = CDbl(acAttrbReference.TextString)
                            End Select

                            acBlkRef.AttributeCollection.AppendAttribute(acAttrbReference)
                            acTransaction.AddNewlyCreatedDBObject(acAttrbReference, True)
                        End If
                    Next
                End Using
                acTransaction.Commit()
            Catch
                MsgBox(Err.Description)
            Finally
                acTransaction.Dispose()
            End Try
        End Sub

        Private Sub DrawLines(acTransaction As Transaction, objInputTerminal As TerminalClass, acDatabase As Database, acInsertPt As Point3d, dblScale As Double)
            Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
            Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Dim acStartPtY As Double = acInsertPt.Y - objInputTerminal.HEIGHT * dblScale / 2

            If objInputTerminal.WiresLeftList.Count <> 0 And CableInList(objInputTerminal.WiresLeftList) = -1 Then
                Dim acStartLPtX As Double = acInsertPt.X - objInputTerminal.WIDTH * dblScale / 2
                Using acLine As Line = New Line(New Point3d(acStartLPtX, acStartPtY, 0),
                                            New Point3d(acStartLPtX - 9, acStartPtY, 0))
                    acBlockTableRecord.AppendEntity(acLine)
                    acTransaction.AddNewlyCreatedDBObject(acLine, True)
                End Using
            End If

            If objInputTerminal.WiresRigthList.Count <> 0 And CableInList(objInputTerminal.WiresLeftList) = -1 Then
                Dim acStartRPtX As Double = acInsertPt.X + objInputTerminal.WIDTH * dblScale / 2
                Using acLine As Line = New Line(New Point3d(acStartRPtX, acStartPtY, 0),
                                            New Point3d(acStartRPtX + 9, acStartPtY, 0))
                    acBlockTableRecord.AppendEntity(acLine)
                    acTransaction.AddNewlyCreatedDBObject(acLine, True)
                End Using
            End If
        End Sub

    End Module
End Namespace

