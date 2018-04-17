Imports System.IO
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.TerminalModules

    Module DrawTerminalBlock
        Public Sub DrawTerminalBlock(objInputTerminal As TerminalClass, acInsertionPoint As Point3d,
                        eOrientation As TerminalOrientationEnum, eDucktSide As DuctSideEnum)

            CheckLayers()

            Dim acDatabase As Database = HostApplicationServices.WorkingDatabase()
            Dim acEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor

            Dim opPt As New PromptPointOptions("pick the insert point of the blk")
            Dim resPt As PromptPointResult = acEditor.GetPoint(opPt)

            If resPt.Status <> PromptStatus.OK Then
                MsgBox("fail to get the insert point")
                Exit Sub
            End If

            Dim acInsertPt As Point3d = resPt.Value


            Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
            Try
                Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
                Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim acInsObjectID As ObjectId
                If acBlockTable.Has(objInputTerminal.CAT) Then
                    Dim acCurrBlkTblRcd As BlockTableRecord
                    acCurrBlkTblRcd = acTransaction.GetObject(acBlockTable.Item(objInputTerminal.CAT), OpenMode.ForRead)
                    acInsObjectID = acCurrBlkTblRcd.Id
                Else
                    Dim acNewDbDwg As New Database(False, True)
                    acNewDbDwg.ReadDwgFile(objInputTerminal.BLOCK, FileShare.Read, True, "")
                    Dim s As String = objInputTerminal.CAT
                    'acInsObjectID = acDatabase.Insert(s, acNewDbDwg, True)
                    acInsObjectID = acDatabase.Insert("blckName", acNewDbDwg, True)
                    acNewDbDwg.Dispose()
                End If

                Dim acBlkRef As New BlockReference(acInsertPt, acInsObjectID)
                acBlkRef.Layer = "PSYMS"
                acBlkRef.ScaleFactors = New Scale3d(0.6)

                acBlockTableRecord.AppendEntity(acBlkRef)
                acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                ' modify the attribute of the block reference
                Dim acBlkTblAttrbRec As BlockTableRecord = acTransaction.GetObject(acInsObjectID, OpenMode.ForRead)
                Dim acAttrbObjectID As ObjectId
                For Each acAttrbObjectID In acBlkTblAttrbRec
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
                            Case "LOC"
                                acAttrbReference.TextString = objInputTerminal.LOC
                            Case "MFG"
                                acAttrbReference.TextString = objInputTerminal.MFG
                            Case "CAT"
                                acAttrbReference.TextString = objInputTerminal.CAT
                            Case "TERM"
                                acAttrbReference.TextString = objInputTerminal.TERM
                            Case "WIRENOL"
                                If objInputTerminal.WiresLeftList.Count <> 0 Then acAttrbReference.TextString = objInputTerminal.WiresLeftList.Item(0).Wireno
                            Case "WIRENOR"
                                If objInputTerminal.WiresRigthList.Count <> 0 Then acAttrbReference.TextString = objInputTerminal.WiresRigthList.Item(0).Wireno
                            Case "TERMDESCL"
                                If objInputTerminal.WiresLeftList.Count <> 0 AndAlso Not CableInList(objInputTerminal.WiresLeftList) Then
                                    For lngA = 1 To objInputTerminal.WiresLeftList.Count
                                        strTermdesc = strTermdesc & objInputTerminal.WiresLeftList.Item(lngA - 1).TERMDESC
                                        If lngA <> objInputTerminal.WiresLeftList.Count Then
                                            strTermdesc = strTermdesc & ", "
                                        End If
                                    Next
                                End If
                                acAttrbReference.TextString = strTermdesc
                            Case "TERMDESCR"
                                If objInputTerminal.WiresRigthList.Count <> 0 AndAlso Not CableInList(objInputTerminal.WiresRigthList) Then
                                    For lngA = 1 To objInputTerminal.WiresRigthList.Count
                                        strTermdesc = strTermdesc & objInputTerminal.WiresRigthList.Item(lngA - 1).TERMDESC
                                        If lngA <> objInputTerminal.WiresRigthList.Count Then
                                            strTermdesc = acAttrbReference.TextString & ", "
                                        End If
                                    Next
                                End If
                                acAttrbReference.TextString = strTermdesc
                        End Select

                        'Dim ptBase As New Point3d(blkRef.Position.X + attDef.Position.X, blkRef.Position.Y + attDef.Position.Y, blkRef.Position.Z + attDef.Position.Z)
                        'attRef.Position = ptBase
                        'attRef.Rotation = attDef.Rotation
                        'attRef.Tag = attDef.Tag
                        'attRef.TextString = "input your data here"
                        'attRef.Height = attDef.Height
                        'attRef.FieldLength = attDef.FieldLength


                        acBlkRef.AttributeCollection.AppendAttribute(acAttrbReference)
                        acTransaction.AddNewlyCreatedDBObject(acAttrbReference, True)
                    End If
                Next
                acTransaction.Commit()

            Catch
                MsgBox(Err.Description)
            Finally
                acTransaction.Dispose()
            End Try


            'Dim acDatabase As Database = Application.DocumentManager.MdiActiveDocument.Database

            'Using acNewDatabase As Database = New Database(False, True)
            '    acNewDatabase.ReadDwgFile(objInputTerminal.BLOCK, FileOpenMode.OpenForReadAndAllShare, True, "")

            '    Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
            '        Try
            '            Dim acCurDwg As ObjectId = acDatabase.Insert(objInputTerminal.CAT, acNewDatabase, True)
            '            If IsNothing(acCurDwg) Then
            '                Throw New NullReferenceException
            '            End If

            '            Dim acBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acDatabase.CurrentSpaceId, OpenMode.ForWrite)
            '            Dim acBlkRfrnc As BlockReference = New BlockReference(New Point3d(), acCurDwg)
            '            acBlkTblRcd.AppendEntity(acBlkRfrnc)
            '            acTransaction.AddNewlyCreatedDBObject(acBlkRfrnc, True)

            '            For Each acId As ObjectId In 

            '            acTransaction.Commit()
            '        Catch
            '            MsgBox(Err.Description)
            '        Finally
            '            acTransaction.Dispose()
            '        End Try
            '    End Using
            'End Using

        End Sub

        Private Function Getlayer(strLayerName As String) As LayerTableRecord

        End Function
    End Module
End Namespace

