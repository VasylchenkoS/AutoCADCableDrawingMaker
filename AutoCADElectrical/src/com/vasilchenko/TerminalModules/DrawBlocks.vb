Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.TerminalModules

    Module DrawBlocks
        Public Sub DrawTerminalBlock(acDatabase As Database, acTransaction As Transaction,
                                     ByRef objInputLevelTerminal As TerminalClass, acInsertPt As Point3d, dblScale As Double, Optional dblRotation As Double = 0.0)

            Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(objInputLevelTerminal.BlockPath)
            Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
            Dim acInsObjectId As ObjectId

            If acBlockTable.Has(strBlkName) Then
                Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                acInsObjectId = acCurrBlkTblRcd.Id
            Else
                Dim acNewDbDwg As New Database(False, True)
                acNewDbDwg.ReadDwgFile(objInputLevelTerminal.BlockPath, FileOpenMode.OpenTryForReadShare, True, "")
                acInsObjectId = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                acNewDbDwg.Dispose()
            End If

            Using acBlkRef As New BlockReference(acInsertPt, acInsObjectId)
                acBlkRef.Layer = "PSYMS"
                acBlkRef.ScaleFactors = New Scale3d(dblScale)

                Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                acBlockTableRecord.AppendEntity(acBlkRef)
                acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                Dim acBlockTableAttrbRec As BlockTableRecord = acTransaction.GetObject(acInsObjectId, OpenMode.ForRead)
                Dim acAttrbObjectId As ObjectId

                If objInputLevelTerminal.MainTermNumber = 31 Then
                    Console.Write("")
                End If

                For Each acAttrbObjectId In acBlockTableAttrbRec
                    Dim acAttrbEntity As Entity = acTransaction.GetObject(acAttrbObjectId, OpenMode.ForRead)
                    Dim acAttrbDefinition As AttributeDefinition = TryCast(acAttrbEntity, AttributeDefinition)
                    If (acAttrbDefinition IsNot Nothing) Then
                        Dim acAttrbReference As New AttributeReference()
                        Dim strTermdesc As String = ""
                        acAttrbReference.SetAttributeFromBlock(acAttrbDefinition, acBlkRef.BlockTransform)
                        Select Case acAttrbReference.Tag
                            Case "WIDTH"
                                objInputLevelTerminal.Width = CDbl(acAttrbReference.TextString)
                            Case "HEIGHT"
                                objInputLevelTerminal.Height = CDbl(acAttrbReference.TextString)
                        End Select
                    End If
                Next

                For Each acAttrbObjectId In acBlockTableAttrbRec
                    Dim acAttrbEntity As Entity = acTransaction.GetObject(acAttrbObjectId, OpenMode.ForRead)
                    Dim acAttrbDefinition As AttributeDefinition = TryCast(acAttrbEntity, AttributeDefinition)
                    If (acAttrbDefinition IsNot Nothing) Then
                        Dim acAttrbReference As New AttributeReference()
                        Dim strTermdesc As String = ""
                        acAttrbReference.SetAttributeFromBlock(acAttrbDefinition, acBlkRef.BlockTransform)
                        Select Case acAttrbReference.Tag
                            Case "P_TAGSTRIP"
                                acAttrbReference.TextString = objInputLevelTerminal.TagStrip
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PTAG"
                            Case "INST"
                                acAttrbReference.TextString = objInputLevelTerminal.Instance
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PLOC"
                            Case "LOC"
                                acAttrbReference.TextString = objInputLevelTerminal.Location
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PLOC"
                            Case "MFG"
                                acAttrbReference.TextString = objInputLevelTerminal.Manufacture
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PMFG"
                            Case "CAT"
                                acAttrbReference.TextString = objInputLevelTerminal.Catalog
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PCAT"
                            Case "CNT"
                                acAttrbReference.TextString = IIf(objInputLevelTerminal.Count = 0, 1.0, objInputLevelTerminal.Count)
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "0"
                            Case "TERM"
                                acAttrbReference.TextString = objInputLevelTerminal.BottomLevelTerminal.TerminalNumber
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "PTERM"
                            Case "TERMT"
                                acAttrbReference.TextString = objInputLevelTerminal.TopLevelTerminal.TerminalNumber
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "PTERM"
                            Case "WIRENOL"
                                If objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Count <> 0 AndAlso objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Item(0).WireNumber.ToLower <> "pe" Then
                                    acAttrbReference.TextString = objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Item(0).WireNumber
                                End If
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PWIRE"
                            Case "WIRENOR"
                                If objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Count <> 0 AndAlso objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Item(0).WireNumber.ToLower <> "pe" Then
                                    acAttrbReference.TextString = objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Item(0).WireNumber
                                End If
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PWIRE"
                            Case "WIRENOTL"
                                If objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Count <> 0 AndAlso objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Item(0).WireNumber.ToLower <> "pe" Then
                                    acAttrbReference.TextString = objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Item(0).WireNumber
                                End If
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PWIRE"
                            Case "WIRENOTR"
                                If objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Count <> 0 AndAlso objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Item(0).WireNumber.ToLower <> "pe" Then
                                    acAttrbReference.TextString = objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Item(0).WireNumber
                                End If
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PWIRE"
                            Case "TERMDESCL"
                                Dim isCable = False
                                If objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Count <> 0 Then
                                    If objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Any(Function(x As WireClass) x.HasCable) Then
                                        strTermdesc = "в " & objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Find(Function(x As WireClass) x.HasCable).Cable.Mark
                                        isCable = True
                                    Else
                                        For lngA = 0 To objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Count - 1
                                            If objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Item(lngA).Termdesc <> "" Then
                                                If strTermdesc = "" Then
                                                    strTermdesc = "к " & objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Item(lngA).Termdesc & ", "
                                                Else
                                                    strTermdesc = strTermdesc & objInputLevelTerminal.BottomLevelTerminal.WiresLeftListList.Item(lngA).Termdesc & ", "
                                                End If
                                            End If
                                        Next
                                        If strTermdesc <> "" Then
                                            strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
                                        End If
                                    End If

                                    acAttrbReference.TextString = strTermdesc

                                    If objInputLevelTerminal.Catalog.StartsWith("UT ") OrElse
                                       objInputLevelTerminal.Catalog.StartsWith("UT 6-HESI") OrElse
                                       objInputLevelTerminal.Catalog.StartsWith("AVK") OrElse
                                       objInputLevelTerminal.Catalog.StartsWith("WGO")Then
                                        Dim x = acInsertPt.X - objInputLevelTerminal.Width / 2
                                        Dim y = acInsertPt.Y - objInputLevelTerminal.Height / 2
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), -20, isCable)
                                    ElseIf objInputLevelTerminal.Catalog.Equals("UTTB 2,5-MT-P/P") Then
                                        Dim x = acInsertPt.X - 35
                                        Dim y = acInsertPt.Y - 6
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), -10, isCable)
                                    ElseIf objInputLevelTerminal.Catalog.Equals("URTK 6") Then
                                        Dim x = acInsertPt.X - objInputLevelTerminal.Width / 2
                                        Dim y = acInsertPt.Y - objInputLevelTerminal.Height / 2
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), -26, isCable)
                                    End If
                                End If

                                strTermdesc = ""
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "PDESC"
                            Case "TERMDESCR"
                                Dim isCable = False
                                If objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Count <> 0 Then
                                    If objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Any(Function(x As WireClass) x.HasCable) Then
                                        strTermdesc = "в " & objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Find(Function(x As WireClass) x.HasCable).Cable.Mark
                                        isCable = True
                                    Else
                                        For lngA = 0 To objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Count - 1
                                            If objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Item(lngA).Termdesc <> "" Then
                                                If strTermdesc = "" Then
                                                    strTermdesc = "к " & objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Item(lngA).Termdesc & ", "
                                                Else
                                                    strTermdesc = strTermdesc & objInputLevelTerminal.BottomLevelTerminal.WiresRigthListList.Item(lngA).Termdesc & ", "
                                                End If
                                            End If
                                        Next
                                        If strTermdesc <> "" Then strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
                                    End If
                                    acAttrbReference.TextString = strTermdesc
                                    If objInputLevelTerminal.Catalog.StartsWith("UT ") OrElse
                                       objInputLevelTerminal.Catalog.StartsWith("UT 6-HESI") OrElse
                                       objInputLevelTerminal.Catalog.StartsWith("AVK") OrElse
                                       objInputLevelTerminal.Catalog.StartsWith("WGO")Then
                                        Dim x = acInsertPt.X + objInputLevelTerminal.Width / 2
                                        Dim y = acInsertPt.Y - objInputLevelTerminal.Height / 2
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), 20, isCable)
                                    ElseIf objInputLevelTerminal.Catalog.Equals("UTTB 2,5-MT-P/P") Then
                                        Dim x = acInsertPt.X + 35
                                        Dim y = acInsertPt.Y - 6
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), 10, isCable)
                                    ElseIf objInputLevelTerminal.Catalog.Equals("URTK 6") Then
                                        Dim x = acInsertPt.X + objInputLevelTerminal.Width / 2
                                        Dim y = acInsertPt.Y - objInputLevelTerminal.Height / 2
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), 26, isCable)
                                    End If
                                End If
                                strTermdesc = ""
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "PDESC"
                            Case "TERMDESCTL"
                                Dim isCable = False
                                If objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Count <> 0 Then
                                    If objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Any(Function(x As WireClass) x.HasCable) Then
                                        strTermdesc = "в " & objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Find(Function(x As WireClass) x.HasCable).Cable.Mark
                                        isCable = True
                                    Else
                                        For lngA = 0 To objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Count - 1
                                            If objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Item(lngA).Termdesc <> "" Then
                                                If strTermdesc = "" Then
                                                    strTermdesc = "к " & objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Item(lngA).Termdesc & ", "
                                                Else
                                                    strTermdesc = strTermdesc & objInputLevelTerminal.TopLevelTerminal.WiresLeftListList.Item(lngA).Termdesc & ", "
                                                End If
                                            End If
                                        Next
                                        If strTermdesc <> "" Then strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
                                    End If
                                    acAttrbReference.TextString = strTermdesc
                                    If objInputLevelTerminal.Catalog.Equals("UTTB 2,5-MT-P/P") Then
                                        Dim x = acInsertPt.X - 20
                                        Dim y = acInsertPt.Y - 3.5
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), -25, isCable)
                                    End If
                                End If
                                strTermdesc = ""
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "PDESC"
                            Case "TERMDESCTR"
                                Dim isCable = False
                                If objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Count <> 0 Then
                                    If objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Any(Function(x As WireClass) x.HasCable) Then
                                        strTermdesc = "в " & objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Find(Function(x As WireClass) x.HasCable).Cable.Mark
                                        isCable = True
                                    Else
                                        For lngA = 0 To objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Count - 1
                                            If objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Item(lngA).Termdesc <> "" Then
                                                If strTermdesc = "" Then
                                                    strTermdesc = "к " & objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Item(lngA).Termdesc & ", "
                                                Else
                                                    strTermdesc = strTermdesc & objInputLevelTerminal.TopLevelTerminal.WiresRigthListList.Item(lngA).Termdesc & ", "
                                                End If
                                            End If
                                        Next
                                        If strTermdesc <> "" Then strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
                                    End If
                                    acAttrbReference.TextString = strTermdesc
                                    If objInputLevelTerminal.Catalog.Equals("UTTB 2,5-MT-P/P") Then
                                        Dim x = acInsertPt.X + 10
                                        Dim y = acInsertPt.Y - 3.5
                                        DrawLine(acDatabase, acTransaction, New Point3d(x, y, 0), 35, isCable)
                                    End If
                                End If
                                strTermdesc = ""
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "PDESC"
                        End Select

                        acBlkRef.AttributeCollection.AppendAttribute(acAttrbReference)
                        acTransaction.AddNewlyCreatedDBObject(acAttrbReference, True)

                        'If dblRotation = 180 Then
                        '    Dim acPtFrom As Point3d = acInsertPt
                        '    Dim acPtTo As Point3d = New Point3d(acInsertPt.X, acInsertPt.Y - 5, acInsertPt.Z)
                        '    Dim acLine3d As Line3d = New Line3d(acPtFrom, acPtTo)
                        '    acBlkRef.TransformBy(Matrix3d.Mirroring(acLine3d))
                        '    acAttrbReference.TransformBy(Matrix3d.Mirroring(acLine3d))
                        'End If

                    End If
                Next
            End Using

            'DrawLines(acDatabase, acTransaction, objInputLevelTerminal, acInsertPt, dblScale)

        End Sub

        Private Sub DrawLine(acDatabase As Database, acTransaction As Transaction, acInsertionPoint As Point3d, lenght As Integer, isCable As Boolean)
            Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
            Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Using acLine As Line = New Line(acInsertionPoint,
                                            New Point3d(acInsertionPoint.X + lenght, acInsertionPoint.Y, 0))
                If isCable Then acLine.Layer = "Cable"
                acBlockTableRecord.AppendEntity(acLine)
                acTransaction.AddNewlyCreatedDBObject(acLine, True)
            End Using
        End Sub

        Friend Sub DrawAccesoriesBlock(acDatabase As Database, acTransaction As Transaction,
                                       ByRef objInputTerminal As TerminalClass, acInsertPt As Point3d, dblScale As Double)

            Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(objInputTerminal.BlockPath)
            Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
            Dim acInsObjectId As ObjectId

            If acBlockTable.Has(strBlkName) Then
                Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                acInsObjectId = acCurrBlkTblRcd.Id
            Else
                Dim acNewDbDwg As New Database(False, True)
                acNewDbDwg.ReadDwgFile(objInputTerminal.BlockPath, FileOpenMode.OpenTryForReadShare, True, "")
                acInsObjectId = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                acNewDbDwg.Dispose()
            End If

            Using acBlkRef As New BlockReference(acInsertPt, acInsObjectId)
                acBlkRef.Layer = "PSYMS"
                acBlkRef.ScaleFactors = New Scale3d(dblScale)

                Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                acBlockTableRecord.AppendEntity(acBlkRef)
                acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                Dim acBlockTableAttrbRec As BlockTableRecord = acTransaction.GetObject(acInsObjectId, OpenMode.ForRead)
                Dim acAttrbObjectId As ObjectId
                For Each acAttrbObjectId In acBlockTableAttrbRec
                    Dim acAttrbEntity As Entity = acTransaction.GetObject(acAttrbObjectId, OpenMode.ForRead)
                    Dim acAttrbDefinition As AttributeDefinition = TryCast(acAttrbEntity, AttributeDefinition)
                    If (acAttrbDefinition IsNot Nothing) Then
                        Dim acAttrbReference As New AttributeReference()
                        acAttrbReference.SetAttributeFromBlock(acAttrbDefinition, acBlkRef.BlockTransform)
                        Select Case acAttrbReference.Tag
                            Case "P_TAGSTRIP"
                                acAttrbReference.TextString = objInputTerminal.TagStrip
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PTAG"
                            Case "INST"
                                acAttrbReference.TextString = objInputTerminal.Instance
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PLOC"
                            Case "LOC"
                                acAttrbReference.TextString = objInputTerminal.Location
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PLOC"
                            Case "CNT"
                                acAttrbReference.TextString = IIf(objInputTerminal.Count = 0, 1.0, objInputTerminal.Count)
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "0"
                            Case "MFG"
                                acAttrbReference.TextString = objInputTerminal.Manufacture
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PMFG"
                            Case "CAT"
                                acAttrbReference.TextString = objInputTerminal.Catalog
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PCAT"
                            Case "HEIGHT"
                                objInputTerminal.Height = CDbl(acAttrbReference.TextString)
                        End Select

                        acBlkRef.AttributeCollection.AppendAttribute(acAttrbReference)
                        acTransaction.AddNewlyCreatedDBObject(acAttrbReference, True)
                    End If
                Next
            End Using
        End Sub

        Public Sub DrawJumpers(acDatabase As Database, acTransaction As Transaction,
                               objInputJumper As JumperClass, acInsertPt As Point3d, dblScale As Double)

            Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(objInputJumper.Jumper.BlockPath)
            Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
            Dim acInsObjectId As ObjectId

            If acBlockTable.Has(strBlkName) Then
                Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                acInsObjectId = acCurrBlkTblRcd.Id
            Else
                Dim acNewDbDwg As New Database(False, True)
                acNewDbDwg.ReadDwgFile(objInputJumper.Jumper.BlockPath, FileOpenMode.OpenTryForReadShare, True, "")
                acInsObjectId = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                acNewDbDwg.Dispose()
            End If

            Using acBlkRef As New BlockReference(acInsertPt, acInsObjectId)
                acBlkRef.Layer = "PSYMS"
                acBlkRef.ScaleFactors = New Scale3d(dblScale)

                Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                acBlockTableRecord.AppendEntity(acBlkRef)
                acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                Dim acBlockTableAttrbRec As BlockTableRecord = acTransaction.GetObject(acInsObjectId, OpenMode.ForRead)
                Dim acAttrbObjectId As ObjectId
                For Each acAttrbObjectId In acBlockTableAttrbRec
                    Dim acAttrbEntity As Entity = acTransaction.GetObject(acAttrbObjectId, OpenMode.ForRead)
                    Dim acAttrbDefinition As AttributeDefinition = TryCast(acAttrbEntity, AttributeDefinition)
                    If (acAttrbDefinition IsNot Nothing) Then
                        Dim acAttrbReference As New AttributeReference()
                        acAttrbReference.SetAttributeFromBlock(acAttrbDefinition, acBlkRef.BlockTransform)
                        Select Case acAttrbReference.Tag
                            Case "P_TAGSTRIP"
                                acAttrbReference.TextString = objInputJumper.Jumper.TagStrip
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PTAG"
                            Case "INST"
                                acAttrbReference.TextString = objInputJumper.Jumper.Instance
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PLOC"
                            Case "LOC"
                                acAttrbReference.TextString = objInputJumper.Jumper.Location
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PLOC"
                            Case "MFG"
                                acAttrbReference.TextString = objInputJumper.Jumper.Manufacture
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PMFG"
                            Case "CNT"
                                acAttrbReference.TextString = IIf(objInputJumper.Jumper.Count = 0, 1.0, objInputJumper.Jumper.Count)
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD")
                                acAttrbReference.Layer = "0"
                            Case "CAT"
                                acAttrbReference.TextString = objInputJumper.Jumper.Catalog
                                acAttrbReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                acAttrbReference.Layer = "PCAT"
                            Case "HEIGHT"
                                objInputJumper.Jumper.Height = CDbl(acAttrbReference.TextString)
                        End Select

                        acBlkRef.AttributeCollection.AppendAttribute(acAttrbReference)
                        acTransaction.AddNewlyCreatedDBObject(acAttrbReference, True)
                    End If
                Next
            End Using
        End Sub

    End Module
End Namespace
