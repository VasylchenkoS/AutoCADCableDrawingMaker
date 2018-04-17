Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices

Namespace com.vasilchenko.TerminalModules
    Module DrawingLayerChecker
        Public Sub CheckLayers()
            Dim acadDocument As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim acadCurDB As Database = acadDocument.Database
            Dim objLayerList As New List(Of String) _
                ({"PASSY", "PCAT", "PDESC", "PGRP", "PITEM", "PLOC", "PMFG", "PMISC", "PMNT", "PRTG", "PSYMS", "PTAG", "PTERM", "PWIRE"})

            Using acadTransanction As Transaction = acadCurDB.TransactionManager.StartTransaction()
                Dim acadLayerTbl As LayerTable
                acadLayerTbl = acadTransanction.GetObject(acadCurDB.LayerTableId, OpenMode.ForRead)

                For Each sLayerName As String In objLayerList
                    If acadLayerTbl.Has(sLayerName) = False Then
                        Using acadLayerTblRec As LayerTableRecord = New LayerTableRecord()
                            acadLayerTblRec.Name = sLayerName
                            Select Case sLayerName
                                Case "PASSY"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 134)
                                Case "PCAT"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 9)
                                Case "PDESC"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1)
                                Case "PGRP"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8)
                                Case "PITEM"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 20)
                                Case "PLOC"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8)
                                Case "PMFG"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 9)
                                Case "PMISC"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 2)
                                Case "PMNT"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8)
                                Case "PRTG"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 130)
                                Case "PSYMS"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
                                Case "PTAG"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 51)
                                Case "PTERM"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 6)
                                Case "PWIRE"
                                    acadLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 84)
                            End Select

                            acadLayerTbl.UpgradeOpen()
                            acadLayerTbl.Add(acadLayerTblRec)

                            acadTransanction.AddNewlyCreatedDBObject(acadLayerTblRec, True)
                        End Using
                    End If
                Next

                acadTransanction.Commit()
            End Using
        End Sub
    End Module
End Namespace