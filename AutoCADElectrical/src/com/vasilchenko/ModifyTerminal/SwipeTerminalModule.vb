Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.ModifyTerminal
    Module SwipeTerminalModule
        Public Sub StartSwipe()
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor
            Dim acPointPosition As Point3d

            Dim i As Short = 1
            While i <> 0
                Dim acPromptSelOpt As New PromptSelectionOptions()
                acPromptSelOpt.MessageForAdding = vbLf & "Select Terminal to swipe:"
                acPromptSelOpt.SingleOnly = True
                Dim acResult As PromptSelectionResult = acEditor.GetSelection(acPromptSelOpt)

                If acResult.Status <> PromptStatus.OK Then
                    If i = 1 Then Application.ShowAlertDialog("Try to select a block next time.")
                    Exit Sub
                End If

                Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
                    Dim acBlckRef As BlockReference = acTransaction.GetObject(acResult.Value(0).ObjectId, OpenMode.ForRead)

                    Dim strWNL = "", strWNR = "", strTDL = "", strTDR = ""
                    Dim dblWidth As Double = 0, dblHeight As Double = 0

                    For Each id As ObjectId In acBlckRef.AttributeCollection
                        Dim acAttrbReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                        Select Case acAttrbReference.Tag
                            Case "WIRENOL"
                                strWNL = acAttrbReference.TextString
                            Case "WIRENOR"
                                strWNR = acAttrbReference.TextString
                            Case "TERMDESCL"
                                strTDL = acAttrbReference.TextString
                            Case "TERMDESCR"
                                strTDR = acAttrbReference.TextString
                            Case "WIDTH"
                                dblWidth = CDbl(acAttrbReference.TextString)
                            Case "HEIGHT"
                                dblHeight = CDbl(acAttrbReference.TextString)
                        End Select
                    Next
                    For Each id As ObjectId In acBlckRef.AttributeCollection
                        Dim acAttrbReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                        Select Case acAttrbReference.Tag
                            Case "WIRENOL"
                                acAttrbReference.UpgradeOpen()
                                acAttrbReference.TextString = strWNR
                            Case "WIRENOR"
                                acAttrbReference.UpgradeOpen()
                                acAttrbReference.TextString = strWNL
                            Case "TERMDESCL"
                                acAttrbReference.UpgradeOpen()
                                acAttrbReference.TextString = strTDR
                                If strTDL = "" Then
                                    acPointPosition = New Point3d(acBlckRef.Position.X + (dblWidth / 2), acBlckRef.Position.Y - (dblHeight / 2), 0)
                                    SwipeLine(acDatabase, acEditor, acTransaction, acPointPosition, -dblWidth)
                                End If
                            Case "TERMDESCR"
                                acAttrbReference.UpgradeOpen()
                                acAttrbReference.TextString = strTDL
                                If strTDR = "" Then
                                    acPointPosition = New Point3d(acBlckRef.Position.X - (dblWidth / 2), acBlckRef.Position.Y - (dblHeight / 2), 0)
                                    SwipeLine(acDatabase, acEditor, acTransaction, acPointPosition, dblWidth)
                                End If
                        End Select
                    Next
                    acTransaction.Commit()
                End Using
                i = 2
            End While
        End Sub

        Private Sub SwipeLine(acDatabase As Database, acEditor As Editor, acTransaction As Transaction, acStartPoint As Point3d, dblXOffset As Double)
            Dim acBlockTable As BlockTable = acDatabase.BlockTableId.GetObject(OpenMode.ForRead)
            Dim acBlockTableRecord As BlockTableRecord = acBlockTable(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)
            For Each objId As ObjectId In acBlockTableRecord
                Dim acEntity As Entity = objId.GetObject(OpenMode.ForRead)
                Dim acLine As Line = TryCast(acEntity, Line)
                If (acLine IsNot Nothing) Then
                    If (Math.Round(acLine.StartPoint.X, 4) = Math.Round(acStartPoint.X, 4) And
                        Math.Round(acLine.StartPoint.Y, 4) = Math.Round(acStartPoint.Y, 4)) OrElse
                       (Math.Round(acLine.EndPoint.X, 4) = Math.Round(acStartPoint.X, 4) And
                        Math.Round(acLine.EndPoint.Y, 4) = Math.Round(acStartPoint.Y, 4)) Then

                        acLine.UpgradeOpen()
                        Dim acLineLenght = acLine.Length

                        acLine.StartPoint = New Point3d(acLine.StartPoint.X + dblXOffset, acLine.StartPoint.Y, acLine.StartPoint.Z)

                        If dblXOffset < 0 Then
                            acLine.EndPoint = New Point3d(acLine.EndPoint.X + (dblXOffset - acLineLenght * 2), acLine.EndPoint.Y, acLine.EndPoint.Z)
                        Else
                            acLine.EndPoint = New Point3d(acLine.EndPoint.X + (dblXOffset + acLineLenght * 2), acLine.EndPoint.Y, acLine.EndPoint.Z)
                        End If
                    End If
                End If
            Next
        End Sub
    End Module
End Namespace
