Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.ModifyTerminal
    Module MoveTerminalDescription
        Friend Sub MoveRight()
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

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

                    If Not strWNL.Equals(strWNR) Then
                        MsgBox("Маркировка проводов не совпадает. Невозможно объединить")
                    Else
                        For Each id As ObjectId In acBlckRef.AttributeCollection
                            Dim acAttrbReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                            Select Case acAttrbReference.Tag
                                Case "WIRENOL"
                                    acAttrbReference.UpgradeOpen()
                                    acAttrbReference.TextString = ""
                                Case "TERMDESCL"
                                    acAttrbReference.UpgradeOpen()
                                    acAttrbReference.TextString = ""
                                Case "TERMDESCR"
                                    acAttrbReference.UpgradeOpen()
                                    acAttrbReference.TextString = strTDR & ", " & strTDL.Remove(0, 2)
                            End Select
                        Next
                        Dim acPointPositionLeft = New Point3d(acBlckRef.Position.X - (dblWidth / 2), acBlckRef.Position.Y - (dblHeight / 2), 0)
                        RemoveLeftLine(acDatabase, acEditor, acPointPositionLeft)
                        acTransaction.Commit()
                    End If
                End Using
                i = 2
            End While
        End Sub

        Private Sub RemoveLeftLine(acDatabase As Database, acEditor As Editor, acStartPointLeft As Point3d)
            Dim acBlockTable As BlockTable = acDatabase.BlockTableId.GetObject(OpenMode.ForRead)
            Dim acBlockTableRecord As BlockTableRecord = acBlockTable(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

            For Each objId As ObjectId In acBlockTableRecord
                Dim acEntity As Entity = objId.GetObject(OpenMode.ForRead)
                Dim acLine As Line = TryCast(acEntity, Line)
                If (acLine IsNot Nothing) Then
                    If (Math.Round(acLine.StartPoint.X, 4) = Math.Round(acStartPointLeft.X, 4) And
                        Math.Round(acLine.StartPoint.Y, 4) = Math.Round(acStartPointLeft.Y, 4)) OrElse
                       (Math.Round(acLine.EndPoint.X, 4) = Math.Round(acStartPointLeft.X, 4) And
                        Math.Round(acLine.EndPoint.Y, 4) = Math.Round(acStartPointLeft.Y, 4)) Then

                        acLine.UpgradeOpen()
                        acLine.Erase()
                    End If
                End If
            Next
        End Sub
    End Module
End Namespace
