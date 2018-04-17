Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.Main
    Public Class Commands

        <CommandMethod("DebStart", CommandFlags.Session)>
        Public Shared Sub Main()
            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            'Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim objForm = New ufTerminalSelector

                Try
                    objForm.ShowDialog()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using
        End Sub


        <CommandMethod("InsertBlk")>
        Public Sub InsertBlk()
            'Dim acCurDb As Database = Application.DocumentManager.MdiActiveDocument.Database

            'Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction
            '    Try
            '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            '        'Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            '        Dim acObjID As ObjectId

            '        If acBlkTbl.Has("FPT_002_UT2_5") Then
            '            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl.Item("FPT_002_UT2_5"), OpenMode.ForRead)
            '            acObjID = acBlkTblRec.Id
            '        Else
            '            Dim dbDwg As New Database(False, True)
            '            dbDwg.ReadDwgFile("D:/Autocad Additional Files/MyDatabase/User Graphics/Panel/2D_Graphics/FPT_002_UT2_5.dwg", IO.FileShare.Read, True, "")
            '            acObjID = acCurDb.Insert("FPT_002_UT2_5", dbDwg, False)
            '            dbDwg.Dispose()
            '        End If


            '        acTrans.Commit()
            '    Catch
            '        MsgBox(Err.Description)
            '    Finally
            '        acTrans.Dispose()
            '    End Try
            'End Using


        End Sub


    End Class
End Namespace