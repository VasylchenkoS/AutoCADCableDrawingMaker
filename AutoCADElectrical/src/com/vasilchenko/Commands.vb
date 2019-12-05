Imports System.Runtime.InteropServices
Imports AutoCADElectrical.com.vasilchenko.TerminalModules
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports AutoCADElectricalSpecifications.com.vasilchenko.Forms
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules
Imports AutoCADVBNETLayoutCreator
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput


Namespace com.vasilchenko
    Public Class Commands

        <DllImport("accore.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="acedTrans")>
        Public Shared Function acedTrans(ByVal point As Double(), ByVal fromRb As IntPtr, ByVal toRb As IntPtr, ByVal disp As Integer, ByVal result As Double()) As Integer
        End Function

        <CommandMethod("ASU_Terminal_Builder", CommandFlags.Session)>
        Public Shared Sub Builder()

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            If Application.GetSystemVariable("MIRRTEXT") = "1" Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("MIRRTEXT variable set to 0")
                Application.SetSystemVariable("MIRRTEXT", 0)
            End If

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim objForm = New ufTerminalSelector
                Try
                    objForm.ShowDialog()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using

        End Sub

        <CommandMethod("ASU_Terminal_MirrorDescription", CommandFlags.Session)>
        Public Shared Sub StartSwipe()
            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim objForm = New ufTerminalSelector
                Try
                    ModifyTerminal.SwipeTerminalModule.StartSwipe()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using
        End Sub

        <CommandMethod("ASU_Terminal_Redraw", CommandFlags.Session)>
        Public Shared Sub StartCheck()
            
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim objForm = New ufTerminalSelector
                Try
                    ModifyTerminal.RedrawTerminalModule.StartRedraw()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using
        End Sub

        <CommandMethod("ASU_Terminal_MoveTermdescToRight", CommandFlags.Session)>
        Public Shared Sub MoveDescriptionToRight()
            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim objForm = New ufTerminalSelector
                Try
                    ModifyTerminal.MoveTerminalDescription.MoveRight()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using
        End Sub

        <CommandMethod("ASU_Specification", CommandFlags.Session)>
        Public Shared Sub Specification()

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction()
                Try
                    Dim uf As New SpecSelector
                    uf.ShowDialog()
                    If uf.rbProjUpdate.Checked = True Then
                    ElseIf uf.rbProjCreate.Checked = True Then
                        ProjectTableDrawing.ProjectTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    ElseIf uf.rbSheetCreate.Checked = True Then
                        KDTableDrawing.DrawSheetTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    ElseIf uf.rbSheetUpdate.Checked = True Then
                        KDTableUpdater.UpdateSheetTable(acDocument, acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно обновлена")
                    ElseIf uf.rbPageCreate.Checked = True Then
                        PageTableDrawing.DrawPagesTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    ElseIf uf.rbPageUpdate.Checked = True Then
                        PageTableUpdater.UpdatePageTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    Else
                        Exit Sub
                    End If
                    uf.Dispose()
                    acTransaction.Commit()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                    acTransaction.Abort()
                Finally
                    acTransaction.Dispose()
                End Try
            End Using

        End Sub

        <CommandMethod("ASU_LayoutCreator", CommandFlags.Session)>
        Public Shared Sub Main()
            Dim swTimer = New Stopwatch
            swTimer.Start()

            Dim strMessage As String = ""

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Dim ufStart As New ufStartForm
            Try
                ufStart.ShowDialog()
            Catch ex As Exception
                MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                strMessage = "[ERROR] Не удалось завершить программу корректно"
            Finally
                swTimer.Stop()
                If strMessage = "" Then strMessage = "Успех! Программа выполнилась за {HH:MM:SS.ms}" & swTimer.Elapsed.ToString
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("---------------------------------------------")
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(strMessage)
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("---------------------------------------------")
                ufStart.Dispose()
            End Try
        End Sub

        <CommandMethod("ASU_CableJournal_RZA", CommandFlags.Session)>
        Public Shared Sub CableJournal()

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            'Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction()
                Try
                    CableTableDrawing.CableTable(acDocument, acDatabase, acTransaction, acEditor)
                    acTransaction.Commit()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                    acTransaction.Abort()
                Finally
                    acTransaction.Dispose()
                End Try
            End Using

        End Sub

        <CommandMethod("ASU_Marking_Maker", CommandFlags.Session)>
        Public Shared Sub AddressMarking()
            AddressMarkingModule.CreateFileWithAddressMarking()
        End Sub

        <CommandMethod("ASU_TAG_CHANGE", CommandFlags.Session)>
        Public Shared Sub TagChanger()

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Try
                    TagChangerModule.ChangeTag()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using

        End Sub

    End Class
End Namespace