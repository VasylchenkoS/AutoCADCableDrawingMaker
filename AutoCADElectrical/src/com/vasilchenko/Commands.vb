Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko.Main
    Public Class Commands

        <CommandMethod("DebStart", CommandFlags.Session)>
        Public Shared Sub Main()

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

    End Class
End Namespace