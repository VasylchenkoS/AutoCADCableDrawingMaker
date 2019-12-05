Imports System.IO
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.Electrical.Project

Namespace com.vasilchenko.TerminalModules
    Module AddressMarkingModule
        Friend Sub CreateFileWithAddressMarking()
            Dim acAddrList As List(Of String) = GetMarkingList()

            Dim ioFilePath = Path.GetDirectoryName(Application.DocumentManager.MdiActiveDocument.Name)
            Dim ioFileName = ProjectManager.GetInstance().GetActiveProject().GetProjectID
            ioFileName = ioFileName.Substring(ioFileName.LastIndexOf("\", StringComparison.Ordinal))
            ioFileName = ioFileName.Substring(0,ioFileName.Length - (ioFileName.Length - ioFileName.LastIndexOf(".", StringComparison.Ordinal))) & ".txt"

            If Not File.Exists(ioFilePath & ioFileName) Then
                Dim fs As FileStream = File.Create(ioFilePath & ioFileName)
                fs.Close()
            Else
                My.Computer.FileSystem.WriteAllText(ioFilePath & ioFileName, "", False)
            End If

            Using objStreamWriter As New StreamWriter(ioFilePath & ioFileName)
                For i = 0 To acAddrList.Count - 1
                    objStreamWriter.WriteLine(acAddrList(i))
                Next
            End Using
        End Sub
    End Module
End Namespace
