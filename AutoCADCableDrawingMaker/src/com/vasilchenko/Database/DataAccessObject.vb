Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.Electrical.Project
Imports AutoCADCableDrawingMaker.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices

Namespace com.vasilchenko.Database

    Module DataAccessObject

        Private strConstProjectDatabasePath As String = ""

        Friend Function GetAllLocations() As List(Of String)
            Dim locationList As New List(Of String)

            strConstProjectDatabasePath = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            If Not IO.File.Exists(strConstProjectDatabasePath) Then
                MsgBox("Source file not found. File way: " & strConstProjectDatabasePath & " Please open some project file and repeat.", vbCritical, "File Not Found")
                Throw New ArgumentNullException
            End If

            Dim strOleQuery As String = "SELECT DISTINCT [LOC] FROM [WNETLST] WHERE [CBL] IS NOT NULL ORDER BY [LOC]"
            Dim dbOleDataTable = DatabaseConnections.GetOleDataTable(strOleQuery, strConstProjectDatabasePath)

            Try
                For Each dbRow In dbOleDataTable.Rows
                    locationList.Add(dbRow.Item("LOC"))
                Next dbRow
            Catch ex As Exception
                MsgBox("Что-то не так в методе GetAllLocations: " & vbCrLf & ex.Message)
            End Try

            Return locationList
        End Function

        Friend Function FillCableData(strLocation As String) As SortedList(Of Integer, CableClass)
            Dim acCableList As New SortedList(Of Integer, CableClass)

            Dim strOleQuery As String = String.Format("SELECT [TAGNAME], [FAMILY], [DESC1], [DESC2], [DESC3], [MFG], [CAT], [INST], [LOC], [CNT] FROM [COMP] WHERE [TAGNAME] IN (SELECT DISTINCT [CBL] FROM [WNETLST]	WHERE [LOC] = '{0}') AND [PAR1_CHLD2] = '1'", strLocation)
            Dim dbOleDataTable = DatabaseConnections.GetOleDataTable(strOleQuery, strConstProjectDatabasePath)

            Try
                For Each dbRow In dbOleDataTable.Rows
                    Dim acCable As New CableClass With {
                        .Tag = dbRow.Item("TAGNAME"),
                        .Family = "CBL",
                        .Manufacture = dbRow.Item("MFG"),
                        .CatalogName = dbRow.Item("CAT"),
                        .Instance = IIf(IsDBNull(dbRow.Item("INST")), "", dbRow.Item("INST")),
                        .Location = "CABLE",
                        .Desc1 = IIf(IsDBNull(dbRow.Item("DESC1")), "", dbRow.Item("DESC1")),
                        .Desc2 = IIf(IsDBNull(dbRow.Item("DESC2")), "", dbRow.Item("DESC2")),
                        .Desc3 = IIf(IsDBNull(dbRow.Item("DESC3")), "", dbRow.Item("DESC3"))
                    }

                    If IsDBNull(dbRow.Item("LOC")) OrElse Not dbRow.Item("LOC").Equals("CABLE") Then
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(String.Format("Для кабеля {0} неверно установлено значение LOCATION. Пожалуйста, измените собственнормучно! {1}", acCable.Tag, vbCrLf))
                    End If

                    If IsDBNull(dbRow.Item("CNT")) OrElse dbRow.Item("CNT").Equals("0") Then
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(String.Format("Для кабеля {0} не указана длинна {1}", acCable.Tag, vbCrLf))
                        acCable.Count = 0
                    Else
                        acCable.Count = CDbl(dbRow.Item("CNT"))
                    End If

                    acCable.TerminalConnections = FillConnectionData(acCable, strLocation)
                    acCable.BlockPath = String.Format("{0}MSH_CBL_{1}W.dwg", My.Resources.MountGraphicsPath, acCable.TerminalConnections.Count)

                    If acCable.TerminalConnections.Exists(Function(x) x.ConnectionPoint.X = (acCable.TerminalConnections(0).ConnectionPoint.X)) Then
                        acCable.Orientation = Enums.OrientationEnum.Horisontal
                    Else
                        acCable.Orientation = Enums.OrientationEnum.Vertical
                    End If

                    acCableList.Add(CInt(acCable.TerminalConnections(0).TermNumber), acCable)
                Next dbRow
            Catch ex As Exception
                MsgBox("Что-то не так в методе FillCableData: " & vbCrLf & ex.Message)
            End Try

            Return acCableList

        End Function

        Private Function FillConnectionData(acCable As CableClass, strLocation As String) As List(Of TerminalConnectionClass)
            Dim acConnectionList As New List(Of TerminalConnectionClass)

            Dim strOleQuery As String = String.Format("SELECT [WNETLST].[WIRENO], [WNETLST].[NAM], [WNETLST].[PIN], [TERMS].[CAT] FROM [WNETLST] LEFT JOIN [TERMS] ON [WNETLST].[NAM]=[TERMS].[TAGSTRIP] AND [WNETLST].[PIN]=[TERMS].[TERM] WHERE [WNETLST].[LOC] = '{0}' AND [WNETLST].[CBL] = '{1}' AND [TERMS].[LOC] = '{0}' AND [TERMS].[PAR1_CHLD2]='1'", strLocation, acCable.Tag)
            Dim dbOleDataTable = DatabaseConnections.GetOleDataTable(strOleQuery, strConstProjectDatabasePath)

            Try
                For Each dbRow In dbOleDataTable.Rows
                    Dim acConnection As New TerminalConnectionClass With {
                        .TagStrip = dbRow.Item("NAM"),
                        .Wireno = dbRow.Item("WIRENO"),
                        .TermNumber = dbRow.Item("PIN"),
                        .CatalogName = dbRow.Item("CAT")
                    }
                    acConnectionList.Add(acConnection)
                Next
                acConnectionList.Sort(Function(x, y) CInt(x.TermNumber).CompareTo(CInt(y.TermNumber)))
                GetConnectionPoints(acConnectionList, strLocation)
            Catch ex As Exception
                MsgBox("Что-то не так в методе FillConnectionData: " & vbCrLf & ex.Message)
            End Try

            Return acConnectionList
        End Function

        Private Sub GetConnectionPoints(acConnectionList As List(Of TerminalConnectionClass), strLocation As String)
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor
            Using docLock As DocumentLock = acDocument.LockDocument()
                Using acTransaction = acDatabase.TransactionManager.StartTransaction()
                    Dim acTypedValue(1) As TypedValue
                    acTypedValue.SetValue(New TypedValue(DxfCode.Start, "INSERT"), 0)
                    acTypedValue.SetValue(New TypedValue(DxfCode.LayerName, "PSYMS"), 1)

                    Dim acSelFilter As SelectionFilter = New SelectionFilter(acTypedValue)
                    Dim acPromSelResult As PromptSelectionResult = acEditor.SelectAll(acSelFilter)

                    If acPromSelResult.Value Is Nothing Then
                        MsgBox("На листе нет блоков", MsgBoxStyle.Critical)
                        Exit Sub
                    End If

                    Dim acObjIdCol As ObjectIdCollection = New ObjectIdCollection(acPromSelResult.Value.GetObjectIds)

                    Try
                        For Each acCurId As ObjectId In acObjIdCol
                            Dim acBlckRef As BlockReference = acTransaction.GetObject(acCurId, OpenMode.ForRead)
                            Dim strTagstrip = "", strTerm = ""
                            Dim dblWidth As Double = 0, dblHeight As Double = 0
                            For Each id As ObjectId In acBlckRef.AttributeCollection
                                Dim acAttrbReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                                Select Case acAttrbReference.Tag
                                    Case "LOC"
                                        If Not strLocation.Equals(acAttrbReference.TextString) Then Exit For
                                    Case "P_TAGSTRIP"
                                        If Not acConnectionList.Any(Function(x) x.TagStrip.Equals(acAttrbReference.TextString)) Then Exit For
                                        strTagstrip = acAttrbReference.TextString
                                    Case "TERM"
                                        If Not acConnectionList.Any(Function(x) x.TermNumber.Equals(acAttrbReference.TextString)) Then Exit For
                                        strTerm = acAttrbReference.TextString
                                    Case "WIDTH"
                                        dblWidth = CDbl(acAttrbReference.TextString)
                                    Case "HEIGHT"
                                        dblHeight = CDbl(acAttrbReference.TextString)
                                End Select

                                If acConnectionList.Any(Function(x) x.TermNumber.Equals(strTerm) And x.TagStrip.Equals(strTagstrip)) Then
                                    If (acBlckRef.GeometricExtents.MaxPoint.X - acBlckRef.GeometricExtents.MinPoint.X) < (acBlckRef.GeometricExtents.MaxPoint.Y - acBlckRef.GeometricExtents.MinPoint.Y) Then
                                        acConnectionList.Find(Function(x) x.TermNumber.Equals(strTerm) And x.TagStrip.Equals(strTagstrip)).ConnectionPoint = New Point3d((acBlckRef.Position.X + (acBlckRef.ScaleFactors.X * (dblHeight / 2))), (acBlckRef.Position.Y - (acBlckRef.ScaleFactors.Y * (dblWidth / 2))), (acBlckRef.Position.Z))
                                    Else
                                        acConnectionList.Find(Function(x) x.TermNumber.Equals(strTerm) And x.TagStrip.Equals(strTagstrip)).ConnectionPoint = New Point3d((acBlckRef.Position.X - (acBlckRef.ScaleFactors.X * (dblWidth / 2))), (acBlckRef.Position.Y - (acBlckRef.ScaleFactors.Y * (dblHeight / 2))), (acBlckRef.Position.Z))
                                    End If
                                End If
                            Next id
                        Next acCurId
                    Catch ex As Exception
                        MsgBox("Что-то не так в методе GetConnectionPoints: " & vbCrLf & ex.Message)
                    End Try

                End Using
            End Using
        End Sub
    End Module
End Namespace
