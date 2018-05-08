Imports Autodesk.AutoCAD.ApplicationServices
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.DBAccessConnection
    Module DBDataAccessObject
        'Private Const strConstFootprintPath As String = "D:\Autocad Additional Files\MyDatabase\Sources\ru-RU\Catalogs\footprint_lookup.mdb"
        'Private Const strConstDefaultCatPath As String = "D:\Autocad Additional Files\MyDatabase\Sources\ru-RU\Catalogs\default_cat.mdb"
        Private strConstProjectDatabasePath As String

        Public Function GetAllLocations() As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList

            If strConstProjectDatabasePath = "" Then
                strConstProjectDatabasePath = FillProjectDataPath()
            ElseIf Right(strConstProjectDatabasePath, Len(strConstProjectDatabasePath) - InStrRev(strConstProjectDatabasePath, "\",, CompareMethod.Text)) <>
                    Right(Application.AcadApplication.ActiveDocument.Path, Len(Application.AcadApplication.ActiveDocument.Path) - InStrRev(Application.AcadApplication.ActiveDocument.Path, "\",, CompareMethod.Text)) & ".mdb" Then
                strConstProjectDatabasePath = FillProjectDataPath()
            End If

            If Not IO.File.Exists(strConstProjectDatabasePath) Then
                MsgBox("Source file not found. File way: " & strConstProjectDatabasePath & " Please open some project file and repeat.", vbCritical, "File Not Found")
                strConstProjectDatabasePath = ""
                Throw New ArgumentNullException
            End If

            strSQLQuery = "SELECT DISTINCT [LOC] FROM TERMS ORDER BY [LOC] DESC"

            objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("LOC"))
                Next objRow
            Else
                Throw New ArgumentException("Не назначены Location для клемм")
            End If

            Return objLocationList
        End Function

        Private Function FillProjectDataPath() As String
            Dim strCustomIconPath As String = Left(Application.AcadApplication.Preferences.Files.ToolPalettePath, InStrRev(Application.AcadApplication.Preferences.Files.ToolPalettePath, "\",, CompareMethod.Text))
            Dim strProjectName As String = Right(Application.AcadApplication.ActiveDocument.Path, Len(Application.AcadApplication.ActiveDocument.Path) - InStrRev(Application.AcadApplication.ActiveDocument.Path, "\",, CompareMethod.Text))
            FillProjectDataPath = strCustomIconPath & "User\" & UCase(strProjectName) & ".mdb"
        End Function

        Public Function GetAllTagstripInLocation(strLocation As String) As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList

            strSQLQuery = "SELECT DISTINCT [TAGSTRIP] " &
                            "FROM TERMS " &
                            "WHERE LOC = '" & UCase(strLocation) & "' " &
                            "ORDER BY [TAGSTRIP] ASC"

            objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("TAGSTRIP"))
                Next objRow
            End If

            Return objLocationList
        End Function

        Public Function GetAllTermsInLocation(strLocation As String, strTagstrip As String) As List(Of String)
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New List(Of String)

            strSQLQuery = "SELECT DISTINCT [TERM] " &
                            "FROM TERMS " &
                            "WHERE [TAGSTRIP] = '" & strTagstrip & "' AND [LOC] = '" & strLocation & "' AND [TERM] IS NOT NULL " &
                            "ORDER BY [TERM] ASC"

            objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("TERM"))
                Next objRow
            End If

            Return objLocationList
        End Function

        Public Function FillTermTypeData(strTagstrip As String, strTermValue As String) As TerminalClass
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objResultTerminal As New TerminalClass

            objResultTerminal.P_TAGSTRIP = strTagstrip
            objResultTerminal.TERM = strTermValue

            strSQLQuery = "SELECT [INST], [LOC], [MFG], [CAT], [HDL] " &
                    "FROM TERMS " &
                    "WHERE [TAGSTRIP] = '" & strTagstrip & "' AND [TERM] = '" & strTermValue & "' AND [MFG] IS NOT NULL AND [JUMPER_ID] IS NULL"

            objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    With objRow
                        objResultTerminal.INST = .item("INST")
                        objResultTerminal.LOC = .item("LOC")
                        objResultTerminal.MFG = .item("MFG")
                        objResultTerminal.CAT = .item("CAT")
                        objResultTerminal.HDL = .item("HDL")
                    End With
                Next objRow
            End If

            Return objResultTerminal
        End Function

        Public Sub FillTerminalBlockPath(ByRef objInputList As List(Of TerminalAccessoriesClass))
            For Each objInputTerminal As TerminalClass In objInputList
                FillSingleTerminalBlockPath(objInputTerminal)
            Next
        End Sub
        Public Sub FillSingleTerminalBlockPath(ByRef objInputTerminal As TerminalAccessoriesClass)
            Dim objDataTable As DataTable
            Dim strSQLQuery As String

            strSQLQuery = "SELECT [BLKNAM] " &
                    "FROM [" & objInputTerminal.MFG & "] " &
                    "WHERE [CATALOG] = '" & objInputTerminal.CAT & "'"
            objDataTable = DBConnection.GetSQLDataTable(strSQLQuery, My.Settings.footprint_lookupSQLConnectionString)
            If Not IsNothing(objDataTable) Then objInputTerminal.BLOCK = objDataTable.Rows(0).Item("BLKNAM")
        End Sub

        'тут хуйняхуйовая )))

        Public Sub FillTerminalConnectionsData(ByRef objInputList As List(Of TerminalAccessoriesClass), eDucktSide As SideEnum)
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim acEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor

            For Each objInputTerminal As TerminalClass In objInputList

                Dim objWiresDictionary As New TerminalDictionaryClass(Of String, List(Of WireClass))

                If objInputTerminal.TERM = 16 Then
                    Debug.Print("")
                End If

                strSQLQuery = "SELECT [WIRENO], [INST1], [LOC1],[NAM1], [PIN1], [INST2], [LOC2],[NAM2], [PIN2], [CBL], [TERMDESC1], [TERMDESC2] " &
                    "FROM WFRM2ALL " &
                    "WHERE ([NAMHDL1] = '" & objInputTerminal.HDL & "' OR [NAMHDL2] = '" & objInputTerminal.HDL & "') " &
                    "AND ([NAM1] IS NOT NULL AND [PIN1] IS NOT NULL AND [NAM2] IS NOT NULL AND [PIN2] IS NOT NULL) " &
                    "ORDER BY [WIRENO]"

                objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

                If Not IsNothing(objDataTable) Then
                    For Each objRow In objDataTable.Rows
                        Dim objWire As New WireClass
                        Dim objCable As New CableClass
                        Dim objWiresList As New List(Of WireClass)
                        Dim strWireno As String = ""
                        With objRow
                            If IsDBNull(.item("WIRENO")) Then
                                MsgBox("Не промеркирована цепь для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM & ". Промаркируйте и перезапустите программу", vbCritical, "Error")
                                Throw New ArgumentNullException
                            Else
                                strWireno = .item("WIRENO")
                                objWire.Wireno = strWireno
                            End If

                            If Not .item("NAM1") = objInputTerminal.P_TAGSTRIP Then
                                objWire.Instance = .item("INST1").ToString
                                objWire.Name = .item("NAM1").ToString
                                objWire.Pin = .item("PIN1").ToString
                                objCable.Location = .item("LOC1").ToString
                                objWire.TERMDESC = .item("TERMDESC1").ToString
                            ElseIf Not .item("PIN1") = objInputTerminal.TERM.ToString Then
                                objWire.Instance = .item("INST1").ToString
                                objWire.Name = .item("NAM1").ToString
                                objWire.Pin = .item("PIN1").ToString
                                objCable.Location = .item("LOC1").ToString
                                objWire.TERMDESC = .item("TERMDESC1").ToString
                            Else
                                objWire.Instance = .item("INST2").ToString
                                objWire.Name = .item("NAM2").ToString
                                objWire.Pin = .item("PIN2").ToString
                                objCable.Location = .item("LOC2").ToString
                                objWire.TERMDESC = .item("TERMDESC2").ToString
                            End If

                            If Not IsDBNull(.Item("CBL")) Then
                                objCable.Mark = .item("CBL")
                                objWire.Cable = objCable
                            Else objWire.Cable = Nothing
                            End If
                            If objWiresDictionary.ContainsKey(strWireno) Then
                                objWiresDictionary.Item(strWireno).Add(objWire)
                            Else
                                objWiresList.Add(objWire)
                                objWiresDictionary.Add(key:=strWireno, value:=objWiresList)
                            End If
                        End With
                    Next objRow
                End If

                If objWiresDictionary.Count <> 0 Then

                    WiresAdditionalFunctions.IdentityTags(objWiresDictionary)
                    WiresAdditionalFunctions.SortCollectionByInstAndCables(objWiresDictionary)
                    WiresAdditionalFunctions.SortDictionaryByInstAndCables(objWiresDictionary, objInputTerminal.LOC)

                    If objWiresDictionary.Count > 2 Then
                        acEditor.WriteMessage("[WARNING]:Слишком много разной маркировки проводов для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM)
                        Throw New ArgumentOutOfRangeException
                    ElseIf objWiresDictionary.Count = 1 Then
                        Dim objTempList As List(Of WireClass) = objWiresDictionary.Item(0)
                        Dim intCableNum As Integer = WiresAdditionalFunctions.CableInList(objTempList)
                        If eDucktSide.Equals(SideEnum.Rigth) Then
                            For intA As Integer = 0 To objTempList.Count - 1
                                If intA <> intCableNum And objInputTerminal.WiresRigthList.Count < 2 Then
                                    objInputTerminal.WireRigth = objTempList.Item(intA)
                                ElseIf objInputTerminal.WiresLeftList.Count < 2 Then
                                    objInputTerminal.WireLeft = objTempList.Item(intA)
                                Else
                                    acEditor.WriteMessage("[WARNING]:Слишком много соединений для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM)
                                End If
                            Next
                        Else
                            For intA As Integer = 0 To objTempList.Count - 1
                                If intA <> intCableNum And objInputTerminal.WiresLeftList.Count < 2 Then
                                    objInputTerminal.WireLeft = objTempList.Item(intA)
                                ElseIf objInputTerminal.WiresRigthList.Count < 2 Then
                                    objInputTerminal.WireRigth = objTempList.Item(intA)
                                Else
                                    acEditor.WriteMessage("[WARNING]:Слишком много соединений для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM)
                                End If
                            Next
                        End If
                    ElseIf objWiresDictionary.Count = 2 Then
                        For intI As Integer = 0 To objWiresDictionary.Count - 1
                            Dim objTempList = objWiresDictionary.Item(intI)
                            Dim intCableNum = WiresAdditionalFunctions.CableInList(objTempList)
                            If eDucktSide.Equals(SideEnum.Rigth) And objInputTerminal.WiresRigthList.Count = 0 Then
                                For intA As Integer = 0 To objTempList.Count - 1
                                    objInputTerminal.WireRigth = objTempList.Item(intA)
                                    If objInputTerminal.WiresRigthList.Count > 2 Then
                                        acEditor.WriteMessage("[WARNING]:Слишком много соединений для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM)
                                    End If
                                Next
                            Else
                                For intA As Integer = 0 To objTempList.Count - 1
                                    objInputTerminal.WireLeft = objTempList.Item(intA)
                                    If objInputTerminal.WiresLeftList.Count > 2 Then
                                        acEditor.WriteMessage("[WARNING]:Слишком много соединений для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM)
                                    End If
                                Next
                            End If
                        Next
                    End If

                End If
            Next

        End Sub

        Public Function FillJumperData(strTagstrip As String) As List(Of JumperClass)
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objResultJumper As JumperClass
            Dim objAccJumper As TerminalAccessoriesClass
            Dim objJumperList As New List(Of JumperClass)

            strSQLQuery = "SELECT V1 AS [TERM], V3 AS [CAT], V4 AS [MFG], V5 AS [INST], V6 AS [LOC] " &
                            "FROM " &
                            "(SELECT Min(CInt([TERM])) AS V1, FIRST([CAT]) AS V3, FIRST([MFG]) AS V4 , FIRST([INST]) AS V5 , FIRST([LOC]) AS V6 " &
                            "FROM TERMS " &
                            "WHERE [TAGSTRIP]='" & strTagstrip & "' AND [JUMPER_ID] <> '' " &
                            "GROUP BY [JUMPER_ID])  AS T1"

            objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objResultJumper = New JumperClass
                    objAccJumper = New TerminalAccessoriesClass
                    With objRow
                        objAccJumper.P_TAGSTRIP = strTagstrip
                        objAccJumper.INST = .item("INST")
                        objAccJumper.LOC = .item("LOC")
                        objAccJumper.MFG = .item("MFG")
                        objAccJumper.CAT = .item("CAT")
                        objResultJumper.Jumper = objAccJumper
                        objResultJumper.StartTermNum = .item("TERM")
                    End With
                    If objResultJumper IsNot Nothing Then
                        DBDataAccessObject.FillSingleTerminalBlockPath(objResultJumper.Jumper)
                        objResultJumper.TermCount = CInt(Replace(Mid(objResultJumper.Jumper.CAT, 5, 2), "-", ""))
                        objJumperList.Add(objResultJumper)
                    End If
                Next objRow
            End If

            Return objJumperList
        End Function

    End Module
End Namespace
