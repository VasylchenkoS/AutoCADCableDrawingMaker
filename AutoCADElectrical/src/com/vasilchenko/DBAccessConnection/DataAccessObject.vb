Imports Autodesk.AutoCAD.ApplicationServices
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Electrical.Project

Namespace com.vasilchenko.DBAccessConnection
    Module DataAccessObject
        'Private Const strConstFootprintPath As String = "D:\Autocad Additional Files\MyDatabase\Sources\ru-RU\Catalogs\footprint_lookup.mdb"
        'Private Const strConstDefaultCatPath As String = "D:\Autocad Additional Files\MyDatabase\Sources\ru-RU\Catalogs\default_cat.mdb"
        Private Const PANEL_GRAPHICS_PATH As String = "D:\Autocad Additional Files\MyDatabase\User Graphics\Panel\2D_Graphics\"

        Public Function GetAllLocations() As ArrayList
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Dim objLocationList As New ArrayList
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            If Not IO.File.Exists(strConstProjectDatabasePath) Then
                MsgBox("Source file not found. File way: " & strConstProjectDatabasePath & " Please open some project file and repeat.", vbCritical, "File Not Found")
                strConstProjectDatabasePath = ""
                Throw New ArgumentNullException
            End If

            sqlQuery = "SELECT DISTINCT [LOC] FROM TERMS ORDER BY [LOC] DESC"

            objDataTable = DatabaseConnection.GetOleDataTable(sqlQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("LOC"))
                Next objRow
            Else
                Throw New ArgumentException("Не назначены Location для клемм")
            End If

            Return objLocationList
        End Function

        Public Function GetAllTagstripInLocation(strLocation As String) As ArrayList
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Dim objLocationList As New ArrayList
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            sqlQuery = "SELECT DISTINCT [TAGSTRIP] " &
                            "FROM TERMS " &
                            "WHERE LOC = '" & UCase(strLocation) & "' " &
                            "ORDER BY [TAGSTRIP] ASC"

            objDataTable = DatabaseConnection.GetOleDataTable(sqlQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("TAGSTRIP"))
                Next objRow
            End If

            Return objLocationList
        End Function
        Public Function GetAllTermsInLocation(strLocation As String, strTagstrip As String) As List(Of Short)
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Dim objLocationList As New List(Of Short)
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            sqlQuery = "SELECT DISTINCT [TERM] " &
                            "FROM TERMS " &
                            "WHERE [TAGSTRIP] = '" & strTagstrip & "' AND [LOC] = '" & strLocation & "' AND [TERM] IS NOT NULL " &
                            "ORDER BY [TERM] ASC"

            objDataTable = DatabaseConnection.GetOleDataTable(sqlQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("TERM"))
                Next objRow
            End If

            objLocationList = objLocationList.OrderBy(Function(x) x).ToList()

            Return objLocationList
        End Function
        Public Function FillTermTypeData(strTagstrip As String, shtTermNumber As Short, strLocation As string) As TerminalClass
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Dim objDoubleTerminal As New TerminalClass
            Dim objTerminal As New LevelTerminalClass
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            objDoubleTerminal.TagStrip = strTagstrip
            objDoubleTerminal.Location = strLocation 
            objTerminal.TerminalNumber = shtTermNumber

            sqlQuery = "SELECT [INST], [LOC], [MFG], [CAT], [HDL], [STRIPSEQ], [LNUMBER], [LTOTAL], [CNT] " &
                    "FROM TERMS " &
                    "WHERE [TAGSTRIP] = '" & strTagstrip & "' AND [TERM] = '" & shtTermNumber & "' AND [LOC] = '" & strLocation & "' AND [MFG] IS NOT NULL AND [JUMPER_ID] IS NULL"

            objDataTable = DatabaseConnection.GetOleDataTable(sqlQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows

                    With objRow
                        objDoubleTerminal.Instance = .item("INST")
                        objDoubleTerminal.Manufacture = .item("MFG")
                        objDoubleTerminal.Catalog = .item("CAT")
                        objTerminal.Handle = .item("HDL")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            objDoubleTerminal.Count = .item("CNT")
                        Else
                            objDoubleTerminal.Count = 1
                        End If
                        If Not IsDBNull(.item("STRIPSEQ")) And (Not IsDBNull(.item("LTOTAL")) AndAlso .item("LTOTAL") <> "1") Then
                            objDoubleTerminal.MainTermNumber = .item("STRIPSEQ")
                        Else
                            objDoubleTerminal.MainTermNumber = shtTermNumber
                        End If
                        If Not IsDBNull(.item("LNUMBER")) Then
                            objTerminal.Level = CShort(.item("LNUMBER"))
                            If objTerminal.Level = 1 Then
                                objDoubleTerminal.BottomLevelTerminal = objTerminal
                            Else
                                objDoubleTerminal.TopLevelTerminal = objTerminal
                            End If
                        Else
                            objDoubleTerminal.BottomLevelTerminal = objTerminal
                        End If
                    End With
                Next objRow
            End If

            Return objDoubleTerminal
        End Function
        Public Sub FillTerminalBlockPath(ByRef objInputList As List(Of TerminalClass))
            For Each objInputTerminal As TerminalClass In objInputList
                FillSingleTerminalBlockPath(objInputTerminal)
            Next
        End Sub
        Public Sub FillSingleTerminalBlockPath(ByRef objInputTerminal As TerminalClass)
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Try
                sqlQuery = "SELECT [BLKNAM] " &
                        "FROM [" & objInputTerminal.Manufacture & "] " &
                        "WHERE [CATALOG] = '" & objInputTerminal.Catalog & "'"
                objDataTable = DatabaseConnection.GetSqlDataTable(sqlQuery, My.Settings.footprint_lookupSQLConnectionString)
                If Not IsNothing(objDataTable) Then objInputTerminal.BlockPath = PANEL_GRAPHICS_PATH & objDataTable.Rows(0).Item("BLKNAM")
            Catch ex As Exception
                MsgBox ("Объект не найден в графической базе. Запрос: " & sqlQuery )
            end try
        End Sub
        Public Sub FillTerminalConnectionsData(ByRef objInputList As List(Of TerminalClass), eDucktSide As SideEnum)
            For Each objInputTerminal As TerminalClass In objInputList
                If Not IsNothing(objInputTerminal.TopLevelTerminal) Then FillSingleTerminalConnectionsData(objInputTerminal.TopLevelTerminal, objInputTerminal.Location, objInputTerminal.TagStrip)
                If Not IsNothing(objInputTerminal.BottomLevelTerminal) Then FillSingleTerminalConnectionsData(objInputTerminal.BottomLevelTerminal, objInputTerminal.Location, objInputTerminal.TagStrip)
            Next
        End Sub

        Public Sub FillSingleTerminalConnectionsData(ByRef objInputLevelTerminal As LevelTerminalClass, strLocation As String, strTagStrip As String)
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Dim acEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            Dim objWiresDictionary As New TerminalDictionaryClass(Of String, List(Of WireClass))

            sqlQuery = "SELECT [WIRENO], [INST1], [LOC1], [NAM1], [PIN1], [INST2], [LOC2], [NAM2], [PIN2], [CBL], [TERMDESC1], [TERMDESC2], [NAMHDL1] " &
                            "FROM WFRM2ALL " &
                            "WHERE ([NAMHDL1] = '" & objInputLevelTerminal.Handle & "' OR [NAMHDL2] = '" & objInputLevelTerminal.Handle & "') " &
                            "AND (([NAM1] IS NOT NULL AND [PIN1] IS NOT NULL) OR ([NAM2] IS NOT NULL AND [PIN2] IS NOT NULL)) " &
                            "ORDER BY [WIRENO]"

            objDataTable = DatabaseConnection.GetOleDataTable(sqlQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                Dim strWireno As String = ""

                For Each objRow In objDataTable.Rows

                    Dim objWire As New WireClass
                    Dim objCable As New CableClass
                    Dim objWiresList As New List(Of WireClass)

                    'For shtI As Short = 0 To 8
                    '    If IsDBNull(objRow.ItemArray(shtI)) Then
                    '        MsgBox("Не указано значение " & objRow.Table.Columns(shtI).ToString & " для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM & ". Исправте и перезапустите программу", MsgBoxStyle.Critical, "Error")
                    '        Throw New ArgumentNullException
                    '    End If
                    'Next

                    With objRow
                        If .Item("NAMHDL1").Equals(objInputLevelTerminal.Handle) Then
                            objWire.Instance = .Item("INST2").ToString
                            objWire.ConnTag = .Item("NAM2").ToString
                            objWire.ConnPin = .Item("PIN2").ToString
                            objCable.Destination = .Item("LOC2").ToString
                            objWire.Termdesc = .Item("TERMDESC2").ToString
                        Else
                            objWire.Instance = .Item("INST1").ToString
                            objWire.ConnTag = .Item("NAM1").ToString
                            objWire.ConnPin = .Item("PIN1").ToString
                            objCable.Destination = .Item("LOC1").ToString
                            objWire.Termdesc = .Item("TERMDESC1").ToString
                        End If

                        If Not IsDBNull(.Item("WIRENO")) Then
                            strWireno = .Item("WIRENO")
                            objWire.WireNumber = strWireno
                        Else
                            MsgBox("Для вывода " & objWire.ConnTag & ":" & objWire.ConnPin & " не назначено номера провода")
                            Exit Sub
                        End If

                        If objCable.Destination = "(??)" Then objCable.Destination = ""

                        If IsDBNull(.Item("CBL")) Then
                            objWire.Cable = Nothing
                        Else
                            objCable.Mark = .Item("CBL")
                            objWire.Cable = objCable
                        End If

                        If objWire.Termdesc.Equals("E") OrElse objWire.Termdesc.Equals("I") Then
                            objWire.Termdesc = ""
                        End If
                    End With

                    If Not (objWire.ConnTag = "" And objWire.ConnPin = "" And objCable.Mark = "") Then
                        If objWiresDictionary.ContainsKey(strWireno) Then
                            objWiresDictionary.Item(strWireno).Add(objWire)
                        Else
                            objWiresList.Add(objWire)
                            objWiresDictionary.Add(key:=strWireno, value:=objWiresList)
                        End If
                    End If

                Next
            End If

            If objWiresDictionary.Count <> 0 Then

                WiresAdditionalFunctions.IdentityTags(objWiresDictionary)
                WiresAdditionalFunctions.SortCollectionByInstAndCables(objWiresDictionary)
                WiresAdditionalFunctions.SortDictionaryByInstAndCables(objWiresDictionary, strLocation)

                If objInputLevelTerminal.TerminalNumber = 5 Then
                    Debug.Print("")
                End If

                If objWiresDictionary.Count > 2 Then
                    acEditor.WriteMessage("[WARNING]:Слишком много разной маркировки проводов для клеммы " & strTagStrip & ":" & objInputLevelTerminal.TerminalNumber)
                    Throw New ArgumentOutOfRangeException
                ElseIf objWiresDictionary.Count = 1 Then
                    Dim objTempList As List(Of WireClass) = objWiresDictionary.Item(0)
                    Dim intCableNum As Integer = WiresAdditionalFunctions.CableInList(objTempList)

                    Dim objA = objTempList.FindAll(Function(x) x.WireNumber.StartsWith("D"))
                    Dim objB = objTempList.Find(Function(x) x.WireNumber.StartsWith("S"))
                    If objTempList.Find(Function(x) x.WireNumber.StartsWith("S")) IsNot Nothing Then
                        objB = objTempList.Find(Function(x) x.WireNumber.StartsWith("S"))
                    ElseIf objTempList.Find(Function(x) x.WireNumber.StartsWith("C")) IsNot Nothing Then
                        objB = objTempList.Find(Function(x) x.WireNumber.StartsWith("C"))
                    ElseIf objTempList.Find(Function(x) x.WireNumber.StartsWith("M")) IsNot Nothing Then
                        objB = objTempList.Find(Function(x) x.WireNumber.StartsWith("M"))
                    End If

                    'If eDucktSide.Equals(SideEnum.Rigth) Then
                    If intCableNum <> -1 Then
                        objInputLevelTerminal.AddWireLeft = objTempList.Item(intCableNum)
                        objTempList.RemoveAt(intCableNum)
                        objInputLevelTerminal.WiresRigthListList = (From l In objTempList
                                                                    Select l).ToList
                    Else
                        If objA IsNot Nothing Then
                            objInputLevelTerminal.WiresLeftListList = objA
                            objTempList.RemoveAll(Function(x) x.WireNumber.StartsWith("D"))
                        ElseIf objB IsNot Nothing Then
                            objInputLevelTerminal.AddWireRigth = objB
                            objTempList.Remove(objB)
                        End If

                        For i As Short = 0 To objTempList.Count - 1
                            If (i + 1) Mod 2 = 0 Then
                                objInputLevelTerminal.AddWireLeft = objTempList.Item(i)
                            Else
                                objInputLevelTerminal.AddWireRigth = objTempList.Item(i)
                            End If
                        Next
                    End If
                    ''Else
                    'If intCableNum <> -1 Then
                    '    objInputTerminal.AddWireRigth = objTempList.Item(intCableNum)
                    '    objTempList.RemoveAt(intCableNum)
                    '    objInputTerminal.WiresLeftList = (From l In objTempList
                    '                                      Select l).ToList
                    'Else
                    '    If objA IsNot Nothing Then
                    '        objInputTerminal.WiresRigthList = objA
                    '        objTempList.RemoveAll(Function(x) x.Wireno.StartsWith("D"))
                    '    ElseIf objB IsNot Nothing Then
                    '        objInputTerminal.AddWireLeft = objB
                    '        objTempList.Remove(objB)
                    '    End If

                    '    For i As Short = 0 To objTempList.Count - 1
                    '        If (i + 1) Mod 2 = 0 Then
                    '            objInputTerminal.AddWireRigth = objTempList.Item(i)
                    '        Else
                    '            objInputTerminal.AddWireLeft = objTempList.Item(i)
                    '        End If
                    '    Next
                    'End If
                    'End If
                ElseIf objWiresDictionary.Count = 2 Then

                    Dim objTempListA As List(Of WireClass) = objWiresDictionary.Item(0)
                    Dim objTempListB As List(Of WireClass) = objWiresDictionary.Item(1)

                    'If eDucktSide.Equals(SideEnum.Rigth) Then
                    'objTempListA = objWiresDictionary.Item(0)
                    'objTempListB = objWiresDictionary.Item(1)
                    'Else
                    '    objTempListA = objWiresDictionary.Item(1)
                    '    objTempListB = objWiresDictionary.Item(0)
                    'End If

                    If WiresAdditionalFunctions.CableInList(objTempListA) <> -1 Then
                        objInputLevelTerminal.WiresLeftListList = objTempListA
                        objInputLevelTerminal.WiresRigthListList = objTempListB
                    Else
                        Dim blnFlagA = objTempListA.Any(Function(x) x.WireNumber.StartsWith("D"))
                        Dim blnFlagB = objTempListB.Any(Function(x) x.WireNumber.StartsWith("S")) OrElse
                        objTempListB.Any(Function(x) x.WireNumber.StartsWith("C")) OrElse
                        objTempListB.Any(Function(x) x.WireNumber.StartsWith("M"))

                        If blnFlagB Or blnFlagA Then
                            objInputLevelTerminal.WiresLeftListList = objTempListA
                            objInputLevelTerminal.WiresRigthListList = objTempListB
                        Else
                            objInputLevelTerminal.WiresRigthListList = objTempListA
                            objInputLevelTerminal.WiresLeftListList = objTempListB
                        End If
                    End If

                End If
            End If
        End Sub

        Public Function FillJumperData(strTagstrip As String) As List(Of JumperClass)
            Dim objDataTable As DataTable
            Dim sqlQuery As String
            Dim objResultJumper As JumperClass
            Dim objAccJumper As TerminalClass
            Dim objJumperList As New List(Of JumperClass)
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            sqlQuery = "SELECT V1 AS [TERM], V3 AS [CAT], V4 AS [MFG], V5 AS [INST], V6 AS [LOC], V7 AS [CNT] " &
                            "FROM " &
                            "(SELECT Min(CInt([TERM])) AS V1, FIRST([CAT]) AS V3, FIRST([MFG]) AS V4 , FIRST([INST]) AS V5 , FIRST([LOC]) AS V6 , FIRST([CNT]) AS V7 " &
                            "FROM TERMS " &
                            "WHERE [TAGSTRIP]='" & strTagstrip & "' AND [JUMPER_ID] <> '' " &
                            "GROUP BY [JUMPER_ID]) AS T1"

            objDataTable = DatabaseConnection.GetOleDataTable(sqlQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objResultJumper = New JumperClass
                    objAccJumper = New TerminalClass
                    With objRow
                        objAccJumper.TagStrip = strTagstrip
                        objAccJumper.Instance = "163.КЛЕММЫ-ПЕРЕМЫЧКИ"
                        objAccJumper.Location = .item("LOC")
                        objAccJumper.Manufacture = .item("MFG")
                        objAccJumper.Catalog = .item("CAT")
                        objResultJumper.StartTermNum = .item("TERM")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            objAccJumper.Count = .item("CNT")
                        Else
                            objAccJumper.Count = 1
                        End If
                        objResultJumper.Jumper = objAccJumper
                    End With
                    If objResultJumper IsNot Nothing Then
                        DataAccessObject.FillSingleTerminalBlockPath(objResultJumper.Jumper)
                        If objAccJumper.Manufacture.ToLower.Equals("phoenix contact") then
                            objResultJumper.TermCount = CInt(Replace(Mid(objResultJumper.Jumper.Catalog, 5, 2), "-", ""))
                        Else if objAccJumper.Manufacture.ToLower.Equals("klemsan") then
                            objResultJumper.TermCount = CInt(Mid(objResultJumper.Jumper.Catalog, InStr(objResultJumper.Jumper.Catalog,"/")+1, 2))
                        end if 
                        objJumperList.Add(objResultJumper)
                    End If
                Next objRow
            End If

            Return objJumperList
        End Function

    End Module
End Namespace
