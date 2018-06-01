Imports Autodesk.AutoCAD.ApplicationServices
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Electrical.Project

Namespace com.vasilchenko.DBAccessConnection
    Module DBDataAccessObject
        'Private Const strConstFootprintPath As String = "D:\Autocad Additional Files\MyDatabase\Sources\ru-RU\Catalogs\footprint_lookup.mdb"
        'Private Const strConstDefaultCatPath As String = "D:\Autocad Additional Files\MyDatabase\Sources\ru-RU\Catalogs\default_cat.mdb"
        Public Function GetAllLocations() As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

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

        Public Function GetAllTagstripInLocation(strLocation As String) As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

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
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

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
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

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
        Public Sub FillTerminalConnectionsData(ByRef objInputList As List(Of TerminalAccessoriesClass), eDucktSide As SideEnum)
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim acEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            For Each objInputTerminal As TerminalClass In objInputList

                Dim objWiresDictionary As New TerminalDictionaryClass(Of String, List(Of WireClass))

                strSQLQuery = "SELECT [WIRENO], [INST1], [LOC1], [NAM1], [PIN1], [INST2], [LOC2], [NAM2], [PIN2], [CBL], [TERMDESC1], [TERMDESC2], [NAMHDL1] " &
                                "FROM WFRM2ALL " &
                                "WHERE ([NAMHDL1] = '" & objInputTerminal.HDL & "' OR [NAMHDL2] = '" & objInputTerminal.HDL & "') " &
                                "AND ([NAM1] IS NOT NULL AND [PIN1] IS NOT NULL AND [NAM2] IS NOT NULL AND [PIN2] IS NOT NULL) " &
                                "ORDER BY [WIRENO]"

                objDataTable = DBConnection.GetOleDataTable(strSQLQuery, strConstProjectDatabasePath)

                If Not IsNothing(objDataTable) Then
                    Dim strWireno As String = ""

                    For Each objRow In objDataTable.Rows

                        Dim objWire As New WireClass
                        Dim objCable As New CableClass
                        Dim objWiresList As New List(Of WireClass)

                        For shtI As Short = 0 To 8
                            If IsDBNull(objRow.ItemArray(shtI)) Then
                                MsgBox("Не указано значение " & objRow.Table.Columns(shtI).ToString & " для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM & ". Исправте и перезапустите программу", MsgBoxStyle.Critical, "Error")
                                Throw New ArgumentNullException
                            End If
                        Next

                        With objRow
                            strWireno = .Item("WIRENO")
                            objWire.Wireno = strWireno

                            If .Item("NAMHDL1").Equals(objInputTerminal.HDL) Then
                                objWire.Instance = .Item("INST2").ToString
                                objWire.Name = .Item("NAM2").ToString
                                objWire.Pin = .Item("PIN2").ToString
                                objCable.Destination = .Item("LOC2").ToString
                                objWire.TERMDESC = .Item("TERMDESC2").ToString
                            Else
                                objWire.Instance = .Item("INST1").ToString
                                objWire.Name = .Item("NAM1").ToString
                                objWire.Pin = .Item("PIN1").ToString
                                objCable.Destination = .Item("LOC1").ToString
                                objWire.TERMDESC = .Item("TERMDESC1").ToString
                            End If

                            If IsDBNull(.Item("CBL")) Then
                                objWire.Cable = Nothing
                            Else
                                objCable.Mark = .Item("CBL")
                                objWire.Cable = objCable
                            End If

                            If objWire.TERMDESC.Equals("E") OrElse objWire.TERMDESC.Equals("I") Then
                                objWire.TERMDESC = ""
                            End If
                        End With

                        If objWiresDictionary.ContainsKey(strWireno) Then
                            objWiresDictionary.Item(strWireno).Add(objWire)
                        Else
                            objWiresList.Add(objWire)
                            objWiresDictionary.Add(key:=strWireno, value:=objWiresList)
                        End If

                    Next
                End If

                If objWiresDictionary.Count <> 0 Then

                    WiresAdditionalFunctions.IdentityTags(objWiresDictionary)
                    WiresAdditionalFunctions.SortCollectionByInstAndCables(objWiresDictionary)
                    WiresAdditionalFunctions.SortDictionaryByInstAndCables(objWiresDictionary, objInputTerminal.LOC)

                    If objInputTerminal.TERM = 5 Then
                        Debug.Print("")
                    End If

                    If objWiresDictionary.Count > 2 Then
                        acEditor.WriteMessage("[WARNING]:Слишком много разной маркировки проводов для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM)
                        Throw New ArgumentOutOfRangeException
                    ElseIf objWiresDictionary.Count = 1 Then
                        Dim objTempList As List(Of WireClass) = objWiresDictionary.Item(0)
                        Dim intCableNum As Integer = WiresAdditionalFunctions.CableInList(objTempList)

                        Dim objA = objTempList.FindAll(Function(x) x.Wireno.StartsWith("D"))
                        Dim objB = objTempList.Find(Function(x) x.Wireno.StartsWith("S"))
                        If objTempList.Find(Function(x) x.Wireno.StartsWith("S")) IsNot Nothing Then
                            objB = objTempList.Find(Function(x) x.Wireno.StartsWith("S"))
                        ElseIf objTempList.Find(Function(x) x.Wireno.StartsWith("C")) IsNot Nothing Then
                            objB = objTempList.Find(Function(x) x.Wireno.StartsWith("C"))
                        ElseIf objTempList.Find(Function(x) x.Wireno.StartsWith("M")) IsNot Nothing Then
                            objB = objTempList.Find(Function(x) x.Wireno.StartsWith("M"))
                        End If

                        'If eDucktSide.Equals(SideEnum.Rigth) Then
                        If intCableNum <> -1 Then
                            objInputTerminal.AddWireLeft = objTempList.Item(intCableNum)
                            objTempList.RemoveAt(intCableNum)
                            objInputTerminal.WiresRigthList = (From l In objTempList
                                                               Select l).ToList
                        Else
                            If objA IsNot Nothing Then
                                objInputTerminal.WiresLeftList = objA
                                objTempList.RemoveAll(Function(x) x.Wireno.StartsWith("D"))
                            ElseIf objB IsNot Nothing Then
                                objInputTerminal.AddWireRigth = objB
                                objTempList.Remove(objB)
                            End If

                            For i As Short = 0 To objTempList.Count - 1
                                If (i + 1) Mod 2 = 0 Then
                                    objInputTerminal.AddWireLeft = objTempList.Item(i)
                                Else
                                    objInputTerminal.AddWireRigth = objTempList.Item(i)
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
                            objInputTerminal.WiresLeftList = objTempListA
                            objInputTerminal.WiresRigthList = objTempListB
                        Else
                            Dim blnFlagA = objTempListA.Any(Function(x) x.Wireno.StartsWith("D"))
                            Dim blnFlagB = objTempListB.Any(Function(x) x.Wireno.StartsWith("S")) OrElse
                            objTempListB.Any(Function(x) x.Wireno.StartsWith("C")) OrElse
                            objTempListB.Any(Function(x) x.Wireno.StartsWith("M"))

                            If blnFlagB Or blnFlagA Then
                                objInputTerminal.WiresLeftList = objTempListA
                                objInputTerminal.WiresRigthList = objTempListB
                            Else
                                objInputTerminal.WiresRigthList = objTempListA
                                objInputTerminal.WiresLeftList = objTempListB
                            End If
                        End If

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
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            strSQLQuery = "SELECT V1 AS [TERM], V3 AS [CAT], V4 AS [MFG], V5 AS [INST], V6 AS [LOC] " &
                            "FROM " &
                            "(SELECT Min(CInt([TERM])) AS V1, FIRST([CAT]) AS V3, FIRST([MFG]) AS V4 , FIRST([INST]) AS V5 , FIRST([LOC]) AS V6 " &
                            "FROM TERMS " &
                            "WHERE [TAGSTRIP]='" & strTagstrip & "' AND [JUMPER_ID] <> '' " &
                            "GROUP BY [JUMPER_ID]) AS T1"

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
