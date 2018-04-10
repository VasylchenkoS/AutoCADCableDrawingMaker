Imports System.Data.OleDb
Imports Autodesk.AutoCAD.ApplicationServices
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports System.Collections.ObjectModel

Namespace com.vasilchenko.DBAccessConnection
    Module DBDataAccessObject
        Private Const strConstFootprintPath As String = "C:\Temp\MyDatabase\Sources\ru-RU\Catalogs\footprint_lookup.mdb"
        Private Const strConstDefaultCatPath As String = "C:\Temp\MyDatabase\Sources\ru-RU\Catalogs\default_cat.mdb"
        Private strConstProjectDatabasePath As String

        Public Function GetAllLocations() As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList

            If strConstProjectDatabasePath = "" Then
                Dim strCustomIconPath As String = Left(Application.AcadApplication.Preferences.Files.CustomIconPath, InStrRev(Application.AcadApplication.Preferences.Files.CustomIconPath, "\",, CompareMethod.Text))
                Dim strProjectName As String = Right(Application.AcadApplication.ActiveDocument.Path, Len(Application.AcadApplication.ActiveDocument.Path) - InStrRev(Application.AcadApplication.ActiveDocument.Path, "\",, CompareMethod.Text))
                strConstProjectDatabasePath = strCustomIconPath & "User\" & UCase(strProjectName) & ".mdb"
            End If

            If Not IO.File.Exists(strConstProjectDatabasePath) Then
                MsgBox("Source file not found. File way: " & strConstProjectDatabasePath & " Please open some project file and repeat.", vbCritical, "File Not Found")
                strConstProjectDatabasePath = ""
                Throw New ArgumentNullException
            End If

            strSQLQuery = "SELECT DISTINCT [LOC] FROM TERMS ORDER BY [LOC] DESC"

            objDataTable = GetOleBdDataReader(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("LOC"))
                Next objRow
            End If

            Return objLocationList
        End Function

        Public Function GetAllTagstripInLocation(strLocation As String) As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList

            strSQLQuery = "SELECT DISTINCT [TAGSTRIP] " &
                            "FROM TERMS " &
                            "WHERE LOC = '" & UCase(strLocation) & "' " &
                            "ORDER BY [TAGSTRIP] ASC"

            objDataTable = GetOleBdDataReader(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    objLocationList.Add(objRow.Item("TAGSTRIP"))
                Next objRow
            End If

            Return objLocationList
        End Function

        Public Function GetAllTermsInLocation(strLocation As String, strTagstrip As String) As ArrayList
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objLocationList As New ArrayList

            strSQLQuery = "SELECT DISTINCT [TERM] " &
                            "FROM TERMS " &
                            "WHERE [TAGSTRIP] = '" & strTagstrip & "' AND [LOC] = '" & strLocation &
                            "' AND [TERM] <> '' AND [TERM] IS NOT NULL " &
                            "ORDER BY [TERM] ASC"

            objDataTable = GetOleBdDataReader(strSQLQuery, strConstProjectDatabasePath)

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
                    "WHERE [TAGSTRIP] = '" & strTagstrip & "' AND [TERM] = '" & strTermValue & "'"

            objDataTable = GetOleBdDataReader(strSQLQuery, strConstProjectDatabasePath)

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
        Public Function FillTerminalBlockPath(objInputTerminal As TerminalClass) As TerminalClass
            Dim objDataTable As DataTable
            Dim strSQLQuery As String

            strSQLQuery = "SELECT [BLKNAM] " &
                    "FROM [" & objInputTerminal.MFG & "] " &
                    "WHERE [CATALOG] = '" & objInputTerminal.CAT & "'"

            objDataTable = GetOleBdDataReader(strSQLQuery, strConstFootprintPath)

            If Not IsNothing(objDataTable) Then objInputTerminal.BLOCK = objDataTable.Rows(0).Item("BLKNAM")

            Return objInputTerminal
        End Function
        Public Function FillTerminalConnectionsData(objInputTerminal As TerminalClass) As TerminalClass
            Dim objDataTable As DataTable
            Dim strSQLQuery As String
            Dim objWiresDic As New Dictionary(Of String, ArrayList)

            strSQLQuery = "SELECT [WIRENO], [INST1], [LOC1],[NAM1], [PIN1], [INST2], [LOC2],[NAM2], [PIN2], [CBL] " &
                    "FROM WFRM2ALL " &
                    "WHERE [NAMHDL1] = '" & objInputTerminal.HDL & "' OR [NAMHDL2] = '" & objInputTerminal.HDL & "' " &
                    "ORDER BY [WIRENO]"

            objDataTable = GetOleBdDataReader(strSQLQuery, strConstProjectDatabasePath)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim objWire As New WireClass
                    Dim objCable As New CableClass
                    Dim objWiresList As New ArrayList
                    Dim strWireno As String = ""
                    With objRow
                        If IsDBNull(.item("WIRENO")) Then
                            MsgBox("Не промеркирована цепь для клеммы " & objInputTerminal.P_TAGSTRIP & ":" & objInputTerminal.TERM & ". Промаркируйте и перезапустите программу", vbCritical, "Error")
                            Throw New ArgumentNullException
                        Else
                            strWireno = .item("WIRENO")
                            objWire.Wireno = strWireno
                        End If
                        If .item("NAM1") <> objInputTerminal.P_TAGSTRIP And .item("PIN1") <> objInputTerminal.TERM.ToString Then
                            objWire.Instance = .item("INST1")
                            objWire.Name = .item("NAM1")
                            objWire.Pin = .item("PIN1")
                            objCable.Location = .item("LOC1")
                        Else
                            objWire.Instance = .item("INST2")
                            objWire.Name = .item("NAM2")
                            objWire.Pin = .item("PIN2")
                            objCable.Location = .item("LOC2")
                        End If

                        If Not IsDBNull(.Item("CBL")) Then
                            objCable.Mark = .item("CBL")
                            objWire.Cable = objCable
                        Else objWire.Cable = Nothing
                        End If
                        If objWiresDic.ContainsKey(strWireno) Then
                            objWiresDic(strWireno).Add(objWire)
                        Else
                            objWiresList.Add(objWire)
                            objWiresDic.Add(key:=strWireno, value:=objWiresList)
                        End If
                    End With
                Next objRow
            Else
                MsgBox("Данные не обнаружены", vbCritical, "ERROR")
                Throw New ArgumentNullException
            End If

            IdentityTags(objWiresDic)
            'SortCollectionByInstAndCables objWiresDictionary
            'SortDictionaryByInstAndCables objWiresDictionary, objInputTerminal.LOC

            Return objInputTerminal
        End Function

    End Module
End Namespace
