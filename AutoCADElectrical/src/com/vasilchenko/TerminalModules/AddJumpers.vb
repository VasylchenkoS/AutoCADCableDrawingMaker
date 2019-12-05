
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums

Namespace com.vasilchenko.TerminalModules
    Module AddJumpers
        Public Sub FillStripByJumpers(ByRef objTerminalStripList As TerminalStripClass)
            Dim objPrevItem As JumperClass = Nothing

            objTerminalStripList.JumperList = DataAccessObject.FillJumperData(objTerminalStripList.TerminalList.Item(0).TagStrip)

            objTerminalStripList.JumperList.Sort(Function(x, y) x.StartTermNum < y.StartTermNum)

            For Each objCurItem As JumperClass In objTerminalStripList.JumperList
                If objPrevItem Is Nothing Then
                    objCurItem.Side = SideEnum.Right
                ElseIf ((objPrevItem.StartTermNum + objPrevItem.TermCount - 1) = objCurItem.StartTermNum) Then
                    Dim temp = objTerminalStripList.TerminalList.Find(Function(x) x.MainTermNumber = objCurItem.StartTermNum)
                    If temp.Manufacture.ToLower.Equals("klemsan") andalso not temp.Catalog.Equals("AVK 2,5 R Gray") Then
                        Throw New Exception ("На клемме " & objCurItem.StartTermNum & " невозможно присоединить 2 перемычки")
                    End If
                    If objPrevItem.Side = SideEnum.Right Then
                        objCurItem.Side = SideEnum.Left
                    Else
                        objCurItem.Side = SideEnum.Right
                    End If
                Else
                        objCurItem.Side = SideEnum.Right
                End If
                objPrevItem = objCurItem
            Next

        End Sub
    End Module
End Namespace