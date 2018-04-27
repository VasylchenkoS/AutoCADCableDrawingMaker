Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection

Namespace com.vasilchenko.TerminalModules
    Module AddPhoenixTerminalAccessories
        Public Sub AddPhoenixAccForSignalisation(objMappingTermsList As List(Of TerminalAccessoriesClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing

            For index As Integer = 0 To objMappingTermsList.Count - 1

                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.TERM = 1 Then
                    objMappingTermsList.Insert(index, EnumFunctions.Convert(AccessoriesPhoenixContactEnum.TerminalMarker, objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index = index + 2
                ElseIf objCurTerminal.TERM = 8 Or objCurTerminal.TERM = 24 Then
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index = index + 1
                Else
                    If objCurTerminal.TERM Mod 32 = 0 Then
                        objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                        index = index + 1
                    ElseIf objCurTerminal.TERM Mod 16 = 0 Then
                        objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                        index = index + 1
                    End If
                End If
            Next
            objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(EnumFunctions.Convert(AccessoriesPhoenixContactEnum.EndClamp35, objCurTerminal))
        End Sub

        Friend Sub AddPhoenixAccForControl(objMappingTermsList As List(Of TerminalAccessoriesClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing

            For index As Integer = 0 To objMappingTermsList.Count - 1
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.TERM = 1 Then
                    objMappingTermsList.Insert(index, EnumFunctions.Convert(AccessoriesPhoenixContactEnum.TerminalMarker, objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index = index + 2
                ElseIf objCurTerminal.TERM Mod 5 = 0 Then
                    objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index = index + 1
                End If
            Next
            objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(EnumFunctions.Convert(AccessoriesPhoenixContactEnum.EndClamp35, objCurTerminal))
        End Sub

        Friend Sub AddPhoenixAccForMeasurement(objMappingTermsList As List(Of TerminalAccessoriesClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing

            For index As Integer = 0 To objMappingTermsList.Count - 1

                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.TERM = 1 Then
                    objMappingTermsList.Insert(index, EnumFunctions.Convert(AccessoriesPhoenixContactEnum.TerminalMarker, objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index = index + 2
                ElseIf objCurTerminal.TERM Mod 16 = 0 Then
                    objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index = index + 1
                End If
            Next
            objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(EnumFunctions.Convert(AccessoriesPhoenixContactEnum.EndClamp35, objCurTerminal))
        End Sub

        Private Function GetCoverObject(objCurTerminal As TerminalClass) As TerminalAccessoriesClass
            If objCurTerminal.CAT = "UT-2,5" Or objCurTerminal.CAT = "UT-2,5-PE" Then
                Return EnumFunctions.Convert(AccessoriesPhoenixContactEnum.EndCoverForUT2_5, objCurTerminal)
            ElseIf objCurTerminal.CAT = "UT-2,5-MT" Then
                Return EnumFunctions.Convert(AccessoriesPhoenixContactEnum.EndCoverForUT2_5MT, objCurTerminal)
            Else Return Nothing
            End If
        End Function
        Private Function GetPlateObject(objCurTerminal As TerminalClass) As TerminalAccessoriesClass
            If objCurTerminal.CAT = "UT-2,5" Or objCurTerminal.CAT = "UT-2,5-PE" Then
                Return EnumFunctions.Convert(AccessoriesPhoenixContactEnum.PartitionPlateForUT2_5, objCurTerminal)
            ElseIf objCurTerminal.CAT = "UT-2,5-MT" Then
                Return EnumFunctions.Convert(AccessoriesPhoenixContactEnum.PartitionPlateForUT2_5MT, objCurTerminal)
            ElseIf objCurTerminal.CAT = "UT6-HESI(6,3X32)" Then
                Return EnumFunctions.Convert(AccessoriesPhoenixContactEnum.PartitionPlateForUT6, objCurTerminal)
            Else Return Nothing
            End If
        End Function

    End Module
End Namespace
