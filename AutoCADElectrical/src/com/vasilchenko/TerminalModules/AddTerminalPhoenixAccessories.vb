Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection

Namespace com.vasilchenko.TerminalModules
    Module AddTerminalPhoenixAccessories
        Public Sub AddPhoenixAccForSignalisation(objMappingTermsCollection As List(Of TerminalAccessoriesClass), eDucktSide As DuctSideEnum)
            If objMappingTermsCollection.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim obPrevTerminal As TerminalClass = Nothing

            For index As Integer = 0 To objMappingTermsCollection.Count - 1
                objCurTerminal = objMappingTermsCollection.Item(index)
                If objCurTerminal.TERM = "1" Then
                    objMappingTermsCollection.Insert(index, EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.TerminalMarker, objCurTerminal))
                    objMappingTermsCollection.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index = index + 2
                ElseIf objCurTerminal.TERM = "8" Or objCurTerminal.TERM = "24" Then
                    objMappingTermsCollection.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index = index + 1
                ElseIf objCurTerminal.TERM = "16" Then
                    objMappingTermsCollection.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index = index + 1
                ElseIf eDucktSide = DuctSideEnum.Left AndAlso Not IsNothing(obPrevTerminal) Then
                    If Left(objCurTerminal.WiresLeftList.Item(0).Wireno, 1) <> Left(obPrevTerminal.WiresLeftList.Item(0).Wireno, 1) Then
                        If CInt(objCurTerminal.TERM) Mod 32 = 1 Then
                            objMappingTermsCollection.Insert(index + 1, GetPlateObject(objCurTerminal))
                        Else
                            objMappingTermsCollection.Insert(index + 1, GetCoverObject(objCurTerminal))
                        End If
                        index = index + 1
                    ElseIf eDucktSide = DuctSideEnum.Rigth AndAlso Not IsNothing(obPrevTerminal) Then
                        If Left(objCurTerminal.WiresRigthList.Item(0).Wireno, 1) <> Left(obPrevTerminal.WiresRigthList.Item(0).Wireno, 1) Then
                            If CInt(objCurTerminal.TERM) Mod 32 = 1 Then
                                objMappingTermsCollection.Insert(index + 1, GetPlateObject(objCurTerminal))
                            Else
                                objMappingTermsCollection.Insert(index + 1, GetCoverObject(objCurTerminal))
                            End If
                            index = index + 1
                        End If
                    End If
                End If
                obPrevTerminal = objCurTerminal
            Next
            objMappingTermsCollection.Add(GetCoverObject(objCurTerminal))
            objMappingTermsCollection.Add(EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.EndClamp35, objCurTerminal))
        End Sub

        Private Function GetCoverObject(objCurTerminal As TerminalClass) As TerminalAccessoriesClass
            If objCurTerminal.CAT = "UT-2,5" Or objCurTerminal.CAT = "UT-2,5-PE" Then
                Return EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.EndCoverForUT2_5, objCurTerminal)
            ElseIf objCurTerminal.CAT = "UT-2,5-MT" Then
                Return EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.EndCoverForUT2_5MT, objCurTerminal)
            Else Return Nothing
            End If
        End Function
        Private Function GetPlateObject(objCurTerminal As TerminalClass) As TerminalAccessoriesClass
            If objCurTerminal.CAT = "UT-2,5" Or objCurTerminal.CAT = "UT-2,5-PE" Then
                Return EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.PartitionPlateForUT2_5, objCurTerminal)
            ElseIf objCurTerminal.CAT = "UT-2,5-MT" Then
                Return EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.PartitionPlateForUT2_5MT, objCurTerminal)
            ElseIf objCurTerminal.CAT = "UT6-HESI(6,3X32)" Then
                Return EnumToTerminalClassConvertor(AccessoriesPhoenixContactEnum.PartitionPlateForUT6, objCurTerminal)
            Else Return Nothing
            End If
        End Function

        Private Function EnumToTerminalClassConvertor(eAccEnum As AccessoriesPhoenixContactEnum, objCurTerminal As TerminalClass) As TerminalAccessoriesClass
            Dim objTerminalAcc As New TerminalAccessoriesClass

            objTerminalAcc.P_TAGSTRIP = objCurTerminal.P_TAGSTRIP
            objTerminalAcc.INST = objCurTerminal.INST
            objTerminalAcc.LOC = objCurTerminal.LOC
            objTerminalAcc.MFG = "Phoenix Contact"
            objTerminalAcc.CAT = New EnumDescriptor(Of AccessoriesPhoenixContactEnum)(eAccEnum).ToString
            DBDataAccessObject.FillTerminalBlockPath(objTerminalAcc)
            Return objTerminalAcc
        End Function

    End Module
End Namespace
