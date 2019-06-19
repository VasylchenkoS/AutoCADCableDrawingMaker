Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums

Namespace com.vasilchenko.TerminalModules
    Module AddTerminalAccessories
        Private ReadOnly UncoverObjects As new List(Of String) From  {"UT 4-MT", "AVK 2,5/4T Yellow-Green"}
        Friend Sub AddAccessoriesForSignalisation(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim intCount As Integer = objMappingTermsList.Count - 2
            Dim index As Integer = 0

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber = 1 Then
                    objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index += 2
                    intCount += 2
                ElseIf objCurTerminal.MainTermNumber = 8 Or objCurTerminal.MainTermNumber = 24 Then
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index += 1
                    intCount += 1
                Else
                    If objCurTerminal.MainTermNumber Mod 32 = 0 Then
                        objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                        index += 1
                        intCount += 1
                    ElseIf objCurTerminal.MainTermNumber Mod 16 = 0 Then
                        objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                        index += 1
                        intCount += 1
                    End If
                End If
                index += 1
            Loop
            objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Friend Sub AddAccessoriesForPower(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim objPrevTerminal As TerminalClass = Nothing
            Dim intCount As Integer = objMappingTermsList.Count
            Dim index As Integer = 0

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber = 1 Then
                    objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index += 2
                    intCount += 2
                ElseIf (objCurTerminal.MainTermNumber = 2 Or objCurTerminal.MainTermNumber = 4) And objCurTerminal.Catalog = "UT 6-HESI (6,3X32)" Then
                    objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index += 1
                    intCount += 1
                ElseIf objCurTerminal.Catalog <> "UT 6-HESI (6,3X32)" Then
                    If ((objPrevTerminal.BottomLevelTerminal.WireContains("L") Or objPrevTerminal.BottomLevelTerminal.WireContains("PP") Or objPrevTerminal.BottomLevelTerminal.WireContains("+") Or objPrevTerminal.BottomLevelTerminal.WireContains("PV")) And
                        (objCurTerminal.BottomLevelTerminal.WireContains("N") Or objCurTerminal.BottomLevelTerminal.WireContains("PM") Or objCurTerminal.BottomLevelTerminal.WireContains("-") Or objCurTerminal.BottomLevelTerminal.WireContains("PT")))  Then
                        objMappingTermsList.Insert(index, GetCoverObject(objCurTerminal))
                        index += 1
                        intCount += 1
                    ElseIf (objPrevTerminal.BottomLevelTerminal.WireContains("N") Or objPrevTerminal.BottomLevelTerminal.WireContains("PM") Or objPrevTerminal.BottomLevelTerminal.WireContains("-") Or objPrevTerminal.BottomLevelTerminal.WireContains("PT")) And
                        (objCurTerminal.BottomLevelTerminal.WireContains("L") Or objCurTerminal.BottomLevelTerminal.WireContains("PP") Or objCurTerminal.BottomLevelTerminal.WireContains("+") Or objCurTerminal.BottomLevelTerminal.WireContains("PV")) Then
                        objMappingTermsList.Insert(index, GetPlateObject(objCurTerminal))
                        index += 1
                        intCount += 1
                    Else if objCurTerminal.BottomLevelTerminal.WireNameStartsWith("PE") And Not objPrevTerminal.BottomLevelTerminal.WireContains("PE") then
                        objMappingTermsList.Insert(index, GetPlateObject(objPrevTerminal))
                        index += 1
                        intCount += 1
                    End If
                End If
                index += 1
                
                objPrevTerminal = objCurTerminal
            Loop
            if Not UncoverObjects.Contains (objPrevTerminal.Catalog) Then objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Friend Sub AddAccessoriesForControl(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim intCount As Integer = objMappingTermsList.Count - 2
            Dim index As Integer = 0
            Dim shtCount As Short = 0

            Dim msgResult = MsgBox("Yes - пяти проводная, No - четырех проводная, Cancel - трех проводная", MsgBoxStyle.YesNoCancel)
            Select Case msgResult
                Case MsgBoxResult.YES
                    shtCount = 5
                Case MsgBoxResult.No
                    shtCount = 4
                Case MsgBoxResult.Cancel
                    shtCount = 3
            End Select

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber = 1 Then
                    objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index += 2
                    intCount += 2
                ElseIf objCurTerminal.MainTermNumber Mod shtCount = 0 Then
                    objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index += 1
                    intCount += 1
                End If
                index += 1
            Loop
            objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Friend Sub AddAccessoriesForMeasurement(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim intCount As Integer = objMappingTermsList.Count - 2
            Dim index As Integer = 0

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber = 1 Then
                    objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                    index += 2
                    intCount += 2
                ElseIf objCurTerminal.MainTermNumber Mod 16 = 0 Then
                    objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index += 1
                    intCount += 1
                End If
                index += 1
            Loop
            objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Public Sub AddAccessoriesForMetering(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim blnPrevWirenum = False
            Dim blnCurWirenum as Boolean
            Dim index = 0
            Dim blnIsTt = False

            Dim msgResult = MsgBox("Клеммник цепей ТТ?", MsgBoxStyle.YesNo)
            Select Case msgResult
                Case MsgBoxResult.Yes
                    blnIsTt = True
                Case MsgBoxResult.No
                    blnIsTt = False
            End Select

            Do While index <> objMappingTermsList.Count
                objCurTerminal = objMappingTermsList.Item(index)
                blnCurWirenum = objCurTerminal.BottomLevelTerminal.WireNameStartsWith("A")
                If objCurTerminal.MainTermNumber = 1 Then
                    objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                    if not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then 
                        objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                        index += 1
                    End If
                    index += 1
                ElseIf blnCurWirenum And blnIsTt Then
                    if not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then
                        objMappingTermsList.Insert(index, GetCoverObject(objCurTerminal))
                        objMappingTermsList.Insert(index + 1, AddTerminalMarker(objCurTerminal))
                        objMappingTermsList.Insert(index + 2, GetCoverObject(objCurTerminal))
                    index += 3
                    Else 
                        objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                        index += 1
                    end if 
                ElseIf blnPrevWirenum And blnCurWirenum And Not blnIsTt Then
                    if not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then
                        objMappingTermsList.Insert(index, GetCoverObject(objCurTerminal))
                        index += 1
                    End If
                    blnPrevWirenum = False
                End If
                index += 1

                If objCurTerminal.BottomLevelTerminal.WiresLeftListList.Count <> 0 Or objCurTerminal.BottomLevelTerminal.WiresRigthListList.Count <> 0 Then
                    blnPrevWirenum = objCurTerminal.BottomLevelTerminal.WireNameStartsWith("N") Or objCurTerminal.BottomLevelTerminal.WireNameStartsWith("O")
                End If
            Loop
            if not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Private Function GetCoverObject(objCurTerminal As TerminalClass) As TerminalClass
            Select Case objCurTerminal.Catalog
                Case "UT 2,5", "UT 2,5-PE", "UT 4", "UT 6"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_EndCoverForUT2_5, objCurTerminal)
                Case "UT 2,5-MT", "UT 4-MTD-DIO/R-L"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_EndCoverForUT2_5MT, objCurTerminal)
                Case "UTTB 2,5-MT-P/P"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_EndCoverForUTTB2_5, objCurTerminal)
                Case "URTK 6"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_EndCoverForURTK_6, objCurTerminal)
                Case "UT 6-HESI (6,3X32)"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_PartitionPlateForUT6, objCurTerminal)
                Case "AVK 4 A Gray", "AVK 2,5 A Gray", "AVK 2,5 R Gray"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.K_EndCoverForAVK2_5_R_AVK_4_A , objCurTerminal)
                Case "AVK 2,5 Gray", "AVK 4 Gray", "AVK 6 Gray", "AVK 10 Gray"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.K_EndCoverForForAVK2_5_10, objCurTerminal)
                Case Else
                    Return Nothing
            End Select
        End Function
        Private Function GetPlateObject(objCurTerminal As TerminalClass) As TerminalClass

            If objCurTerminal.Manufacture = "Klemsan" Then Return GetCoverObject(objCurTerminal)

            Select Case objCurTerminal.Catalog
                Case "UT 2,5", "UT 2,5-PE", "UT 4", "UT 6"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_PartitionPlateForUT2_5, objCurTerminal)
                Case "UT 2,5-MT", "UT 4-MTD-DIO/R-L"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_PartitionPlateForUT2_5MT, objCurTerminal)
                Case "UTTB 2,5-MT-P/P"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_PartitionPlateForUTTB2_5, objCurTerminal)
                Case "UT 6-HESI (6,3X32)"
                    Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_PartitionPlateForUT6, objCurTerminal)
                Case Else
                    Return Nothing
            End Select
        End Function

        Private Function AddEndClamp(objCurTerminal As TerminalClass) As TerminalClass
            if objCurTerminal.Manufacture.ToLower.Equals("phoenix contact")
                Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_EndClamp35, objCurTerminal)
            Else if objCurTerminal.Manufacture.ToLower.Equals("klemsan")
                Return  EnumFunctions.Convert(TerminaAccessoriesEnum.K_EndClamp , objCurTerminal)
            end if 
            Return Nothing 
        End Function

        Private Function AddTerminalMarker(objCurTerminal As TerminalClass) As TerminalClass
            if objCurTerminal.Manufacture.ToLower.Equals("phoenix contact")
                Return EnumFunctions.Convert(TerminaAccessoriesEnum.PC_TerminalMarker, objCurTerminal)
            Else if objCurTerminal.Manufacture.ToLower.Equals("klemsan")
                Return  EnumFunctions.Convert(TerminaAccessoriesEnum.K_TerminalMarker, objCurTerminal)
            end if 
            Return Nothing 
        End Function

    End Module
End Namespace
