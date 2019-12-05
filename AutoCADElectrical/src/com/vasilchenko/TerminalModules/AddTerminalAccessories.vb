Imports AutoCADElectrical.com.vasilchenko.TerminalClasses
Imports AutoCADElectrical.com.vasilchenko.TerminalEnums

Namespace com.vasilchenko.TerminalModules
    Module AddTerminalAccessories
        Private ReadOnly UncoverObjects As New List(Of String) From {"UT 4-MT", "AVK 2,5/4T Yellow-Green"}

        Friend Sub AddAccessoriesForSignalisation(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim intCount As Integer = objMappingTermsList.Count - 2
            Dim index As Integer = 0

            objMappingTermsList.Insert(index, AddTerminalMarker(objMappingTermsList.Item(index)))
            objMappingTermsList.Insert(index + 1, GetCoverObject(objMappingTermsList.Item(index + 1)))
            index += 2
            intCount += 2

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber = 8 Or objCurTerminal.MainTermNumber = 24 Then
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
            Dim index = 0


            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)

                If objCurTerminal.MainTermNumber = 1 Then
                    objMappingTermsList.Insert(index, AddTerminalMarker(objMappingTermsList.Item(index)))
                    objMappingTermsList.Insert(index + 1, GetCoverObject(objMappingTermsList.Item(index + 1)))
                    index += 2
                    intCount += 2
                ElseIf _
                    (objCurTerminal.MainTermNumber = 2 Or objCurTerminal.MainTermNumber = 4) And
                    objCurTerminal.Catalog = "UT 6-HESI (6,3X32)" Then
                    objMappingTermsList.Insert(index + 1, GetPlateObject(objCurTerminal))
                    index += 1
                    intCount += 1
                ElseIf objCurTerminal.Catalog <> "UT 6-HESI (6,3X32)" Then
                    If _
                        ((objPrevTerminal.BottomLevelTerminal.WireContains("L") Or
                          objPrevTerminal.BottomLevelTerminal.WireContains("PP") Or
                          objPrevTerminal.BottomLevelTerminal.WireContains("+") Or
                          objPrevTerminal.BottomLevelTerminal.WireContains("PV")) And
                         (objCurTerminal.BottomLevelTerminal.WireContains("N") Or
                          objCurTerminal.BottomLevelTerminal.WireContains("PM") Or
                          objCurTerminal.BottomLevelTerminal.WireContains("-") Or
                          objCurTerminal.BottomLevelTerminal.WireContains("PT"))) Then
                        objMappingTermsList.Insert(index, GetCoverObject(objCurTerminal))
                        index += 1
                        intCount += 1
                    ElseIf _
                        (objPrevTerminal.BottomLevelTerminal.WireContains("N") Or
                         objPrevTerminal.BottomLevelTerminal.WireContains("PM") Or
                         objPrevTerminal.BottomLevelTerminal.WireContains("-") Or
                         objPrevTerminal.BottomLevelTerminal.WireContains("PT")) And
                        (objCurTerminal.BottomLevelTerminal.WireContains("L") Or
                         objCurTerminal.BottomLevelTerminal.WireContains("PP") Or
                         objCurTerminal.BottomLevelTerminal.WireContains("+") Or
                         objCurTerminal.BottomLevelTerminal.WireContains("PV")) Then
                        objMappingTermsList.Insert(index, GetPlateObject(objCurTerminal))
                        index += 1
                        intCount += 1
                    ElseIf _
                        objCurTerminal.BottomLevelTerminal.WireNameStartsWith("PE") And
                        Not objPrevTerminal.BottomLevelTerminal.WireContains("PE") Then
                        objMappingTermsList.Insert(index, GetPlateObject(objPrevTerminal))
                        index += 1
                        intCount += 1
                    End If
                End If
                index += 1

                objPrevTerminal = objCurTerminal
            Loop
            If Not UncoverObjects.Contains(objPrevTerminal.Catalog) Then _
                objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Friend Sub AddAccessoriesForControl(objMappingTermsList As List(Of TerminalClass))
            If objMappingTermsList.Count = 0 Then Exit Sub
            Dim objCurTerminal As TerminalClass = Nothing
            Dim intCount As Integer = objMappingTermsList.Count - 2
            Dim index As Integer = 0
            Dim shtCount As Short = 0

            Try
                shtCount = cint(InputBox("Сколько проводов в схеме ТУ?", "ControlCount", "5", 100, 100))
            Catch
                MsgBox("Нужно было вводить целое число) Установлено по-умолчанию (5)")
                shtCount = 5
            End Try

            objMappingTermsList.Insert(index, AddTerminalMarker(objMappingTermsList.Item(index)))
            objMappingTermsList.Insert(index + 1, GetCoverObject(objMappingTermsList.Item(index + 1)))
            index += 2
            intCount += 2

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber Mod shtCount = 0 Then
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
            Dim shtCount As Short = 0

            objMappingTermsList.Insert(index, AddTerminalMarker(objMappingTermsList.Item(index)))
            objMappingTermsList.Insert(index + 1, GetCoverObject(objMappingTermsList.Item(index + 1)))
            index += 2
            intCount += 2

            Dim msgResult = MsgBox("Модуль 8-миканальный (560AIRxx) ?", MsgBoxStyle.YesNo)
            Select Case msgResult
                Case MsgBoxResult.Yes
                    shtCount = 8
                Case MsgBoxResult.No
                    shtCount = 6
            End Select

            Do While index <> intCount
                objCurTerminal = objMappingTermsList.Item(index)
                If objCurTerminal.MainTermNumber Mod shtCount*2 = 0 Then
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
            Dim blnCurWirenum As Boolean
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
                    If Not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then
                        objMappingTermsList.Insert(index + 1, GetCoverObject(objCurTerminal))
                        index += 1
                    End If
                    index += 1
                ElseIf blnCurWirenum And blnIsTt Then
                    If Not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then
                        objMappingTermsList.Insert(index, GetCoverObject(objCurTerminal))
                        objMappingTermsList.Insert(index + 1, AddTerminalMarker(objCurTerminal))
                        objMappingTermsList.Insert(index + 2, GetCoverObject(objCurTerminal))
                        index += 3
                    Else
                        objMappingTermsList.Insert(index, AddTerminalMarker(objCurTerminal))
                        index += 1
                    End If
                ElseIf blnPrevWirenum And blnCurWirenum And Not blnIsTt Then
                    If Not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then
                        objMappingTermsList.Insert(index, GetCoverObject(objCurTerminal))
                        index += 1
                    End If
                    blnPrevWirenum = False
                End If
                index += 1

                If _
                    objCurTerminal.BottomLevelTerminal.WiresLeftListList.Count <> 0 Or
                    objCurTerminal.BottomLevelTerminal.WiresRigthListList.Count <> 0 Then
                    blnPrevWirenum = objCurTerminal.BottomLevelTerminal.WireNameStartsWith("N") Or
                                     objCurTerminal.BottomLevelTerminal.WireNameStartsWith("O")
                End If
            Loop
            If Not objCurTerminal.Catalog.ToLower.StartsWith("wgo 4") Then _
                objMappingTermsList.Add(GetCoverObject(objCurTerminal))
            objMappingTermsList.Add(AddEndClamp(objCurTerminal))
        End Sub

        Private Function GetCoverObject(objCurTerminal As TerminalClass) As TerminalClass
            Select Case objCurTerminal.Catalog
                Case "UT 2,5", "UT 2,5-PE", "UT 4", "UT 6", "UT 6-PE"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForUT2_5, objCurTerminal)
                Case "UT 16"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForUT16, objCurTerminal)
                Case "UT 2,5-MT", "UT 4-MTD-DIO/R-L"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForUT2_5MT, objCurTerminal)
                Case "UTTB 2,5", "UTTB 2,5-PV", "UTTB 2,5-MT-P/P"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForUTTB2_5, objCurTerminal)
                Case "UT 2,5-QUATTRO"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForUT2_5QUATTRO, objCurTerminal)
                Case "URTK 6"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForURTK_6, objCurTerminal)
                Case "UT 6-HESI (6,3X32)"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUT6, objCurTerminal)
                Case "ST 2,5"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForST2_5, objCurTerminal)
                Case "STTB 2,5-L/N", "STTB 2,5-PE"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForSTTB2_5, objCurTerminal)
                Case "AVK 4 A Gray", "AVK 2,5 A Gray", "AVK 2,5 R Gray"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.K_EndCoverForAVK2_5_R_AVK_4_A, objCurTerminal)
                Case "AVK 2,5 Gray", "AVK 4 Gray", "AVK 6 Gray", "AVK 10 Gray"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.K_EndCoverForForAVK2_5_10, objCurTerminal)
                Case "WDU 2.5", "WPE 2.5"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.K_EndCoverForForAVK2_5_10, objCurTerminal)
                Case "280-601", "280-607"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.WAGO_EndCoverFor280_601_7, objCurTerminal)
                Case "280-833"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.WAGO_EndCoverFor280_833, objCurTerminal)
                Case Else
                    Return Nothing
            End Select
        End Function

        Private Function GetPlateObject(objCurTerminal As TerminalClass) As TerminalClass

            If objCurTerminal.Manufacture = "Klemsan" Then Return GetCoverObject(objCurTerminal)

            Select Case objCurTerminal.Catalog
                Case "UT 2,5", "UT 2,5-PE", "UT 4", "UT 6", "UT 6-PE"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUT2_5, objCurTerminal)
                Case "UT 16"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUT16, objCurTerminal)
                Case "UT 2,5-MT", "UT 4-MTD-DIO/R-L"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUT2_5MT, objCurTerminal)
                Case "UTTB 2,5", "UTTB 2,5-PV", "UTTB 2,5-MT-P/P"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUTTB2_5, objCurTerminal)
                Case "UT 2,5-QUATTRO"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUT2_5QUATTRO, objCurTerminal)
                Case "UT 6-HESI (6,3X32)"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForUT6, objCurTerminal)
                Case "ST 2,5"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_PartitionPlateForST4, objCurTerminal)
                Case "STTB 2,5-L/N", "STTB 2,5-PE"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndCoverForSTTB2_5, objCurTerminal)
                Case "WDU 2.5", "WPE 2.5"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.WM_EndCoverForForWDU2_5, objCurTerminal)
                Case "280-601", "280-607"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.WAGO_EndCoverFor280_601_7, objCurTerminal)
                Case "280-833"
                    Return EnumFunctions.Convert(TerminalAccessoriesEnum.WAGO_PartitionPlateFor280_833, objCurTerminal)
                Case Else
                    Return Nothing
            End Select
        End Function

        Private Function AddEndClamp(objCurTerminal As TerminalClass) As TerminalClass
            If objCurTerminal.Manufacture.ToLower.Equals("phoenix contact") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_EndClamp35, objCurTerminal)
            ElseIf objCurTerminal.Manufacture.ToLower.Equals("klemsan") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.K_EndClamp, objCurTerminal)
            ElseIf objCurTerminal.Manufacture.ToLower.Equals("weidmuller") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.WM_EndClamp, objCurTerminal)
            ElseIf objCurTerminal.Manufacture.ToLower.Equals("wago") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.WAGO_EndClamp, objCurTerminal)
            End If
            Return Nothing
        End Function

        Private Function AddTerminalMarker(objCurTerminal As TerminalClass) As TerminalClass
            If objCurTerminal.Manufacture.ToLower.Equals("phoenix contact") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.PC_TerminalMarker, objCurTerminal)
            ElseIf objCurTerminal.Manufacture.ToLower.Equals("klemsan") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.K_TerminalMarker, objCurTerminal)
            ElseIf objCurTerminal.Manufacture.ToLower.Equals("weidmuller") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.WM_TerminalMarker, objCurTerminal)
            ElseIf objCurTerminal.Manufacture.ToLower.Equals("wago") Then
                Return EnumFunctions.Convert(TerminalAccessoriesEnum.WAGO_TerminalMarker, objCurTerminal)
            End If
            Return Nothing
        End Function
    End Module
End Namespace
