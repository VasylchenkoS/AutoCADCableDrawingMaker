Imports System.Text.RegularExpressions
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.DBAccessConnection
    Module WiresAdditionalFunctions

        Public Sub IdentityTags(ByVal objInputDictionary As TerminalDictionaryClass(Of String, List(Of WireClass)))

            Dim objTempWireI As WireClass
            Dim objTempWireJ As WireClass

            For Each objConnectionList As List(Of WireClass) In objInputDictionary.Items
                For i As Short = objConnectionList.Count - 2 To 1 Step -1
                    objTempWireI = objConnectionList.Item(i)
                    For j As Short = i + 1 To 0 Step -1
                        objTempWireJ = objConnectionList.Item(j)

                        If objTempWireI.Instance = objTempWireJ.Instance AndAlso Not objTempWireI.Termdesc.Equals(objTempWireJ.Termdesc) AndAlso GetNumericFromString(objTempWireI.ConnTag) <> -1 Then
                            If objTempWireI.ConnTag.Length > 3 AndAlso objTempWireJ.ConnTag.Length > 3 AndAlso objTempWireI.ConnTag.Remove(3) = objTempWireJ.ConnTag.Remove(3) Then
                                objConnectionList.RemoveAt(i)
                            End If
                        End If

                    Next
                Next
            Next
        End Sub
        Private Function GetNumericFromString(s As String) As Double
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                Return Math.Abs(Double.Parse(matches(matches.Count - 1).Value))
            Else
                Return -1
            End If
        End Function

        Public Sub SortCollectionByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, List(Of WireClass)))

            Dim objTempWire As WireClass
            Dim shtFalse As Short = 0

            For Each objConnectionList As List(Of WireClass) In objInputDictionary.Items
                If Not objConnectionList.Count = 1 Then
                    For intY As Short = 0 To objConnectionList.Count - 2
                        If GetNumericFromString(objConnectionList.Item(intY).Instance) = -1 OrElse
                            (GetNumericFromString(objConnectionList.Item(intY).Instance) > GetNumericFromString(objConnectionList.Item(intY + 1).Instance) And GetNumericFromString(objConnectionList.Item(intY + 1).Instance) <> -1) Then
                            objTempWire = objConnectionList(intY)
                            objConnectionList.RemoveAt(intY)
                            objConnectionList.Insert(intY + 1, objTempWire)
                            intY = -1
                        ElseIf objConnectionList.Item(intY).HasCable Then
                            objTempWire = objConnectionList(intY)
                            objConnectionList.RemoveAt(intY)
                            objConnectionList.Insert(intY + 1, objTempWire)
                            intY = -1
                        ElseIf objConnectionList.Item(intY + 1).HasCable And (intY + 2) < objConnectionList.Count Then
                            objTempWire = objConnectionList(intY + 1)
                            objConnectionList.RemoveAt(intY + 1)
                            objConnectionList.Insert(intY + 2, objTempWire)
                            intY = -1
                        End If
                        shtFalse += 1
                        If shtFalse = 20 Then
                            MsgBox("В клемме ошибка. Измените подключение " & objConnectionList.Item(intY + 1).Termdesc & ":" & objConnectionList.Item(intY + 1).WireNumber)
                            Exit Sub
                        End If
                    Next
                End If
            Next
        End Sub

        Public Sub SortDictionaryByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, List(Of WireClass)), strPanelLocation As String)
            Dim objConnectionListF As List(Of WireClass)
            Dim objConnectionListS As List(Of WireClass)
            Dim objTempList As List(Of WireClass)
            Dim blnChange = False


            For lngA As Short = 0 To objInputDictionary.Count - 2
                objConnectionListF = objInputDictionary.Item(lngA)
                objConnectionListS = objInputDictionary.Item(lngA + 1)
                If objConnectionListF.Min(Function(x) GetNumericFromString(x.Instance)) = -1 OrElse
                    objConnectionListF.Min(Function(x) GetNumericFromString(x.Instance)) > objConnectionListS.Min(Function(x) GetNumericFromString(x.Instance)) Then
                    blnChange = True
                End If
                If objConnectionListF.Any(Function(x)
                                              If x.HasCable AndAlso (x.Cable.Destination <> "" And x.Cable.Destination <> strPanelLocation) Then
                                                  Return True
                                              Else : Return False
                                              End If
                                          End Function) Then
                    blnChange = True
                End If
                If blnChange Then
                    objTempList = objInputDictionary.Item(lngA)
                    objInputDictionary.RemoveAt(lngA)
                    objInputDictionary.Add(key:=objTempList.Item(0).WireNumber, value:=objTempList)
                End If
            Next
        End Sub

        Public Function CableInList(objInputList As List(Of WireClass)) As Integer
            For lngA As Short = 0 To objInputList.Count - 1
                If objInputList.Item(lngA).HasCable Then
                    CableInList = lngA
                    Exit Function
                End If
            Next
            CableInList = -1
        End Function
    End Module
End Namespace