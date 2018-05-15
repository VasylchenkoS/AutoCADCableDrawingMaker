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

                        If objTempWireI.Instance = objTempWireJ.Instance AndAlso Not objTempWireI.TERMDESC.Equals(objTempWireJ.TERMDESC) AndAlso GetNumericFromString(objTempWireI.Name) <> -1 Then
                            'If Replace(objTempWireI.Name, GetNumericFromString(objTempWireI.Name), "") =
                            '        Replace(objTempWireJ.Name, GetNumericFromString(objTempWireJ.Name), "") Then
                            If objTempWireI.Name.Length > 3 AndAlso objTempWireJ.Name.Length > 3 AndAlso objTempWireI.Name.Remove(3) = objTempWireJ.Name.Remove(3) Then
                                objConnectionList.RemoveAt(i)
                            End If
                        End If

                    Next
                Next
            Next
        End Sub
        Private Function GetNumericFromString(s As String) As String
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                Return matches(matches.Count - 1).Value
            Else
                Return -1
            End If
        End Function

        Public Sub SortCollectionByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, List(Of WireClass)))

            Dim objTempWire As WireClass

            For Each objConnectionList As List(Of WireClass) In objInputDictionary.Items
                For intY As Short = 0 To objConnectionList.Count - 2
                    If GetNumericFromString(objConnectionList.Item(intY).Instance) > GetNumericFromString(objConnectionList.Item(intY + 1).Instance) Then
                        objTempWire = objConnectionList(intY)
                        objConnectionList.RemoveAt(intY)
                        objConnectionList.Insert(intY + 1, objTempWire)
                    ElseIf objConnectionList.Item(intY).HasCable Then
                        objTempWire = objConnectionList(intY)
                        objConnectionList.RemoveAt(intY)
                        objConnectionList.Insert(intY + 1, objTempWire)
                    ElseIf objConnectionList.Item(intY + 1).HasCable And (intY + 2) < objConnectionList.Count Then
                        objTempWire = objConnectionList(intY + 1)
                        objConnectionList.RemoveAt(intY + 1)
                        objConnectionList.Insert(intY + 2, objTempWire)
                    End If
                Next
            Next
        End Sub

        Public Sub SortDictionaryByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, List(Of WireClass)), strPanelLocation As String)
            Dim objConnectionListF As List(Of WireClass)
            Dim objConnectionListS As List(Of WireClass)
            Dim objTempList As List(Of WireClass)

            Dim blnKF, blnKS, blnCF, blnCS As Boolean

            For lngA As Short = 0 To objInputDictionary.Count - 2
                blnKF = False : blnKS = False : blnCF = False : blnCS = False
                objConnectionListF = objInputDictionary.Item(lngA)
                objConnectionListS = objInputDictionary.Item(lngA + 1)
                For lngI = 1 To objConnectionListF.Count - 1
                    If objConnectionListF.Item(lngI).HasCable Then
                        If objConnectionListF.Item(lngI).Cable.Destination <> strPanelLocation Then
                            blnKF = True
                        Else : blnCF = True
                        End If
                    End If
                Next
                For lngI As Short = 1 To objConnectionListS.Count - 1
                    If objConnectionListS.Item(lngI).HasCable Then
                        If objConnectionListS.Item(lngI).Cable.Destination <> strPanelLocation Then
                            blnKS = True
                        Else : blnCS = True
                        End If
                    End If
                Next
                If blnKF And Not blnKS Or blnCF And Not blnCS Then
                    objTempList = objInputDictionary.Item(lngA)
                    objInputDictionary.RemoveAt(lngA)
                    objInputDictionary.Add(key:=objTempList.Item(0).Wireno, value:=objTempList)
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