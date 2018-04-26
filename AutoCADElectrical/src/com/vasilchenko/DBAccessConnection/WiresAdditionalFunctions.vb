Imports System.Text.RegularExpressions
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.DBAccessConnection
    Module WiresAdditionalFunctions
        Public Sub IdentityTags(ByVal objInputDictionary As TerminalDictionaryClass(Of String, ArrayList))
            Dim objConnectionList As ArrayList
            Dim objTempWireI As WireClass
            Dim objTempWireY As WireClass

            For intA As Integer = 0 To objInputDictionary.Count - 1
                objConnectionList = objInputDictionary.Item(intA)
                For intI As Integer = 0 To objConnectionList.Count - 2
                    objTempWireI = objConnectionList.Item(intI)
                    For intY As Integer = intI + 1 To objConnectionList.Count - 1
                        objTempWireY = objConnectionList.Item(intY)

                        If objTempWireI.Instance = objTempWireY.Instance AndAlso GetNumericFromString(objTempWireI.Name) <> -1 Then
                            If Replace(objTempWireI.Name, GetNumericFromString(objTempWireI.Name), "") =
                                    Replace(objTempWireY.Name, GetNumericFromString(objTempWireY.Name), "") Then

                                objConnectionList.RemoveAt(intY)
                            End If
                        End If
                    Next
                Next
            Next
        End Sub
        Private Function GetNumericFromString(s As String) As Integer
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                Return matches(matches.Count - 1).Value
            Else
                Return -1
            End If
        End Function

        Public Sub SortCollectionByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, ArrayList))
            Dim objConnectionList As ArrayList
            Dim objTempWire As WireClass

            For intA As Integer = 0 To objInputDictionary.Count - 1
                objConnectionList = objInputDictionary.Item(intA)
                For intY As Integer = 0 To objConnectionList.Count - 2
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

        Public Sub SortDictionaryByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, ArrayList), strPanelLocation As String)
            Dim objConnectionListF As ArrayList
            Dim objConnectionListS As ArrayList
            Dim objTempList As ArrayList

            Dim blnKF, blnKS, blnCF, blnCS As Boolean

            For lngA As Integer = 0 To objInputDictionary.Count - 2
                blnKF = False : blnKS = False : blnCF = False : blnCS = False
                objConnectionListF = objInputDictionary.Item(lngA)
                objConnectionListS = objInputDictionary.Item(lngA + 1)
                For lngI = 1 To objConnectionListF.Count - 1
                    If objConnectionListF.Item(lngI).HasCable Then
                        If objConnectionListF.Item(lngI).Cable.Location <> strPanelLocation Then
                            blnKF = True
                        Else : blnCF = True
                        End If
                    End If
                Next
                For lngI As Long = 1 To objConnectionListS.Count - 1
                    If objConnectionListS.Item(lngI).HasCable Then
                        If objConnectionListS.Item(lngI).Cable.Location <> strPanelLocation Then
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

        Public Function CableInList(objInputList As ArrayList) As Integer
            For lngA As Long = 0 To objInputList.Count - 1
                If objInputList.Item(lngA).HasCable Then
                    CableInList = lngA
                    Exit Function
                End If
            Next
            CableInList = -1
        End Function
    End Module
End Namespace