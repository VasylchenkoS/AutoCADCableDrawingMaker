Imports System.Linq
Imports System.Text.RegularExpressions
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.DBAccessConnection
    Module WiresSortFunctions
        Public Sub IdentityTags(ByVal objInputDictionary As TerminalDictionaryClass(Of String, ArrayList))
            Dim objConnectionList As ArrayList
            Dim vntTempWireI As WireClass
            Dim vntTempWireY As WireClass

            For lngA As Long = 0 To objInputDictionary.Count - 1
                objConnectionList = objInputDictionary.Item(lngA)
                For lngI As Long = 0 To objConnectionList.Count - 2
                    vntTempWireI = objConnectionList.Item(lngI)
                    For lngY As Long = lngI + 1 To objConnectionList.Count - 1
                        vntTempWireY = objConnectionList.Item(lngY)
                        If vntTempWireI.Instance = vntTempWireY.Instance Then
                            If Replace(vntTempWireI.Name, GetNumericFromString(vntTempWireI.Name), "") =
                                    Replace(vntTempWireY.Name, GetNumericFromString(vntTempWireY.Name), "") Then
                                objConnectionList.RemoveAt(lngY)
                            End If
                        End If
                    Next
                Next
            Next
        End Sub
        Private Function GetNumericFromString(s As String) As Integer
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                Return matches(matches.Count - 1).Value
            End If
        End Function

        Public Sub SortCollectionByInstAndCables(ByRef objInputDictionary As TerminalDictionaryClass(Of String, ArrayList))
            Dim objConnectionList As ArrayList
            Dim vntTempWire As WireClass

            For lngA As Long = 0 To objInputDictionary.Count - 1
                objConnectionList = objInputDictionary.Item(lngA)
                For lngY As Long = 0 To objConnectionList.Count - 2
                    If GetNumericFromString(objConnectionList.Item(lngY).Instance) > GetNumericFromString(objConnectionList.Item(lngY + 1).Instance) Then
                        vntTempWire = objConnectionList(lngY)
                        objConnectionList.RemoveAt(lngY)
                        objConnectionList.Insert(lngY + 1, vntTempWire)
                    ElseIf objConnectionList.Item(lngY).HasCable Then
                        vntTempWire = objConnectionList(lngY)
                        objConnectionList.RemoveAt(lngY)
                        objConnectionList.Insert(lngY + 1, vntTempWire)
                    ElseIf objConnectionList.Item(lngY + 1).HasCable And (lngY + 2) < objConnectionList.Count Then
                        vntTempWire = objConnectionList(lngY + 1)
                        objConnectionList.RemoveAt(lngY + 1)
                        objConnectionList.Insert(lngY + 2, vntTempWire)
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

        Public Function CableInDictionary(objInputDictionary As TerminalDictionaryClass(Of String, ArrayList), strCurrentLocation As String) As Integer
            Dim objTempList As ArrayList
            For lngI As Long = 0 To objInputDictionary.Count - 1
                objTempList = objInputDictionary.Item(lngI)
                For lngA As Long = 0 To objInputDictionary.Item(lngI).Count
                    If objTempList.Item(lngA).HasCable Then
                        If objTempList.Item(lngA).Cable.Location <> strCurrentLocation Then
                            CableInDictionary = lngI
                            Exit Function
                        End If
                    End If
                Next
            Next
        End Function

        Public Function CableInList(objInputList As ArrayList) As Integer
            For lngA As Long = 0 To objInputList.Count
                If objInputList.Item(lngA).HasCable Then
                    CableInList = lngA
                    Exit Function
                End If
            Next
        End Function

    End Module
End Namespace