Imports System.Text.RegularExpressions
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.DBAccessConnection
    Module WiresSortFunctions
        Public Sub IdentityTags(ByVal objInputDictionary As Dictionary(Of String, ArrayList))
            Dim objConnectionList As ArrayList
            Dim vntTempWire1 As WireClass
            Dim vntTempWire2 As WireClass


            For Each pair As KeyValuePair(Of String, ArrayList) In objInputDictionary
                objConnectionList = pair.Value
                For lngI As Long = 0 To objConnectionList.Count - 2
                    vntTempWire1 = objConnectionList.Item(lngI)
                    For lngY As Long = lngI + 1 To objConnectionList.Count - 1
                        vntTempWire2 = objConnectionList.Item(lngY)
                        If vntTempWire1.Instance = vntTempWire2.Instance Then
                            If Strings.Replace(vntTempWire1.Name, GetNumericFromString(vntTempWire1.Name), "") =
                                    Replace(vntTempWire2.Name, GetNumericFromString(vntTempWire2.Name), "") Then
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

    End Module
End Namespace