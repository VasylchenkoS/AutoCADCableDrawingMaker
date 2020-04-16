Imports System.Text.RegularExpressions

Namespace com.vasilchenko.Modules
    Module AdditionalFunctions
        Friend Function GetLastNumericFromString(s As String) As Double

            If s.Equals("") Then Return -1

            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection
            Try
                matches = rgx.Matches(s)
                If matches.Count > 0 Then
                    Return Math.Abs(Double.Parse(matches(matches.Count - 1).Value))
                Else
                    Return -1
                End If
            Catch ex As Exception
                MsgBox("Что-то не так: " & vbCrLf & ex.Message)
            End Try
        End Function

    End Module
End Namespace