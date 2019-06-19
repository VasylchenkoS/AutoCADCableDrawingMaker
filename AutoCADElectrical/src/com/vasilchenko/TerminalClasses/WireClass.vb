﻿Namespace com.vasilchenko.TerminalClasses
    Public Class WireClass
        Private _strTermdesc As String

        Public Sub New()
            Me.Cable = New CableClass
        End Sub

        Public Property WireNumber As String

        Public Property Instance As String

        Public Property ConnTag As String

        Public Property ConnPin As String

        Public Property Cable As CableClass

        Public Property Termdesc As String
            Get
                If Me._strTermdesc.Equals("") And ConnTag <> "" And ConnPin <> "" Then
                    Me._strTermdesc = ConnTag & ":" & ConnPin
                End If
                Return _strTermdesc
            End Get
            Set(value As String)
                Me._strTermdesc = value
            End Set
        End Property

        Public Function HasCable() As Boolean
            HasCable = Not IsNothing(Me.Cable)
        End Function
    End Class

End Namespace