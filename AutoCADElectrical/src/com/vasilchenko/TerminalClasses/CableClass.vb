Namespace com.vasilchenko.TerminalClasses

    Public Class CableClass
        Private strMARK As String
        Private strDESTINATION As String
        Public Property Destination As String
            Get
                Return Me.strDESTINATION
            End Get
            Set(strLocationValue As String)
                Me.strDESTINATION = strLocationValue
            End Set
        End Property
        Public Property Mark As String
            Get
                Return Me.strMARK
            End Get
            Set(strMarkValue As String)
                Me.strMARK = strMarkValue
            End Set
        End Property
    End Class

End Namespace