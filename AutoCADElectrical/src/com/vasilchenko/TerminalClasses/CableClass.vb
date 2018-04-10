Namespace com.vasilchenko.TerminalClasses

    Public Class CableClass
        Private strMARK As String
        Private strLOCATION As String
        Public Property Location As String
            Get
                Return Me.strLOCATION
            End Get
            Set(strLocationValue As String)
                Me.strLOCATION = strLocationValue
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