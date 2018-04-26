Namespace com.vasilchenko.TerminalClasses

    Public Class TerminalStripClass
        Private objTerminalList As List(Of TerminalAccessoriesClass)
        Private objJumperList As List(Of JumperClass)

        Public Property TerminalList As List(Of TerminalAccessoriesClass)
            Get
                Return objTerminalList
            End Get
            Set(value As List(Of TerminalAccessoriesClass))
                Me.objTerminalList = value
            End Set
        End Property

        Public Property JumperList As List(Of JumperClass)
            Get
                Return objJumperList
            End Get
            Set(value As List(Of JumperClass))
                Me.objJumperList = value
            End Set
        End Property
    End Class
End Namespace

