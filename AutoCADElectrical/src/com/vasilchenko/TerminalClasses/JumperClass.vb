Imports AutoCADElectrical.com.vasilchenko.TerminalEnums

Namespace com.vasilchenko.TerminalClasses
    Public Class JumperClass
        Private objJumper As TerminalAccessoriesClass
        Private intStartTermNum As Integer
        Private intTermCount As Integer
        Private eSide As SideEnum

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Public Property Jumper As TerminalAccessoriesClass
            Get
                Return objJumper
            End Get
            Set(value As TerminalAccessoriesClass)
                Me.objJumper = value
            End Set
        End Property

        Public Property StartTermNum As Integer
            Get
                Return intStartTermNum
            End Get
            Set(value As Integer)
                Me.intStartTermNum = value
            End Set
        End Property

        Public Property Side As SideEnum
            Get
                Return eSide
            End Get
            Set(value As SideEnum)
                Me.eSide = value
            End Set
        End Property

        Public Property TermCount As Integer
            Get
                Return intTermCount
            End Get
            Set(value As Integer)
                Me.intTermCount = value
            End Set
        End Property
    End Class
End Namespace

