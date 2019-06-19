Imports AutoCADElectrical.com.vasilchenko.TerminalEnums

Namespace com.vasilchenko.TerminalClasses
    Public Class JumperClass

        Public Property Jumper As TerminalClass

        Public Property StartTermNum As Integer

        Public Property Side As SideEnum

        Public Property TermCount As Integer
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace

