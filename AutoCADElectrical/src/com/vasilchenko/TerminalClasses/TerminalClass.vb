Namespace com.vasilchenko.TerminalClasses
    Public Class TerminalClass
        Public Property BottomLevelTerminal As LevelTerminalClass
        Public Property TopLevelTerminal As LevelTerminalClass
        Public Property MainTermNumber As Short
        Public Property WdBlockName As String
        Public Property BlockPath As String
        Public Property TagStrip As String
        Public Property Instance As String
        Public Property Location As String
        Public Property Manufacture As String
        Public Property Catalog As String
        Public Property Width As Double
        Public Property Height As Double
        Public Property Count As Double
        Public Property LinkTerm As String


        Public Sub New()
            WdBlockName = "TRMS"
        End Sub

        Protected Overrides Sub Finalize()
            Me.BottomLevelTerminal = Nothing
            Me.TopLevelTerminal = Nothing
            MyBase.Finalize()
        End Sub

    End Class
End Namespace
