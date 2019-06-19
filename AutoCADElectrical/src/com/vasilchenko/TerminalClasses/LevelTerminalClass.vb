Namespace com.vasilchenko.TerminalClasses
    Public Class LevelTerminalClass
        Implements IEquatable(Of LevelTerminalClass)
        Implements IComparable(Of LevelTerminalClass)

        Private _objWiresLeftList As List(Of WireClass)
        Private _objWiresRigthList As List(Of WireClass)
        Public Property TerminalNumber As Short
        Public Property Handle As String
        Public Property Level As Short

        Public WriteOnly Property AddWireLeft As WireClass
            Set(value As WireClass)
                _objWiresLeftList.Add(value)
            End Set
        End Property

        Public WriteOnly Property AddWireRigth As WireClass
            Set(value As WireClass)
                _objWiresRigthList.Add(value)
            End Set
        End Property

        Public Property WiresLeftListList As List(Of WireClass)
            Get
                Return _objWiresLeftList
            End Get
            Set(value As List(Of WireClass))
                _objWiresLeftList = value
            End Set
        End Property

        Public Property WiresRigthListList As List(Of WireClass)
            Get
                Return _objWiresRigthList
            End Get
            Set(value As List(Of WireClass))
                _objWiresRigthList = value
            End Set
        End Property

        Public Sub New()
            Me._objWiresLeftList = New List(Of WireClass)
            Me._objWiresRigthList = New List(Of WireClass)
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Public Overloads Function Equals(other As LevelTerminalClass) As Boolean Implements IEquatable(Of LevelTerminalClass).Equals
            If other Is Nothing Then
                Return 1
            Else
                Return Me.TerminalNumber.Equals(other.TerminalNumber)
            End If
        End Function

        Public Function CompareTo(other As LevelTerminalClass) As Integer Implements IComparable(Of LevelTerminalClass).CompareTo
            If other Is Nothing Then
                Return 1
            Else
                Return Me.TerminalNumber.CompareTo(other.TerminalNumber)
            End If
        End Function

        Friend Function WireNameStartsWith(strWireName As String) As Boolean
            Return Me._objWiresLeftList.Any(Function(x) x.WireNumber.StartsWith(strWireName)) Or Me._objWiresRigthList.Any(Function(x) x.WireNumber.StartsWith(strWireName))
        End Function
        Friend Function WireContains(strWireName As String) As Boolean
            Return Me._objWiresLeftList.Any(Function(x) x.WireNumber.Contains(strWireName)) Or Me._objWiresRigthList.Any(Function(x) x.WireNumber.Contains(strWireName))
        End Function
    End Class

End Namespace