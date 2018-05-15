Namespace com.vasilchenko.TerminalClasses
    Public Class TerminalClass
        Inherits TerminalAccessoriesClass
        Implements IEquatable(Of TerminalClass)
        Implements IComparable(Of TerminalClass)

        Private intTERM As Integer
        Private objWiresLeft As List(Of WireClass)
        Private objWiresRigth As List(Of WireClass)
        Private objWireLeft As WireClass
        Private objWireRigth As WireClass
        Private strHDL As String
        Private lngWIDTH As Double

        Public Property TERM As Integer
            Get
                Return intTERM
            End Get
            Set(value As Integer)
                intTERM = value
            End Set
        End Property

        Public Property HDL As String
            Get
                Return strHDL
            End Get
            Set(value As String)
                strHDL = value
            End Set
        End Property

        Public Property WIDTH As Double
            Get
                Return lngWIDTH
            End Get
            Set(value As Double)
                lngWIDTH = value
            End Set
        End Property

        Public WriteOnly Property AddWireLeft As WireClass
            Set(value As WireClass)
                objWiresLeft.Add(value)
            End Set
        End Property

        Public WriteOnly Property AddWireRigth As WireClass
            Set(value As WireClass)
                objWiresRigth.Add(value)
            End Set
        End Property

        Public Property WiresLeftList As List(Of WireClass)
            Get
                Return objWiresLeft
            End Get
            Set(value As List(Of WireClass))
                objWiresLeft = value
            End Set
        End Property

        Public Property WiresRigthList As List(Of WireClass)
            Get
                Return objWiresRigth
            End Get
            Set(value As List(Of WireClass))
                objWiresRigth = value
            End Set
        End Property

        Public Sub New()
            Me.objWiresLeft = New List(Of WireClass)
            Me.objWiresRigth = New List(Of WireClass)
        End Sub

        Protected Overrides Sub Finalize()
            Me.objWireLeft = Nothing
            Me.objWireRigth = Nothing
            MyBase.Finalize()
        End Sub

        Public Overloads Function Equals(other As TerminalClass) As Boolean Implements IEquatable(Of TerminalClass).Equals
            If other Is Nothing Then
                Return 1
            Else
                Return Me.intTERM.Equals(other.TERM)
            End If
        End Function

        Public Function CompareTo(other As TerminalClass) As Integer Implements IComparable(Of TerminalClass).CompareTo
            If other Is Nothing Then
                Return 1
            Else
                Return Me.intTERM.CompareTo(other.TERM)
            End If
        End Function

        Friend Function WireNameStartsWith(strWireName As String) As Boolean
            Return Me.objWiresLeft.Any(Function(x) x.Wireno.StartsWith(strWireName)) Or
                Me.objWiresRigth.Any(Function(x) x.Wireno.StartsWith(strWireName))
        End Function

    End Class

End Namespace