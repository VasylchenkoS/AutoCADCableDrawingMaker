﻿Namespace com.vasilchenko.TerminalClasses
    Public Class TerminalClass
        Inherits TerminalAccessoriesClass
        Implements IEquatable(Of TerminalClass)
        Implements IComparable(Of TerminalClass)

        Private intTERM As Integer
        Private objWiresLeft As ArrayList
        Private objWiresRigth As ArrayList
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

        Public WriteOnly Property WireLeft As WireClass
            Set(value As WireClass)
                objWiresLeft.Add(value)
            End Set
        End Property

        Public WriteOnly Property WireRigth As WireClass
            Set(value As WireClass)
                objWiresRigth.Add(value)
            End Set
        End Property

        Public ReadOnly Property WiresLeftList As ArrayList
            Get
                Return objWiresLeft
            End Get
        End Property

        Public ReadOnly Property WiresRigthList As ArrayList
            Get
                Return objWiresRigth
            End Get
        End Property

        Public Sub New()
            Me.objWiresLeft = New ArrayList
            Me.objWiresRigth = New ArrayList
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
    End Class

End Namespace