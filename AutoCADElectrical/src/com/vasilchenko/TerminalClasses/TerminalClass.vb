Namespace com.vasilchenko.TerminalClasses
    Public Class TerminalClass
        Inherits TerminalAccessoriesClass

        Private intTERM As Integer
        Private objWiresLeft As ArrayList
        Private objWiresRigth As ArrayList
        Private objWireLeft As WireClass
        Private objWireRigth As WireClass
        Private strHDL As String
        Private lngWIDTH As Long

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

        Public Property WIDTH As Long
            Get
                Return lngWIDTH
            End Get
            Set(value As Long)
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

    End Class

End Namespace