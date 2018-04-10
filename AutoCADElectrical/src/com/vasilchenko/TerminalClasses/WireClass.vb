Namespace com.vasilchenko.TerminalClasses
    Public Class WireClass
        Private strWIRENO As String
        Private strINST As String
        Private strNAM As String
        Private strPIN As String
        Private strTERMDESC As String
        Private objCABLE As CableClass

        Public Sub New()
            Me.objCABLE = New CableClass
        End Sub

        Public ReadOnly Property TermDesc() As String
            Get
                Me.strTERMDESC = strNAM & ":" & strPIN
                TermDesc = Me.strTERMDESC
            End Get
        End Property

        Public Property Wireno As String
            Get
                Return Me.strWIRENO
            End Get
            Set(strWirenoValue As String)
                Me.strWIRENO = strWirenoValue
            End Set
        End Property
        Public Property Instance As String
            Get
                Return Me.strINST
            End Get
            Set(strInstanceValue As String)
                Me.strINST = strInstanceValue
            End Set
        End Property
        Public Property Name As String
            Get
                Return Me.strNAM
            End Get
            Set(strNameValue As String)
                Me.strNAM = strNameValue
            End Set
        End Property
        Public Property Pin As String
            Get
                Return Me.strPIN
            End Get
            Set(strPinValue As String)
                Me.strPIN = strPinValue
            End Set
        End Property
        Public Property Cable As CableClass
            Get
                Return Me.objCABLE
            End Get
            Set(objCableValue As CableClass)
                Me.objCABLE = objCableValue
            End Set
        End Property

        Public Function HasCable() As Boolean
            HasCable = IsNothing(Me.objCABLE)
        End Function
        Protected Overrides Sub Finalize()
            Me.objCABLE = Nothing
            MyBase.Finalize()
        End Sub

    End Class

End Namespace