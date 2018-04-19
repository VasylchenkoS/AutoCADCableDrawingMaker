Namespace com.vasilchenko.TerminalClasses
    Public Class TerminalAccessoriesClass

        Private strP_TAGSTRIP As String
        Private strINST As String
        Private strLOC As String
        Private strMFG As String
        Private strCAT As String
        Private strWDBLKNAM As String
        Private strBLOCK As String
        Private lngHEIGHT As Double

        Public Property P_TAGSTRIP As String
            Get
                Return strP_TAGSTRIP
            End Get
            Set(value As String)
                strP_TAGSTRIP = value
            End Set
        End Property

        Public Property INST As String
            Get
                Return strINST
            End Get
            Set(value As String)
                strINST = value
            End Set
        End Property

        Public Property LOC As String
            Get
                Return strLOC
            End Get
            Set(value As String)
                strLOC = value
            End Set
        End Property

        Public Property MFG As String
            Get
                Return strMFG
            End Get
            Set(value As String)
                strMFG = value
            End Set
        End Property

        Public Property CAT As String
            Get
                Return strCAT
            End Get
            Set(value As String)
                strCAT = value
            End Set
        End Property

        Public Property WDBLKNAM As String
            Get
                Return strWDBLKNAM
            End Get
            Set(value As String)
                strWDBLKNAM = value
            End Set
        End Property

        Public Property BLOCK As String
            Get
                Return strBLOCK
            End Get
            Set(value As String)
                strBLOCK = value
            End Set
        End Property

        Public Property HEIGHT As Double
            Get
                Return lngHEIGHT
            End Get
            Set(value As Double)
                lngHEIGHT = value
            End Set
        End Property

        Public Sub New()
            strWDBLKNAM = "TRMS"
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

End Namespace