Imports AutoCADCableDrawingMaker.com.vasilchenko.Enums
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.Classes
    Public Class CableClass
        Public Property Manufacture As String
        Public Property CatalogName As String
        Public Property Desc1 As String
        Public Property Desc2 As String
        Public Property Desc3 As String
        Public Property Tag As String
        Public Property Location As String
        Public Property Family As String
        Public Property Instance As String
        Public Property TerminalConnections As List(Of TerminalConnectionClass)
        Public Property BlockPath As String
        Public Property Orientation As OrientationEnum
        Public Property Count As Double
        Public Property InsertionPoint As Point3d
    End Class
End Namespace

