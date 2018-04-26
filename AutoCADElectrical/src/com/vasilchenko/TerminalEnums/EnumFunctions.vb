Imports System.ComponentModel
Imports System.Reflection
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.TerminalEnums
    Module EnumFunctions
        Public Function GetEnumFromDescriptionAttribute(Of T)(description As String) As T
            Dim type = GetType(T)
            If Not type.IsEnum Then
                Throw New InvalidOperationException()
            End If
            For Each fieldInfo As FieldInfo In type.GetFields()
                Dim descriptionAttribute = TryCast(Attribute.GetCustomAttribute(fieldInfo, GetType(DescriptionAttribute)), DescriptionAttribute)
                If descriptionAttribute IsNot Nothing Then
                    If descriptionAttribute.Description <> description Then
                        Continue For
                    End If
                    Return DirectCast(fieldInfo.GetValue(Nothing), T)
                End If
                If fieldInfo.Name <> description Then
                    Continue For
                End If
                Return DirectCast(fieldInfo.GetValue(Nothing), T)
            Next
            Return Nothing
        End Function
        Public Function Convert(eAccEnum As AccessoriesPhoenixContactEnum, objCurTerminal As TerminalClass) As TerminalAccessoriesClass
            Dim objTerminalAcc As New TerminalAccessoriesClass

            objTerminalAcc.P_TAGSTRIP = objCurTerminal.P_TAGSTRIP
            objTerminalAcc.INST = objCurTerminal.INST
            objTerminalAcc.LOC = objCurTerminal.LOC
            objTerminalAcc.MFG = "Phoenix Contact"
            objTerminalAcc.CAT = New EnumDescriptor(Of AccessoriesPhoenixContactEnum)(eAccEnum).ToString
            DBDataAccessObject.FillTerminalBlockPath(objTerminalAcc)
            Return objTerminalAcc
        End Function
    End Module
End Namespace

