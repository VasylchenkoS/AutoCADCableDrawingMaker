Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.TerminalModules
    Module TerminalClassMapping

        Public Sub CreateTerminalBlock(strLocation As String, strTagstrip As String,
                                    eOrientation As TerminalOrientationEnum, eDucktSide As DuctSideEnum)
            Dim objTermsArray As ArrayList
            Dim objMappingTermsCollection As New Collection

            objTermsArray = GetAllTermsInLocation(strLocation, strTagstrip)
            FillTerminalData(objTermsArray, strTagstrip)

        End Sub

        Private Function FillTerminalData(objInputList As ArrayList, strTagstrip As String) As ArrayList
            Dim objResultArray As New ArrayList

            For Each strTempItem As String In objInputList
                objResultArray.Add(FillTermTypeData(strTagstrip, strTempItem))
            Next
            For Each objTempItem As TerminalClass In objResultArray
                objResultArray.Add(FillTerminalBlockPath(objTempItem))
                objResultArray.Add(FillTerminalConnectionsData(objTempItem))
            Next

            Return objResultArray

        End Function

    End Module

End Namespace
