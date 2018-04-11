Imports AutoCADElectrical.com.vasilchenko.TerminalEnums
Imports AutoCADElectrical.com.vasilchenko.DBAccessConnection
Imports AutoCADElectrical.com.vasilchenko.TerminalClasses

Namespace com.vasilchenko.TerminalModules
    Module TerminalClassMapping

        Public Sub CreateTerminalBlock(strLocation As String, strTagstrip As String,
                                    eOrientation As TerminalOrientationEnum, eDucktSide As DuctSideEnum)
            Dim objTermsArray As ArrayList
            Dim objMappingTermsCollection As ArrayList

            objTermsArray = GetAllTermsInLocation(strLocation, strTagstrip)
            objMappingTermsCollection = FillTerminalData(objTermsArray, strTagstrip, eDucktSide)

        End Sub

        Private Function FillTerminalData(objInputList As ArrayList, strTagstrip As String, eDucktSide As DuctSideEnum) As ArrayList
            Dim objResultArray As New ArrayList

            For Each strTempItem As String In objInputList
                objResultArray.Add(FillTermTypeData(strTagstrip, strTempItem))
            Next
            For Each objTempItem As TerminalClass In objResultArray
                FillTerminalBlockPath(objTempItem)
                FillTerminalConnectionsData(objTempItem, eDucktSide)
            Next

            Return objResultArray

        End Function

    End Module

End Namespace
