Imports System.ComponentModel

Namespace com.vasilchenko.TerminalEnums
    Public Enum AccessoriesPhoenixContactEnum
        <Description("ATP-URTK/SP")> PartitionPlateForUT6
        <Description("ATP-UT")> PartitionPlateForUT2_5
        <Description("ATP-UT TWIN")> PartitionPlateForUT2_5MT
        <Description("CLIPFIX 35")> EndClamp35
        <Description("CLIPFIX 35-5")> EndClamp35_5
        <Description("D-UT 2,5/10")> EndCoverForUT2_5
        <Description("D-UT 2,5/4-TWIN")> EndCoverForUT2_5MT
        <Description("UBE/D")> TerminalMarker
        <Description("FBS 2-5")> Jumper2
        <Description("FBS 3-5")> Jumper3
        <Description("FBS 4-5")> Jumper4
        <Description("FBS 10-5")> Jumper10
    End Enum

    Public Enum TerminalOrientationEnum
        <Description("Вертикальный")> Vertical
        <Description("Горизонтальный")> Horisontal
    End Enum

    Public Enum DuctSideEnum
        <Description("Справа")> Rigth
        <Description("Слева")> Left
    End Enum

End Namespace