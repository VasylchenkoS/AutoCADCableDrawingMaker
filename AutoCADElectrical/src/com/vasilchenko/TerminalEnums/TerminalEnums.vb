Imports System.ComponentModel

Namespace com.vasilchenko.TerminalEnums

    Public Enum TerminaAccessoriesEnum
        'Phoenix Contact
        <Description("ATP-URTK/SP")> PC_PartitionPlateForUT6
        <Description("ATP-UT")> PC_PartitionPlateForUT2_5
        <Description("ATP-UT TWIN")> PC_PartitionPlateForUT2_5MT
        <Description("CLIPFIX 35")> PC_EndClamp35
        <Description("DP-UTTB 2,5/4")> PC_PartitionPlateForUTTB2_5
        <Description("D-UT 2,5/10")> PC_EndCoverForUT2_5
        <Description("D-UT 2,5/4-TWIN")> PC_EndCoverForUT2_5MT
        <Description("D-UTTB 2,5/4")> PC_EndCoverForUTTB2_5
        <Description("D-URTK 6")> PC_EndCoverForURTK_6
        <Description("UBE/D")> PC_TerminalMarker
        'KLESMAN
        <Description("KD 4 Gray")> K_EndClamp
        <Description("NPP/AVK 2,5-10")> K_EndCoverForForAVK2_5_10
        <Description("NPP/AVK 4R-A")> K_EndCoverForAVK2_5_R_AVK_4_A
        <Description("GE Gray")> K_TerminalMarker
    End Enum

    Public Enum OrientationEnum
        <Description("Вертикальный")> Vertical
        <Description("Горизонтальный")> Horisontal
    End Enum

    Public Enum SideEnum
        <Description("Справа")> Rigth
        <Description("Слева")> Left
    End Enum
    
End Namespace