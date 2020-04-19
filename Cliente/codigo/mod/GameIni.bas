Attribute VB_Name = "GameIni"

Option Explicit

Public Type tCabecera 'Cabecera de los con

    desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public Type tSetupMods

    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean

End Type

Public ClientSetup As tSetupMods
Public MiCabecera  As tCabecera

