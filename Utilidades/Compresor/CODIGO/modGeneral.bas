Attribute VB_Name = "modGeneral"
Option Explicit

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    
    FileExist = (LenB(Dir$(file, FileType)) <> 0)

End Function
