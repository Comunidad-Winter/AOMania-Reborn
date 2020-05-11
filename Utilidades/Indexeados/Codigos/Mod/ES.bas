Attribute VB_Name = "ES"
Option Explicit

Public NumObjDatas As Integer

Sub LoadObjData()
   
   Dim i As Long
   
   Dim Leer As New clsIniManager
   
   Call Leer.Initialize(DatPath & "Obj.Dat")
   
   NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))
   
   ReDim Preserve ObjData(1 To NumObjDatas) As tObjData
   
   For i = 1 To NumObjDatas
       
       ObjData(i).GrhIndex = Val(Leer.GetValue("OBJ" & i, "GrhIndex"))
       ObjData(i).ObjType = Val(Leer.GetValue("OBJ" & i, "ObjType"))
       ObjData(i).NumObj = i
       
   Next i
     
End Sub
