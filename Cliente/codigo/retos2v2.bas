Attribute VB_Name = "retos2v2"
Sub VerificarRetos(ByVal UserIndex As Integer)
On Error GoTo chao
If UserList(UserIndex).Reto.Retando_2 Then
    UserList(OPCDuelos.J1).Reto.Received_Request = False
    UserList(OPCDuelos.J1).Reto.Retando_2 = False
    UserList(OPCDuelos.J1).Reto.Send_Request = False
   
    UserList(OPCDuelos.J2).Reto.Received_Request = False
    UserList(OPCDuelos.J2).Reto.Retando_2 = False
    UserList(OPCDuelos.J2).Reto.Send_Request = False
   
    UserList(OPCDuelos.J3).Reto.Received_Request = False
    UserList(OPCDuelos.J3).Reto.Retando_2 = False
    UserList(OPCDuelos.J3).Reto.Send_Request = False
   
    UserList(OPCDuelos.J4).Reto.Received_Request = False
    UserList(OPCDuelos.J4).Reto.Retando_2 = False
    UserList(OPCDuelos.J4).Reto.Send_Request = False
   
    Call WarpUserChar(OPCDuelos.J1, 1, 50, 50, True)
    Call WarpUserChar(OPCDuelos.J2, 1, 51, 51, True)
    Call WarpUserChar(OPCDuelos.J3, 1, 52, 52, True)
    Call WarpUserChar(OPCDuelos.J4, 1, 53, 53, True)
 
    Call SendData(ToAll, 0, 0, "||2vs2: El reto se cancela porque " & UserList(UserIndex).name & " desconectó." & FONTTYPE_TALK)
 
    frmMain.retos.Enabled = False '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
    OPCDuelos.OCUP = False
    OPCDuelos.Tiempo = 0
    OPCDuelos.J1 = 0
    OPCDuelos.J2 = 0
    OPCDuelos.J3 = 0
    OPCDuelos.J4 = 0
End If
chao:
 End Sub
