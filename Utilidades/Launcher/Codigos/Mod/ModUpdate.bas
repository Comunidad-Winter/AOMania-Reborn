Attribute VB_Name = "ModUpdate"
Option Explicit

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer
Rem Programado por Shedark

Public Sub Analizar()

    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    'lEstado.Caption = "Obteniendo datos..."
    
  If Not Comprobar_Conexión(Url_Path & "VEREXE.TXT") Then
     MsgBox "¡Error! Comprueba tu conexión de internet."
     Exit Sub
  End If
  
    If InStr(frmMain.Inet1.OpenURL(Url_Path & "VEREXE.TXT"), "<title>404 Not Found</title>") Then
        MsgBox "¡Error! No se ha podido ver archivo binario."
        Exit Sub
    End If
    
    iX = frmMain.Inet1.OpenURL(Url_Path & "VEREXE.TXT") 'Host
    tX = LeerInt(FileUpdate)
    
    DifX = iX - tX
    
    If Not (DifX = 0) Then
      frmMain.ProgressBar1.Visible = True

       For i = 1 To DifX
            frmMain.Inet1.AccessType = icUseDefault
            dNum = i + tX
            
            #If BuscarLinks Then 'Buscamos el link en el host (1)
                Inet1.Url = Inet1.OpenURL("http://argentumania.es/cosas/parches/Parche" & dNum & ".zip") 'Host
            #Else                'Generamos Link por defecto (0)
                frmMain.Inet1.Url = "http://argentumania.es/cosas/parches/Parche" & dNum & ".zip" 'Host
            #End If
            
            Directory = App.Path & "\Libs\Configuracion\Parche" & dNum & ".zip"
            bDone = False
            dError = False
            
            'lURL.Caption = Inet1.URL
            'lName.Caption = "Parche" & dNum & ".zip"
            'lDirectorio.Caption = App.Path & "\"
                
            frmMain.Inet1.Execute , "GET"
            
            Do While bDone = False
            DoEvents
            Loop
            
            If dError Then Exit Sub
            
            UnZip Directory, App.Path & "\"
            Kill Directory
        Next i
    End If

    Call GuardarInt(FileUpdate, iX)
    SaveSetting "AoMania", "Updater", "Status", "1"
    

    frmMain.ProgressBar1.Value = 0

   frmMain.ProgressBar1.Visible = False
   
   frmMain.txtUpdate.Visible = True
   'TimerOn = 1
   'SetUpdate = "1"
   frmMain.txtUpdate.Left = 3480
   'Ejecutador.Enabled = True
   frmMain.txtUpdate.Caption = "Actualización OK"
  ' SetUpdateChange = "1"

End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal Data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, Data
    Close #F
End Sub

