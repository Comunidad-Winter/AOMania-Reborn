Attribute VB_Name = "Ranking"
Option Explicit

Public PtDesafio As Long
Public NickPtDesafio As String



Function DirINI()
     DirINI = App.Path & "\Dat\ini\"
End Function


Sub Load_Ranking()
     
     '[Mayor desafiador]
     PtDesafio = val(GetVar(DirINI & "Ranking.ini", "Ranking", "PtDesafio"))
     NickPtDesafio = GetVar(DirINI & "Ranking.ini", "Ranking", "NickPtDesafio")
     
     
     
     
     
     
     
End Sub


Sub Save_Ranking(UserIndex As Integer)
    
    With UserList(UserIndex)
    
       If Criminal(UserIndex) Then
          
          If .Stats.ELV > val(GetVar(DirINI & "Ranking.ini", "Ranking", "NivelMaxCriminal")) Then
              
          End If
          
       End If
       
       
       
    End With
    
    
End Sub
