Attribute VB_Name = "main"
Sub ArchivementD3Bot()
'VERSION: 1.0
'DEVELOPER: Eduardo Inocencio
'Departament: Accounting
    
 Dim vresp As String
 
    vresp = MsgBox("Deseja confirmar essa opera��o?", vbQuestion + vbYesNo, "Mover e-mail e extrair anexos.")
    
    If vresp = vbYes Then
    
        Call MoveItems
        Call saveToD3BotFolder
        MsgBox "Processo realizado com sucesso!", vbInformation, "D3 Attachments"
        
    Else
        
        MsgBox "Opera��o cancelada!", vbCritical, "Cancelado"
    
    End If

End Sub
