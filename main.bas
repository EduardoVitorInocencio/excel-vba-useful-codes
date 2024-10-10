Attribute VB_Name = "main"
Sub ArchivementD3Bot()
'VERSION: 1.0
'DEVELOPER: Eduardo Inocencio
'Departament: Accounting
    
 Dim vresp As String
 
    vresp = MsgBox("Deseja confirmar essa operação?", vbQuestion + vbYesNo, "Mover e-mail e extrair anexos.")
    
    If vresp = vbYes Then
    
        Call MoveItems
        Call saveToD3BotFolder
        MsgBox "Processo realizado com sucesso!", vbInformation, "D3 Attachments"
        
    Else
        
        MsgBox "Operação cancelada!", vbCritical, "Cancelado"
    
    End If

End Sub
