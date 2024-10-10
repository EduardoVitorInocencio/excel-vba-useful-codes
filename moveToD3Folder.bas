Attribute VB_Name = "moveToD3Folder"
Sub saveToD3BotFolder()
'VERSION: 1.0
'DEVELOPER: Eduardo Inocencio
'Departament: Accounting
 
 Dim myNameSpace        As Outlook.namespace
 Dim myFolder           As Outlook.folder
 Dim otherFolder        As Outlook.folder
 Dim myDestFolder       As Outlook.folder
 Dim mailAttachment     As Outlook.attachment
 Dim item               As Outlook.mailItem
 Dim store              As Outlook.store
  
 Dim storeName          As String
 Dim mailSubject        As String
 Dim robotsFolder       As String
 Dim encSubject         As String
 Dim strFolderpath      As String
 Dim attachBegin        As String
 Dim emailCategory      As String
 
 Dim responsaveis(1 To 3)   As String
 Dim i                      As Integer
 
 '----------------------------------------------------------- Remetente dos e-mails
 responsaveis(1) = "Guilherme Catto"
 responsaveis(2) = "William Franco"
 responsaveis(3) = "Azevedo, Guilherme"
 '-----------------------------------------------------------
 
 i = 1
 
 '-----------------------------------------------------------
 
 'robotsFolder = "X:\333\133\Common\ROBO" ' # Pasta correta
 'robotsFolder = "\\badfile\users\edinocencio\My Documents\pastaRoboD3"
 strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
 strFolderpath = strFolderpath & "\Attachments\"
 
 
 storeName = caixaEmail
 Set myNameSpace = Application.GetNamespace("MAPI")
 
 
For Each store In myNameSpace.Stores
   
   If store.DisplayName = storeName Then           '--------> Valida o campo para que seja igual ao nome da caixa do Fiscal da Grenke
   
      Set myFolder = store.GetDefaultFolder(olFolderInbox)
      Set otherFolder = myFolder.Folders("ARQUIVAR ROBO")
      Set myDestFolder = myFolder.Folders("NOTAS EMITIDAS")
      
     
            For Each item In otherFolder.items
            
                If otherFolder.items.Count = 0 Then
                
                    Exit For
                    
                End If
                
                   
                   mailSubject = item.Subject
                   encSubject = Mid(mailSubject, 1, 5)
                   
                   If encSubject = "ENC: " Then
                   
                        subjectPrefix = Left(Mid(mailSubject, 6, Len(mailSubject) - 5), 4)
                        
                   Else
                                 
                        subjectPrefix = Left(mailSubject, 4)
                        
                   End If
                   
                   If subjectPrefix = "P252" Or subjectPrefix = "P133" Or subjectPrefix = "P209" Then
                   
                           For Each mailAttachment In item.Attachments
                           
                               attachBegin = Left(mailAttachment.FileName, 3)
                                                  
                               If (attachBegin = "124" Or attachBegin = "106") Then
                                   
                                   item.Categories = "ARQUIVADO D3 ROBO"
                                   mailAttachment.SaveAsFile robotsFolder & "\" & mailAttachment
                                                                
                               End If
                           
                           Next
                           
                           item.Move myDestFolder
                       
                   End If
            Next
      
      
      
   End If

Next
    
End Sub
