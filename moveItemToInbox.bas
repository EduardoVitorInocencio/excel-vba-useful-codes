Attribute VB_Name = "moveItemToInbox"
Global Const caixaEmail As String = "edinocencio@grenke.com.br" '"GRENKE Brasil, Fiscal"  '--------> Inserir a caixa de E-mail do Fiscal GRENKE


Sub MoveItems()
'VERSION: 1.0
'DEVELOPER: Eduardo Inocencio
'Departament: Accounting

 Dim myNameSpace            As Outlook.namespace
 Dim myInbox                As Outlook.folder
 Dim myDestFolder           As Outlook.folder
 Dim myItems                As Outlook.items
 Dim account                As Outlook.account
 Dim myItem                 As Object
 Dim store                  As Outlook.store
 Dim responsaveis(1 To 3)   As String
 Dim storeName              As String
 Dim i                      As Integer
 
 '----------------------------------------------------------- Remetente dos e-mails
 responsaveis(1) = "Guilherme Catto"
 responsaveis(2) = "William Franco"
 responsaveis(3) = "Azevedo, Guilherme"
 '-----------------------------------------------------------
 
 i = 1
  
 storeName = caixaEmail
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 
 While i <= 3
 
     For Each store In myNameSpace.Stores
        
        If store.DisplayName = storeName Then
            
            Set myInbox = store.GetDefaultFolder(olFolderInbox)
            Set myItems = myInbox.items
            Set myDestFolder = myInbox.Folders("ARQUIVAR ROBO")
            Set myItem = myItems.Find("[SenderName] = '" & responsaveis(i) & "'")
            
            
            While TypeName(myItem) <> "Nothing"
    
                    mailSubject = myItem.Subject
                    encSubject = Mid(mailSubject, 1, 5)
        
                    If encSubject = "ENC: " Then
                        subjectPrefix = Left(Mid(mailSubject, 6, Len(mailSubject) - 5), 4)
                    Else
                        subjectPrefix = Left(mailSubject, 4)
                    End If
        
                    If subjectPrefix = "P252" Or subjectPrefix = "P133" Or subjectPrefix = "P209" Then
                            myItem.Categories = "ARQUIVADO D3 ROBO"
                            myItem.Move myDestFolder
                    End If
    
            Set myItem = myItems.FindNext
            
            Wend
    
        End If
        
     Next
     
  i = i + 1
  
 Wend
 
End Sub
