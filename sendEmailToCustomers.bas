Attribute VB_Name = "sendEmailToCustomers"
Sub EnviarNotasFiscais()
   
    '------------------------------------------------------------------------------------------
    Dim outlookApp              As Outlook.Application
    Dim namespace               As Outlook.namespace
    Dim folder                  As Outlook.folder
    Dim mailItem                As Outlook.mailItem
    Dim newMail                 As Outlook.mailItem
    Dim attachment              As Outlook.attachment
    Dim notaFiscal              As Outlook.attachment
    
    Dim db                      As DAO.Database
    Dim rs                      As DAO.Recordset
    Dim nomeDb                  As String
    Dim staticPath              As String
    Dim sql                     As String
    
    Dim item                    As Object
    Dim recipient               As String
    Dim customBody              As String
    Dim i                       As Integer
    Dim notaEncontrada          As Boolean
    Dim tempPath                As String
    
    Dim tempFileName            As String
    Dim fso                     As Object
    Dim requestNumber           As String
    Dim partes()                As String
    Dim branch                  As String
    Dim fullResquestNumber      As String
    
    Dim lesseeName              As String
    Dim lesseeVATID             As String
    Dim lesseeEmail             As String
    Dim lesseeStreet            As String
    Dim lesseeArea              As String
    Dim lesseeState             As String
    
    '------------------------------------------------------------------------------------------
    staticPath = "X:\333\133\Common\GRENKE Subsidiary\TI\Robot Process Automation\002 - Projects Done\004 - Automations with VBA\Outlook Accounting\version-002\"
    nomeDb = "database\dataBaseAccountingOutlook.accdb"
    
    On Error Resume Next
    
    ' Iniciar a aplicação Outlook
    Set outlookApp = New Outlook.Application
    Set namespace = outlookApp.GetNamespace("MAPI")
    
    storeName = caixaEmail
    Set namespace = Application.GetNamespace("MAPI")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = staticPath & "temp"
    
    If Not fso.FolderExists(tempPath) Then
        fso.CreateFolder (tempPath)
    End If


    For Each store In namespace.Stores
       
       If store.DisplayName = storeName Then           '--------> Valida o campo para que seja igual ao nome da caixa do Fiscal da Grenke
       
          Set myFolder = store.GetDefaultFolder(olFolderInbox)
          Set folder = myFolder.Folders("NOTAS EMITIDAS")
        
        End If
    Next
    
    ' Aplicar filtro para selecionar apenas os e-mails com a categoria "ENVIAR CLIENTE"
    Dim filteredItems As Outlook.items
    Dim filter As String
    filter = "[Categories] = 'ENVIAR CLIENTE/ARQUIVADO D3 ROBO'"
    Set filteredItems = folder.items.Restrict(filter)
    
    ' Defina as categorias que você deseja verificar
    Dim categoria1 As String
    Dim categoria2 As String
    
    Dim progressForm As frm_ProgressBar
    Set progressForm = New frm_ProgressBar
    progressForm.Show vbModeless ' Exibir o UserForm de forma não modal
    
    categoria1 = "ENVIAR CLIENTE"
    categoria2 = "PENDENTE"
    'UserForm1.Show
    
    ' Iterar pelos itens da pasta "notas emitidas"
    For i = 1 To filteredItems.Count
        
        ' Calcular o progresso
        progress = (i / filteredItems.Count) * 100
        If progress > 100 Then progress = 100 ' Limitar a 100%
        
        ' Atualizar o UserForm com o progresso
        progressForm.Update_Progress (progress)
      
        ' Atualizar a interface do usuário
        DoEvents
        
        Set item = folder.items(i)
        Dim AttchQtd As Integer
        Dim getAttchQtd As Integer
        
        If TypeName(item) = "MailItem" Then
            Set mailItem = item
            'If InStr(1, mailItem.Categories, categoria1) > 0 And InStr(1, mailItem.Categories, categoria2) Then
            
                'Inicializar a variável de controle
                notaEncontrada = False
                
                'Procurar pelo anexo de nota fiscal, buscando pela extensão do documento.
                For Each attachment In mailItem.Attachments
                    
                   
                    Dim filesExtension As String
                    Dim attchName As String
                    i = 1
                    
                   AttchQtd = mailItem.Attachments.Count
                   getAttchQtd = 1
                   
                   While getAttchQtd <= AttchQtd
                    
                    If InStr(attachment.FileName, ".pdf") > 0 Then
                        Set notaFiscal = attachment
                        
                        filesExtension = Right(notaFiscal.FileName, 4)
                        partes = Split(notaFiscal.FileName, "_")
                        requestNumber = partes(2)
                        branch = partes(1)
                        
                        If Not filesExtension = ".pdf" Then
                            filesExtension = ".INVÀLIDO"
                        End If
                        
                        attchName = branch & "-" & requestNumber & "_" & CStr(i) & filesExtension
                        tempFileName = tempPath & "\" & attchName
                        notaFiscal.SaveAsFile tempFileName
                        notaEncontrada = True
                        
                        partes = Split(notaFiscal.FileName, "_")
                        
                        fullResquestNumber = branch & "-" & requestNumber
                        
                        If Right(fullResquestNumber, 4) = ".pdf" Then
                            
                           fullResquestNumber = Mid(fullResquestNumber, 1, Len(fullResquestNumber) - 4)
                        
                        End If
                        
                        
                        'Exit For
                    End If
                        
                        i = i + 1
                        getAttchQtd = getAttchQtd + 1
                        
                    Wend
                '--------------------------------------------------------------------------------------------------------------------------------------------
                lesseeName = "NOME DO CLIENTE - PREENCHER"
                lesseeVATID = "CNPJ - PREENCHER"
                lesseeEmail = "EMAIL - PREENCHER"
                lesseeStreet = "RUA - PREENCHER"
                lesseeArea = "BAIRRO - PREENCHER"
                lesseeState = "ESTADO - PREENCHER"
                '--------------------------------------------------------------------------------------------------------------------------------------------
                
                sql = "SELECT lesseeName, lesseeVATID, lesseeEmail, lesseeStreet, lesseeArea, lesseeState FROM lesseeInfo WHERE branchRequest = '" & fullResquestNumber & "'"
                
                Set db = OpenDatabase(staticPath & nomeDb, True, False)
                Set rs = db.OpenRecordset(sql)
                
                rs.MoveFirst
                
                lesseeName = rs![lesseeName]
                lesseeVATID = rs![lesseeVATID]
                lesseeEmail = rs![lesseeEmail]
                lesseeStreet = rs![lesseeStreet]
                lesseeArea = rs![lesseeArea]
                lesseeState = rs![lesseeState]
                lesseeVATID = formatarCNPJ(lesseeVATID)
                
                'Se a nota fiscal for encontrada, eniar o e-mail ao cliente
                If notaEncontrada Then
                    ' Criar um novo e-mail
                    Set newMail = outlookApp.CreateItem(olMailItem)
                    
                    ' Definir destinatário, copor do e-mail e anexar a nota fiscal
                    recipient = mailItem.To
                    customBody = "<html><body style='font-family:Arial; font-size:14pt;'>"
                    
                    customBody = customBody & "<h3>" & lesseeName & "</h3>"
                    customBody = customBody & "<p>CNPJ: " & lesseeVATID & "<br>"
                    customBody = customBody & "Endereço: " & lesseeStreet & "," & lesseeArea & "-" & lesseeState & "</p>"
                    
                    customBody = customBody & "<br><h2>ENVIO DE NFe EMITIDA</h2>"
                    customBody = customBody & "<p>Prezado cliente,</p>"
                    customBody = customBody & "<p>A Grenke Brasil emite todas as suas Notas Fiscais em meio eletrônico, de acordo com a legislação vigente para a NF-e nacional. Esta mensagem refere-se ao envio obrigatório da NF-e emitida para o pedido " & fullResquestNumber & ".</p>"
                    
                    customBody = customBody & "<br><table border='1' style='border-collapse: collapse;'>"
                    customBody = customBody & "<tr style='background-color: #f0f0f0;'><td style='width:180px; font-family:Arial; font-size:12pt; font-weight: bold; border: 1px solid #d3d3d3;'>Razão Social</td><td style='width:580px; font-family:Arial; font-size:12pt; border: 1px solid #d3d3d3;'>GC LOCACAO DE EQUIPAMENTOS LTDA</td></tr>"
                    customBody = customBody & "<tr style='background-color: #f0f0f0;'><td style='width:180px; font-family:Arial; font-size:12pt; font-weight: bold; border: 1px solid #d3d3d3;'>CNPJ / CPF: </td><td style='width:580px; font-family:Arial; font-size:12pt; border: 1px solid #d3d3d3;'>14.262.033/0001-14</td></tr>"
                    customBody = customBody & "<tr style='background-color: #f0f0f0;'><td style='width:180px; font-family:Arial; font-size:12pt; font-weight: bold; border: 1px solid #d3d3d3;'>E-mail</td><td style='width:580px; font-family:Arial; font-size:12pt; border: 1px solid #d3d3d3;'>fiscal@grenke.com.br</td></tr>"
                    customBody = customBody & "</table>"
                    customBody = customBody & "</body></html>"
                    
                    
                    With newMail
                    
                        .To = lesseeEmail
                        .Subject = "Nota Fiscal Eletrônica - Pedido: " & fullResquestNumber
                        
                        .HTMLBody = customBody
                        
                        attchName = Dir(tempPath & "\" & "*.*")
                        i = 1
                        
                        Do While attchName <> ""
                            
                            .Attachments.Add tempFileName, olByValue, 1, "Nota Fiscal Eletrônica - Pedido: " & attchName
                            attchName = Dir
                        Loop
                        
                        .CC = "fiscal@grenke.com.br"
                        mailItem.Categories = "ENVIADO CLIENTE/ARQUIVADO D3 ROBO"
                        .Save
                        
                    End With
                    
                End If
                   
                Next attachment
                
                
                mailItem.Save
                
                rs.Close
                db.Close
            
            'End If
        End If
        
    Next i
    
    ' Fechar o UserForm
    progressForm.CloseForm
    Set progressForm = Nothing
    
    If fso.FolderExists(tempPath) Then
        fso.DeleteFolder tempPath, True
    End If
    
    ' Limpeza
    'Unload UserForm1
    Set mailItem = Nothing
    Set folder = Nothing
    Set namespace = Nothing
    
    Set outlookApp = Nothing
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "Notas enviadas ao clientes, verifique sua caixa de envios.", vbInformation, "Enviar Notas"
    
    
End Sub


