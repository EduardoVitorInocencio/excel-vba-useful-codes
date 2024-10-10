Attribute VB_Name = "functions"
Option Explicit

Function formatarCNPJ(Documento As String) As String
    ' Remove qualquer caractere que não seja número
    Dim DocumentoLimpo As String
    Dim i As Integer
    
    ' Limpar o documento, removendo espaços e caracteres não numéricos
    For i = 1 To Len(Documento)
        If IsNumeric(Mid(Documento, i, 1)) Then
            DocumentoLimpo = DocumentoLimpo & Mid(Documento, i, 1)
        End If
    Next i

    ' Verificar se é CPF (11 dígitos) ou CNPJ (14 dígitos)
    Select Case Len(DocumentoLimpo)
        Case 11 ' CPF
            ' Formatar o CPF no formato XXX.XXX.XXX-XX
            formatarCNPJ = Left(DocumentoLimpo, 3) & "." & Mid(DocumentoLimpo, 4, 3) & "." & Mid(DocumentoLimpo, 7, 3) & "-" & Right(DocumentoLimpo, 2)
        
        Case 14 ' CNPJ
            ' Formatar o CNPJ no formato XX.XXX.XXX/XXXX-XX
            formatarCNPJ = Left(DocumentoLimpo, 2) & "." & Mid(DocumentoLimpo, 3, 3) & "." & Mid(DocumentoLimpo, 6, 3) & "/" & Mid(DocumentoLimpo, 9, 4) & "-" & Right(DocumentoLimpo, 2)
        
        Case Else
            ' Retornar mensagem de erro caso o documento não tenha o tamanho correto
            formatarCNPJ = "Documento inválido"
    End Select
End Function

