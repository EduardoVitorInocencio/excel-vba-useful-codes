VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProgressBar 
   Caption         =   "Lendo e-mails"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   OleObjectBlob   =   "frm_ProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Update_Progress(progress As Single)
      ' Atualizar o UserForm com o progresso
        Me.txtProgress.Caption = "Progresso: " & Format(progress, "0.00") & "%"
        ' Preencher a barra de progresso
        Me.UpdateProgress.Width = 200 * (progress / 100) ' Ajuste o valor para a largura desejada
End Sub

Public Sub CloseForm()
    Unload Me
End Sub

