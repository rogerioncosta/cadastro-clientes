VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAlterarSenha 
   Caption         =   "Alterar Senha"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   OleObjectBlob   =   "FormAlterarSenha.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAlterarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AlterarSenha_Click()

    Sheets("pesquisa").Activate

    Dim senhaNova As String
    Dim senhaAtual As String
    senhaAtual = Sheets("Config").Range("A1").Value
    senhaNova = Me.TextNovaSenha.Value
    
    If Me.TextSenhaAtual = senhaAtual Then
        senhaAtual = senhaNova
        
        Sheets("Config").Range("A1").Value = senhaAtual
        
        '*******
        ThisWorkbook.Save
    
        MsgBox ("A senha para salvar o arquivo foi alterada!"), vbInformation
    
        Unload Me
    Else
        MsgBox "A senha atual está incorreta."
        Exit Sub
    End If
    
End Sub

Private Sub cancelar_Click()
    Range("A1").Select
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Me.TextSenhaAtual.SetFocus
End Sub
