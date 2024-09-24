VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   OleObjectBlob   =   "FormCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelar_Click()
    Range("A1").Select
    Unload Me
End Sub


Private Sub Cadastrar_Click()

    '==
    Sheets("pesquisa").Activate
    
    'remove espaços em branco a direita e esquerda do form
    Me.TextName.Value = WorksheetFunction.Trim(Me.TextName)
    Me.TextCpf.Value = WorksheetFunction.Trim(Me.TextCpf)
    Me.TextSenha.Value = WorksheetFunction.Trim(Me.TextSenha)

    'Verificar se tem campos em branco
    If Me.TextName = "" Or Me.TextCpf = "" Or Me.TextSenha = "" Then
        MsgBox "Não são permitidos campos em branco!", vbInformation
        Exit Sub
    End If
    
    'Forçar somente números
'    If Not IsNumeric(Me.TextCpf.Value) Then
'        MsgBox "São permitidos somente números neste campo!*", vbInformation
'        Exit Sub
'    End If
    
    If Len(Me.TextCpf.Value) < 14 Or Len(Me.TextCpf.Value) > 14 Then
        MsgBox "Digite a quantidade correta de números"
        Exit Sub
    End If

    'Limpa qualquer texto nas caixas de pesquisa da planilha nome e cpf
    Dim tbx As OLEObject
    For Each tbx In ActiveSheet.OLEObjects
        If TypeName(tbx.Object) = "TextBox" Then tbx.Object.Text = ""
    Next

    Range("C6:E6").AutoFilter
    Range("C6:E6").AutoFilter

    'Range("A1").Select

    'Range("E1048576").End(xlUp).Row
    linha = Range("D7").End(xlDown).Row + 1

    lin = 7
    '==
    
    'Forçar não duplicar cpf ou renavam
    While lin < linha
        If Cells(lin, 4) = TextCpf.Value Then
            MsgBox ("Registro já existe!"), vbInformation
            Exit Sub
        End If
        
        lin = lin + 1
    Wend
    
    '==Realizar o registro dos dados
    Range("C6:E6").AutoFilter
    Range("C6:E6").AutoFilter
    Range("C1048576").Select
    Selection.End(xlUp).Select '.Offset(-0, 0)
    ActiveCell.EntireRow.Select
    ActiveCell.EntireRow.Copy
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
   ' ActiveCell.EntireRow.Select
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Selection.End(xlToLeft).Select
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = Me.TextName.Text
    'transforma campo em maiusculo
    maiusc = UCase(ActiveCell.Value)
    ActiveCell.Value = maiusc
    'remove espaços em branco a direita e esquerda
    ActiveCell.Value = WorksheetFunction.Trim(ActiveCell)
        
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Me.TextCpf.Text
    'converta para texto
    ActiveCell.NumberFormat = "@"
    'remove espaços em branco a direita e esquerda
    ActiveCell.Value = WorksheetFunction.Trim(ActiveCell)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Me.TextSenha.Text
    'converta para texto
    ActiveCell.NumberFormat = "@"
    'remove espaços em branco a direita e esquerda
    ActiveCell.Value = WorksheetFunction.Trim(ActiveCell)
    
    ActiveCell.EntireRow.Font.Bold = False
        
'        'ActiveCell.Offset(0, -1).Select
'    If Len(TextCpf.Text) <= 10 Then
'        ActiveCell.Offset(0, 1).Select
'        ActiveCell.Value = "PLACA"
'        ActiveCell.Interior.Color = 65535
'    Else
'        ActiveCell.Offset(0, 1).Select
'        ActiveCell.Value = "MOTORISTA"
'        ActiveCell.Interior.Color = 65535
'        'Selection.Font.Size = 8
'    End If
'
'    If ActiveCell.Value = "MOTORISTA" Then
'        Selection.AutoFilter Field:=3, Criteria1:="MOTORISTA"
'        '*****
'        ActiveCell.EntireRow.Cut
'        Range("C8").EntireRow.Select
'        'ActiveCell.Offset(2, 0).EntireRow.Select
'        Selection.Insert Shift:=xlDown
'
'        Range("E6").Select
'        Columns("E").Find(What:="MOTORISTA").Select
'        linhaInicio = ActiveCell.Offset(0, -2).Row
'
'    Else
'        Selection.AutoFilter Field:=3, Criteria1:="PLACA"
'        Range("E6").Select
'        Columns("E").Find(What:="PLACA").Select
'        linhaInicio = ActiveCell.Offset(0, -2).Row
'    End If
    
    linhaInicio = ActiveCell.Offset(0, -2).Row
    
    linhaFim = Sheets("pesquisa").Range("E1048576").End(xlUp).Row
           
    Range("C" & linhaInicio & ":E" & linhaFim).Select
          
    ActiveWorkbook.Worksheets("pesquisa").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("pesquisa").Sort.SortFields.Add Key:= _
        ActiveCell, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("pesquisa").Sort
        .SetRange Range("E" & linhaInicio & ":C" & linhaFim)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("C6:E6").AutoFilter
    Range("C6:E6").AutoFilter
    Range("A1").Select
    
    '*******
    ThisWorkbook.Save
    
    MsgBox ("O registro foi cadastrado com sucesso!"), vbInformation
    
    Unload Me
    
    'Me.TextName = ""
    'Me.TextCpf = ""
    'Me.TextName.SetFocus
    
    '**************************** Me.Text0.Text = Left(Me.Text0.Text, 20)

End Sub



Private Sub TextName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))


End Sub

Private Sub TextCpf_Change()

    If Me.TextCpf = "" Then
        MsgBox "Não são permitidos campos em branco!", vbInformation
        Exit Sub
    End If
    
    Dim temp As String
    temp = Replace(Replace(Replace(TextCpf.Text, ".", ""), "-", ""), " ", "")
    
    If Len(temp) > 0 Then
        If Len(temp) <= 3 Then
            TextCpf.Text = temp
        ElseIf Len(temp) <= 6 Then
            TextCpf.Text = Left(temp, 3) & "." & Mid(temp, 4)
        ElseIf Len(temp) <= 9 Then
            TextCpf.Text = Left(temp, 3) & "." & Mid(temp, 4, 3) & "." & Mid(temp, 7)
        ElseIf Len(temp) <= 11 Then
            TextCpf.Text = Left(temp, 3) & "." & Mid(temp, 4, 3) & "." & Mid(temp, 7, 3) & "-" & Mid(temp, 10)
        End If
    End If
    
    TextCpf.SelStart = Len(TextCpf.Text)
    
'    If Not IsNumeric(Me.TextCpf.Value) Then
'        MsgBox "São permitidos somente números neste campo!*", vbInformation
'        'remove ponto, vírgulas, traços e barras do cpf ou renavam do form
'        Me.TextCpf.Value = Replace(Me.TextCpf.Value, ".", "")
'        Me.TextCpf.Value = Replace(Me.TextCpf.Value, ",", "")
'        Me.TextCpf.Value = Replace(Me.TextCpf.Value, "-", "")
'        Me.TextCpf.Value = Replace(Me.TextCpf.Value, "/", "")
'        Exit Sub
'    End If
    
    'remove espaços em branco a direita e esquerda do form
    Me.TextCpf.Value = WorksheetFunction.Trim(Me.TextCpf)
    
End Sub


Private Sub TextName_Change()

    If Me.TextName = "" Then
        MsgBox "Não são permitidos campos em branco!", vbInformation
        Exit Sub
    End If

'remove espaços em branco a direita e esquerda do form
'Me.TextName.Value = WorksheetFunction.Trim(Me.TextName)

End Sub

Private Sub TextSenha_Change()
    If Me.TextSenha = "" Then
        MsgBox "Não são permitidos campos em branco!", vbInformation
        Exit Sub
    End If
    
'    Me.TextSenha.Value = Replace(Me.TextSenha.Value, ".", "")
'    Me.TextSenha.Value = Replace(Me.TextSenha.Value, ",", "")
'    Me.TextSenha.Value = Replace(Me.TextSenha.Value, "-", "")
'    Me.TextSenha.Value = Replace(Me.TextSenha.Value, "/", "")
    
    'remove espaços em branco a direita e esquerda do form
    Me.TextSenha.Value = WorksheetFunction.Trim(Me.TextSenha)
End Sub

'Private Sub TextCpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    Cr = KeyAscii
'    Módulo2.bloquerCaractere
'    On Error Resume Next
'    KeyAscii = Valor
'
'End Sub

Private Sub TextCpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(TextCpf.Text) >= 14 Then
        KeyAscii.Value = 0
    Else
        If Len(TextCpf.Text) = 3 Or Len(TextCpf.Text) = 7 Then
            TextCpf.Text = TextCpf.Text & "."
        End If
        If Len(TextCpf.Text) = 11 Then
            TextCpf.Text = TextCpf.Text & "-"
        End If
    End If
End Sub

