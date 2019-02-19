'-----------------------------------------------------------------------------------------------------------------------------------------------
'- Aplicação do código em um evento do formulário (antes de atualizar - before update) para buscar valores inseridos em controles do formulário.
'- O código utiliza a função DlookupRecordset, que deve ser guardada dentro de um módulo do access.
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub Form_BeforeUpdate(Cancel As Integer)

Dim vAux    As Dim String

If IsNull (Me.Id) Then

'- Verificar se existe duplicidade na conta Db.
vAux = ""
vAux = DlookupRecordset("SELECT Id FROM [Parametrização de Integração Contabil] WHERE ObraCC = " & Me.ObraCC & " And DptoCC = " & Me.DptoCC & " And Db = '" & Me.Db & "'")
    If vAux <> "" Then  
        MsgBox "Já existe parametrização cadastrada!",vbInformation,"Parametrização Integração Folha Contabilidade"
        Cancel = True
        Me.Undo
        Exit Sub 
    End If

'- Verificar se existe duplicidade na conta Cr.
vAux = ""
vAux = DlookupRecordset("SELECT Id FROM [Parametrização de Integração Contabil] WHERE ObraCC = " & Me.ObraCC & " And DptoCC = " & Me.DptoCC & " And Cr = '" & Me.Cr & "'")
    If vAux <> "" Then
        MsgBox "Já existe parametrização cadastrada!",vbInformation,"Parametrização Integração Folha Contabilidade"
        Cancel = True
        Me.Undo
        Exit Sub
    End If

'- Verificar se existe duplicidade na conta Db + Cr.
vAux = ""
vAux = DlookupRecordset("SELECT Id FROM [Parametrização de Integração Contabil] WHERE ObraCC = " & Me.ObraCC & " And DptoCC = " & Me.DptoCC & " And Db = '" & Me.Db & "' And Cr = '" & Me.Cr & "'")
    If vAux <> "" Then
        MsgBox "Já existe parametrização cadastrada!",vbInformation,"Parametrização Integração Folha Contabilidade"
        Cancel = True
        Me.Undo
        Exit Sub
    End If
Else
    vAux = ""
    vAux = DlookupRecordset("SELECT Id FROM [Parametrização de Integração Contabil] WHERE Id <> " & Me.Id & " And ObraCC = " & Me.ObraCC & " And DptoCC = " & Me.DptoCC & " And Db = '" & Me.Db & "'")
        If vAux <> "" Then
            MsgBox "Já existe parametrização cadastrada!",vbInformation,"Parametrização Integração Folha Contabilidade"
            Cancel = True
            Me.Undo
            Exit Sub
        End If
    
    vAux = ""
    vAux = DlookupRecordset("SELECT Id FROM [Parametrização de Integração Contabil] WHERE Id <> " & Me.Id & " And ObraCC = " & Me.ObraCC & " And DptoCC = " & Me.DptoCC & " And Cr = '" & Me.Cr & "'")
        If vAux <> "" Then
            MsgBox "Já existe parametrização cadastrada!",vbInformation,"Parametrização Integração Folha Contabilidade"
            Cancel = True
            Me.Undo
            Exit Sub
        End If

    vAux = ""
    vAux = DlookupRecordset("SELECT Id FROM [Parametrização de Integração Contabil] WHERE Id <> " & Me.Id & " And ObraCC = " & Me.ObraCC & " And DptoCC = " & Me.DptoCC & " And Db = '" & Me.Db & "' And Cr = '" & Me.Cr & "'")
        If vAux <> "" Then
            MsgBox "Já existe parametrização cadastrada!", vbInformation,"Parametrização Integração Folha Contabilidade"
            Cancel = True
            Me.Undo
            Exit Sub
        End If
    End If

End Sub

'-----------------------------------------------------------------------------------------------------------------
'- Gravar a função abaixo dentro de um módulo. A declaração da função publica permite a chamada em todo o código.
'- Função DlookupRecordset para buscar valores dentro de uma table ou query, substituir apenas as strings.
'-----------------------------------------------------------------------------------------------------------------

Public Function DlookupRecordset (sSQL As String) As Variant

Dim SQL     As String
Dim Rs      As Recordset

Fnc_DlookupRecordset = ""
SQL = sSQL

Set Rs = CurrentDb.OpenRecordset(sSQL)
    If Rs.EOF = False Then
        Fnc_DlookupRecordset = Rs(0)
    End If

    Rs.Close
Set Rs = Nothing

End Function
