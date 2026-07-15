Attribute VB_Name = "Módulo1"
Option Explicit

'========================
' Constantes
'========================
Const SEPARADOR_SENDTO As String = ","
Const FIRST_FIELD_COL As Long = 2
Const ANEXO_AUTO_MARK As String = "[ANEXO_AUTO]"
Const SEPARADOR_ANEXO As String = ";"

' Coluna nova na base para escolher conta de envio
Const FIELD_CONTA_ENVIO As String = "[conta_envio]"
Const POOL_DESENVOLVIMENTO As String = "pool-id.desenvolvimento_rede@daimlertruck.com"

'========================
' Util: substitui todas ocorrências de Fieldname no texto
'========================
Public Function SetTxtField(ByVal Fieldname As String, ByVal Content As String, ByVal str As String) As String
    Dim pos As Long
    Dim posatual As Long
    
    posatual = 1
    pos = InStr(posatual, str, Fieldname)
    
    If Fieldname = "" Then
        str = Content
    Else
        Do Until pos = 0
            str = Left$(str, pos - 1) & Content & Mid$(str, pos + Len(Fieldname), Len(str))
            posatual = pos + Len(Fieldname)
            pos = InStr(posatual, str, Fieldname)
        Loop
    End If
    
    SetTxtField = str
End Function

'========================
' Lê o conteúdo de um campo (cabeçalho) na linha RecordRow
'========================
Public Function GetContent(sheetmails As Worksheet, ByVal Fieldname As String, ByVal RecordRow As Long) As Variant
    Dim col As Long
    Dim FieldsRow As Long
    
    FieldsRow = sheetmails.Range("FieldsRow").Row
    
    col = FIRST_FIELD_COL
    Do Until sheetmails.Cells(FieldsRow, col).Value = Fieldname Or sheetmails.Cells(FieldsRow, col).Value = ""
        col = col + 1
    Loop
    
    If sheetmails.Cells(FieldsRow, col).Value <> "" Then
        GetContent = sheetmails.Cells(FieldsRow + RecordRow, col).Value
    Else
        GetContent = ""
    End If
End Function

'========================
' Retorna a coluna onde está um Fieldname no cabeçalho
'========================
Public Function GetFieldCol(sheetmails As Worksheet, ByVal Fieldname As String) As Long
    Dim col As Long
    Dim FieldsRow As Long
    
    FieldsRow = sheetmails.Range("FieldsRow").Row
    
    col = FIRST_FIELD_COL
    Do Until sheetmails.Cells(FieldsRow, col).Value = Fieldname Or sheetmails.Cells(FieldsRow, col).Value = ""
        col = col + 1
    Loop
    
    If sheetmails.Cells(FieldsRow, col).Value = Fieldname Then
        GetFieldCol = col
    Else
        GetFieldCol = 0
    End If
End Function

'========================
' Substitui todos os campos do cabeçalho dentro do Content
'========================
Public Function SetAllTxtFields(sheetmails As Worksheet, ByVal Content As String, ByVal RecordRow As Long) As String
    Dim FieldsRow As Long
    Dim FieldAtual As String
    Dim col As Long
    
    FieldsRow = sheetmails.Range("FieldsRow").Row
    
    col = FIRST_FIELD_COL
    FieldAtual = CStr(sheetmails.Cells(FieldsRow, col).Value)
    
    Do
        Content = SetTxtField(FieldAtual, CStr(GetContent(sheetmails, FieldAtual, RecordRow)), Content)
        col = col + 1
        FieldAtual = CStr(sheetmails.Cells(FieldsRow, col).Value)
    Loop Until FieldAtual = ""
    
    SetAllTxtFields = Content
End Function

'========================
' Escreve status na coluna A (mensagem por linha)
'========================
Public Sub SetMSG(sheetmails As Worksheet, ByVal msg As String, ByVal RecordRow As Long)
    Dim FieldsRow As Long
    FieldsRow = sheetmails.Range("FieldsRow").Row
    sheetmails.Cells(FieldsRow + RecordRow, 1).Value = msg
End Sub

'========================
' Resolve a conta SMTP para envio:
' - lê [conta_envio] da linha
' - vazio => usa Outlook padrão (retorna "")
' - "pool" (qualquer case) => usa POOL_DESENVOLVIMENTO
' - e-mail => usa direto
' Também retorna célula de origem para referência de erro.
'========================
Public Function ResolveSendAccountSmtp(sheetmails As Worksheet, ByVal RecordRow As Long, ByRef errCellAddress As String) As String
    Dim contaLinha As String
    Dim c As Long, FieldsRow As Long
    
    errCellAddress = ""
    
    contaLinha = Trim$(CStr(GetContent(sheetmails, FIELD_CONTA_ENVIO, RecordRow)))
    If contaLinha = "" Then
        ResolveSendAccountSmtp = ""   ' usa padrão do Outlook
        Exit Function
    End If
    
    ' célula para referência (linha/coluna)
    FieldsRow = sheetmails.Range("FieldsRow").Row
    c = GetFieldCol(sheetmails, FIELD_CONTA_ENVIO)
    If c > 0 Then errCellAddress = sheetmails.Cells(FieldsRow + RecordRow, c).Address(False, False)
    
    If StrComp(contaLinha, "pool", vbTextCompare) = 0 Then
        ResolveSendAccountSmtp = POOL_DESENVOLVIMENTO
    Else
        ResolveSendAccountSmtp = contaLinha
    End If
End Function

'========================
' Procura uma conta no Outlook por SMTP e DisplayName. Se não achar, falha.
'========================
Public Function TryGetOutlookAccount(ByVal olApp As Object, ByVal smtpOrName As String, ByRef accountOut As Object, ByRef errMsg As String) As Boolean
    Dim acc As Object
    Dim target As String
    Dim smtp As String
    
    target = LCase$(Trim$(smtpOrName))
    smtp = target
    
    ' 1) Tentativa exata por SmtpAddress
    For Each acc In olApp.Session.Accounts
        On Error Resume Next
        If LCase$(Trim$(acc.SmtpAddress)) = smtp Then
            Set accountOut = acc
            TryGetOutlookAccount = True
            Exit Function
        End If
        On Error GoTo 0
    Next
    
    ' 2) Tentativa exata por DisplayName
    For Each acc In olApp.Session.Accounts
        On Error Resume Next
        If LCase$(Trim$(acc.DisplayName)) = target Then
            Set accountOut = acc
            TryGetOutlookAccount = True
            Exit Function
        End If
        On Error GoTo 0
    Next
    
    ' 3) Tentativa "contém" (DisplayName ou SmtpAddress contém o texto)
    For Each acc In olApp.Session.Accounts
        On Error Resume Next
        If InStr(1, LCase$(acc.SmtpAddress), target, vbTextCompare) > 0 Or _
           InStr(1, LCase$(acc.DisplayName), target, vbTextCompare) > 0 Then
            Set accountOut = acc
            TryGetOutlookAccount = True
            Exit Function
        End If
        On Error GoTo 0
    Next
    
    errMsg = "Conta não encontrada/configurada no Outlook: " & smtpOrName
    TryGetOutlookAccount = False
End Function

'========================
' Envia e-mail via Outlook
'========================
Public Function EnviaEmails(ByVal SendTo As String, ByVal CopyTo As String, ByVal Subject As String, ByVal Text As String, ByVal AttachFile As String, _
                            Optional ByRef msgerr As String, Optional ByVal SendAccountSmtp As String = "") As Boolean
    
    Dim Outlook As Object
    Dim EmailOutlook As Object
    
    Dim i As Long
    Dim posini As Long
    Dim posfim As Long
    Dim qtdSendTo As Long
    Dim qtdCopyTo As Long
    Dim qtdAttach As Long
    Dim SendToList() As String
    Dim CopyToList() As String
    Dim AttachList() As String
    
    Dim acc As Object, accErr As String
    
    On Error GoTo ErrEnviaMail
    
    '-----------------------------
    ' Monta lista SendTo
    '-----------------------------
    qtdSendTo = 0
    posini = 1
    Do
        posfim = InStr(posini, SendTo, SEPARADOR_ANEXO)
        ReDim Preserve SendToList(qtdSendTo)
        SendToList(qtdSendTo) = Trim$(Mid$(SendTo, posini, IIf(posfim = 0, Len(SendTo) + 1, posfim) - posini))
        If SendToList(qtdSendTo) <> "" Then qtdSendTo = qtdSendTo + 1
        posini = posfim + 1
    Loop Until posfim = 0
    
    '-----------------------------
    ' Monta lista CopyTo
    '-----------------------------
    qtdCopyTo = 0
    posini = 1
    Do
        posfim = InStr(posini, CopyTo, SEPARADOR_SENDTO)
        ReDim Preserve CopyToList(qtdCopyTo)
        CopyToList(qtdCopyTo) = Trim$(Mid$(CopyTo, posini, IIf(posfim = 0, Len(CopyTo) + 1, posfim) - posini))
        If CopyToList(qtdCopyTo) <> "" Then qtdCopyTo = qtdCopyTo + 1
        posini = posfim + 1
    Loop Until posfim = 0
    
    '-----------------------------
    ' Monta lista Attachments
    '-----------------------------
    qtdAttach = 0
    posini = 1
    Do
        posfim = InStr(posini, AttachFile, ";")
        ReDim Preserve AttachList(qtdAttach)
        AttachList(qtdAttach) = Trim$(Mid$(AttachFile, posini, IIf(posfim = 0, Len(AttachFile) + 1, posfim) - posini))
        If AttachList(qtdAttach) <> "" Then qtdAttach = qtdAttach + 1
        posini = posfim + 1
    Loop Until posfim = 0
    
    '-----------------------------
    ' Cria e-mail
    '-----------------------------
    Set Outlook = CreateObject("Outlook.Application")
    Set EmailOutlook = Outlook.CreateItem(0)
    
    If qtdSendTo = 0 Then
        msgerr = "Sem destinatário (To) após substituição."
        EnviaEmails = False
        GoTo FimEnviaMail
    End If
    
    For i = 0 To qtdSendTo - 1
        EmailOutlook.Recipients.Add(SendToList(i)).Type = 1 'olTo
    Next
    
    For i = 0 To qtdCopyTo - 1
        EmailOutlook.Recipients.Add(CopyToList(i)).Type = 2 'olCC
    Next
    
    EmailOutlook.Subject = Subject
    
    If UCase$(Left$(Text, 6)) = "<HTML>" Then
        EmailOutlook.HTMLBody = Text
    Else
        EmailOutlook.Body = Text
    End If
    
    ' Anexos: se não existir, falha
    For i = 0 To qtdAttach - 1
        If AttachList(i) <> "" Then
            If Dir$(AttachList(i)) = "" Then
                msgerr = "Anexo não encontrado: " & AttachList(i)
                EnviaEmails = False
                GoTo FimEnviaMail
            End If
            EmailOutlook.Attachments.Add AttachList(i)
        End If
    Next
    
    ' Conta de envio: se informado, deve existir
    If Trim$(SendAccountSmtp) <> "" Then
        If Not TryGetOutlookAccount(Outlook, SendAccountSmtp, acc, accErr) Then
            msgerr = accErr
            EnviaEmails = False
            GoTo FimEnviaMail
        End If
        Set EmailOutlook.SendUsingAccount = acc
    End If
    
    EmailOutlook.Send
    
    msgerr = ""
    EnviaEmails = True
    GoTo FimEnviaMail
    
ErrEnviaMail:
    EnviaEmails = False
    msgerr = Err.Description
    
FimEnviaMail:
    On Error GoTo 0
    Set EmailOutlook = Nothing
    Set Outlook = Nothing
End Function

'========================
' ROTINA PRINCIPAL (renomeada) - o botão vai chamar esta
'========================
Public Sub EnviarEmails_Main()

    If MsgBox("Deseja enviar os emails agora ?", vbDefaultButton2 + vbYesNo + vbQuestion, "Envio automático de Emails") <> vbYes Then Exit Sub
    
    Dim sheetmails As Worksheet
    Dim regAtual As Long
    Dim keyField As String
    
    Dim SendTo As String
    Dim CopyTo As String
    Dim Subject As String
    Dim Body As String
    Dim attachment As String
    Dim msg As String
    Dim QtdSent As Long
    Dim QtdErr As Long
    
    Dim contaSmtp As String
    Dim contaErrCell As String
    
    ' >>> mover para o topo
    Dim resolvedAttachment As String
    Dim attachErr As String

    Set sheetmails = Sheets("emailsEnviar")
    keyField = CStr(sheetmails.Range("keyField").Value)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    regAtual = 1
    Do While CStr(GetContent(sheetmails, keyField, regAtual)) <> ""
        
        If UCase$(CStr(GetContent(sheetmails, "*enviar*", regAtual))) = "X" Then
            SetMSG sheetmails, ">>>", regAtual
            
            ' Resolve conta de envio por linha ([conta_envio])
            contaSmtp = ResolveSendAccountSmtp(sheetmails, regAtual, contaErrCell)
            
            SendTo = CStr(sheetmails.Range("SendToTest").Value)
            CopyTo = ""
            
            If SendTo = "" Then
                SendTo = SetAllTxtFields(sheetmails, CStr(sheetmails.Range("SendTo").Value), regAtual)
                CopyTo = SetAllTxtFields(sheetmails, CStr(sheetmails.Range("CopyTo").Value), regAtual)
            End If
            
            Subject = SetAllTxtFields(sheetmails, CStr(sheetmails.Range("Subject").Value), regAtual)
            Body = SetAllTxtFields(sheetmails, CStr(sheetmails.Range("Body").Value), regAtual)
            attachment = SetAllTxtFields(sheetmails, CStr(sheetmails.Range("Attachment").Value), regAtual)
            
            ' Resolve anexo automático (se Attachment tiver [ANEXO_AUTO])
            If Not ResolveAutoAttachment(sheetmails, regAtual, attachment, resolvedAttachment, attachErr) Then
                SetMSG sheetmails, "ERR - " & attachErr, regAtual
                QtdErr = QtdErr + 1
                GoTo ProximaLinha
            End If
            attachment = resolvedAttachment
            
            If EnviaEmails(SendTo, CopyTo, Subject, Body, attachment, msg, contaSmtp) Then
                SetMSG sheetmails, "OK", regAtual
                QtdSent = QtdSent + 1
            Else
                If contaErrCell <> "" Then
                    SetMSG sheetmails, "ERR - " & msg & " (ver " & contaErrCell & ")", regAtual
                Else
                    SetMSG sheetmails, "ERR - " & msg, regAtual
                End If
                QtdErr = QtdErr + 1
            End If
        Else
            SetMSG sheetmails, "", regAtual
        End If
        
ProximaLinha:
        regAtual = regAtual + 1
    Loop
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Tarefa Finalizada" & vbCrLf & "Enviados: " & QtdSent & vbCrLf & "Erros: " & QtdErr

End Sub

' =========================================================
' ANEXO AUTOMÁTICO POR COD1 + TÍTULO DO LOTE (C16)
' Aceita VCD 005-26 (prioridade) e VCD 005.26 (fallback)
' Se achar vários, pega o mais recente.
' =========================================================

Private Function Cod3(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then
        Cod3 = ""
    ElseIf IsNumeric(s) Then
        Cod3 = Right$("000" & CStr(CLng(s)), 3)
    Else
        Cod3 = s
    End If
End Function

Private Function EnsureTrailingSlash(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath <> "" Then
        If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    End If
    EnsureTrailingSlash = folderPath
End Function

Private Function GetTituloLoteAnexo(ByVal sheetmails As Worksheet) As String
    On Error Resume Next
    GetTituloLoteAnexo = Trim$(CStr(sheetmails.Range("TituloLoteAnexo").Value))
    On Error GoTo 0
End Function

' Retorna o arquivo mais recente que bate com o wildcard.
Private Function FindLatestFileByPattern(ByVal folderPath As String, ByVal pattern As String, ByRef foundFullPath As String, ByRef errMsg As String) As Boolean
    Dim f As String, full As String
    Dim bestFull As String
    Dim bestDate As Date, dt As Date
    Dim foundAny As Boolean

    foundFullPath = ""
    errMsg = ""
    folderPath = EnsureTrailingSlash(folderPath)

    f = Dir$(folderPath & pattern)
    foundAny = False

    Do While f <> ""
        full = folderPath & f
        On Error Resume Next
        dt = FileDateTime(full)
        On Error GoTo 0

        If Not foundAny Then
            foundAny = True
            bestDate = dt
            bestFull = full
        ElseIf dt > bestDate Then
            bestDate = dt
            bestFull = full
        End If

        f = Dir$()
    Loop

    If Not foundAny Then
        errMsg = "Anexo não encontrado (padrão): " & folderPath & pattern
        FindLatestFileByPattern = False
    Else
        foundFullPath = bestFull
        FindLatestFileByPattern = True
    End If
End Function

' Resolve o attachment quando estiver em modo automático: ...\[ANEXO_AUTO]
' Usa [cod1] e o título da leva (C16). Prioriza "-26" e faz fallback para ".26".
Private Function ResolveAutoAttachment(ByVal sheetmails As Worksheet, ByVal RecordRow As Long, ByVal attachmentTemplate As String, _
                                      ByRef resolvedAttachment As String, ByRef errMsg As String) As Boolean
    Dim folderPath As String
    Dim cod As String
    Dim titulo As String
    Dim patternDash As String, patternDot As String
    Dim found As String, errLocal As String
    Dim ok As Boolean

    resolvedAttachment = attachmentTemplate
    errMsg = ""

    If InStr(1, attachmentTemplate, ANEXO_AUTO_MARK, vbTextCompare) = 0 Then
        ResolveAutoAttachment = True
        Exit Function
    End If

    folderPath = Replace$(attachmentTemplate, ANEXO_AUTO_MARK, "", 1, -1, vbTextCompare)
    folderPath = Trim$(folderPath)

    cod = Cod3(GetContent(sheetmails, "[cod1]", RecordRow))
    If cod = "" Then
        errMsg = "Campo [cod1] vazio (necessário para anexo automático)."
        ResolveAutoAttachment = False
        Exit Function
    End If

    titulo = GetTituloLoteAnexo(sheetmails)

    ' Prioridade: hífen "-26"
    If titulo <> "" Then
        patternDash = "VCD " & cod & "-26* " & titulo & "*.pdf"
        patternDot = "VCD " & cod & ".26* " & titulo & "*.pdf"
    Else
        patternDash = "VCD " & cod & "-26*.pdf"
        patternDot = "VCD " & cod & ".26*.pdf"
    End If

    ok = FindLatestFileByPattern(folderPath, patternDash, found, errLocal)
    If Not ok Then
        ' fallback para ".26"
        ok = FindLatestFileByPattern(folderPath, patternDot, found, errLocal)
    End If

    If Not ok Then
        errMsg = errLocal
        ResolveAutoAttachment = False
        Exit Function
    End If

    resolvedAttachment = found
    ResolveAutoAttachment = True
End Function
