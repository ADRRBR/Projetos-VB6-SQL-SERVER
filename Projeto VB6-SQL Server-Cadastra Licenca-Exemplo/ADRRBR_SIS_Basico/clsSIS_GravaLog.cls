VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_GravaLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sTituloAplicacao          As String
Private sCodigoAplicacao          As String
Private sNomeAplicacao            As String
Private lCodigoUsuarioAplicacao   As Long
Private sNomeUsuarioAplicacao     As String
Private lUltimo_PK_Log_Gravado_BD As Variant
Private cConexao                  As Object

Private iStatus                   As ADRRBR_SIS_Basico.eStatus
Private sMensagem                 As String

Dim rsLog                         As ADODB.Recordset

Dim oLog                          As Object
Dim oParamLog                     As Object

Private Const sClasse = "SIS_GRAVALOG"

Private Sub Class_Initialize()
    LimpaStatus
    lUltimo_PK_Log_Gravado_BD = "NULL"
End Sub

Private Sub Class_Terminate()
    Set cConexao = Nothing
    Set rsLog = Nothing
    Set oLog = Nothing
    Set oParamLog = Nothing
End Sub

Public Property Let TituloAplicacao(ByVal vNewValue As String)
    sTituloAplicacao = vNewValue
End Property
Public Property Get TituloAplicacao() As String
    TituloAplicacao = sTituloAplicacao
End Property

Public Property Let CodigoAplicacao(ByVal vNewValue As String)
    sCodigoAplicacao = vNewValue
End Property
Public Property Get CodigoAplicacao() As String
    CodigoAplicacao = sCodigoAplicacao
End Property

Public Property Let NomeAplicacao(ByVal vNewValue As String)
    sNomeAplicacao = vNewValue
End Property
Public Property Get NomeAplicacao() As String
    NomeAplicacao = sNomeAplicacao
End Property

Public Property Let CodigoUsuarioAplicacao(ByVal vNewValue As Long)
    lCodigoUsuarioAplicacao = vNewValue
End Property
Public Property Get CodigoUsuarioAplicacao() As Long
    CodigoUsuarioAplicacao = lCodigoUsuarioAplicacao
End Property

Public Property Let NomeUsuarioAplicacao(ByVal vNewValue As String)
    sNomeUsuarioAplicacao = vNewValue
End Property
Public Property Get NomeUsuarioAplicacao() As String
    NomeUsuarioAplicacao = sNomeUsuarioAplicacao
End Property

Public Property Let Ultimo_PK_Log_Gravado_BD(ByVal vNewValue As Variant)
    lUltimo_PK_Log_Gravado_BD = vNewValue
End Property
Public Property Get Ultimo_PK_Log_Gravado_BD() As Variant
    Ultimo_PK_Log_Gravado_BD = lUltimo_PK_Log_Gravado_BD
End Property

Public Property Set Conexao(ByVal vNewValue As Object)
    Set cConexao = vNewValue
End Property
Public Property Get Conexao() As Object
    Set Conexao = cConexao
End Property

Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Sub Grava(pDescricaoLog As String, pIdentificacao As String, pNomeObjeto As String, pNomeRotinaFonte As String)
    Dim sCaminhoArquivoLog As String
    Dim sColunaAux         As String
    
    On Error GoTo ErroRotina

    LimpaStatus
    
    If Not VerificaParametrosLog Then Exit Sub
    If Not VerificaPropriedadesGerais Then Exit Sub
    
    If pDescricaoLog = Empty Then
        iStatus = Erro
        sMensagem = "Informe obrigatoriamente a Descrição do Log. (Ex.: ARQUIVO INVÁLIDO / PROCESSAMENTO EFETUADO COM SUCESSO ) !"
        Exit Sub
    End If
    If pIdentificacao = Empty Then
        iStatus = Erro
        sMensagem = "Informe obrigatoriamente a Identificação do Log. (Ex.: ERRO / PROCESSAMENTO) !"
        Exit Sub
    End If
    If pNomeObjeto = Empty Then
        iStatus = Erro
        sMensagem = "Informe obrigatoriamente o nome do Objeto corrente. (Ex.: clsPED_Clientes / clsSIS_Basico) !"
        Exit Sub
    End If
    If pNomeRotinaFonte = Empty Then
        iStatus = Erro
        sMensagem = "Informe obrigatoriamente o nome da Rotina Fonte Corrente (Sub/Função). (Ex.: Consultar_BD / Conecta) !"
        Exit Sub
    End If
    
    If oParamLog.GeraLogLocal = 1 Then
        Set rsLog = Nothing
        Set rsLog = New ADODB.Recordset
       
        If oParamLog.VerificaAplicacaoInternet Then
            sCaminhoArquivoLog = oParamLog.CaminhoFisicoInternet
        Else
            sCaminhoArquivoLog = oParamLog.CaminhoLogLocal
        End If
        
        sCaminhoArquivoLog = sCaminhoArquivoLog & "\ADRRBR_Log.Log"
        
        If Dir(sCaminhoArquivoLog) <> Empty Then
            rsLog.Open sCaminhoArquivoLog
        Else
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Identificação", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "1000"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Objeto", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "1000"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Rotina Fonte", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "1000"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Descrição", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "9000"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Data Geração", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "100"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Código Aplicação", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "100"
            
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Aplicação", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "100"
            
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Computador", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "1000"
            
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Usuário Local", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "1000"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Código Usuário Aplicação", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "100"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Usuário Aplicação", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "1000"
        
            sColunaAux = oBasico.Geral.EncriptarDecriptar("Fonte Conexão BD", True)
            rsLog.Fields.Append sColunaAux, adVarChar, "4000"
        
            rsLog.Open
        End If
    
        rsLog.AddNew
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Identificação", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(pIdentificacao, True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Objeto", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(pNomeObjeto, True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Rotina Fonte", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(pNomeRotinaFonte, True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Descrição", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(pDescricaoLog, True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Data Geração", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(Format(Now, "dd/mm/yyyy hh:mm:ss"), True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Código Aplicação", True)
        If Trim(sCodigoAplicacao) <> Empty Then
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(sCodigoAplicacao, True)
        Else
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar("NÃO INFORMADO", True)
        End If
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Aplicação", True)
        If Trim(sNomeAplicacao) <> Empty Then
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(sNomeAplicacao, True)
        Else
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar("NÃO INFORMADO", True)
        End If
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Computador", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(oBasico.Geral.NomeComputador, True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Usuário Local", True)
        rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(oBasico.Geral.UsuarioLocal, True)
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Código Usuário Aplicação", True)
        If lCodigoUsuarioAplicacao > 0 Then
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(CStr(lCodigoUsuarioAplicacao), True)
        Else
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar("NÃO INFORMADO", True)
        End If
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Usuário Aplicação", True)
        If Trim(sNomeUsuarioAplicacao) <> Empty Then
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(sNomeUsuarioAplicacao, True)
        Else
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar("NÃO INFORMADO", True)
        End If
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("Fonte Conexão BD", True)
        If cConexao Is Nothing Then
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar("CONEXÃO NÃO INFORMADA", True)
        Else
            rsLog.Fields(sColunaAux).Value = oBasico.Geral.EncriptarDecriptar(cConexao.FonteConexao, True)
        End If
        
        rsLog.Update
        
        If Dir(sCaminhoArquivoLog) <> Empty Then Kill sCaminhoArquivoLog
        rsLog.Save sCaminhoArquivoLog
        rsLog.Close
        Set rsLog = Nothing
    End If
    
    If oParamLog.GeraLogBancoDados = 1 Then
        If cConexao Is Nothing Then
            iStatus = Cancelado
            sMensagem = "Não foi possível gravar o LOG no banco de dados, pois não foi informada a CONEXÃO!"
            Exit Sub
        End If

        Set oLog = Nothing
        Set oLog = CreateObject("ADRRBR_SIS_Log.clsSIS_Log")
                
        oLog.Acao = Incluir
        Set oLog.Conexao = cConexao
        oLog.FK_Log = lUltimo_PK_Log_Gravado_BD
        oLog.VCH_Identificacao = pIdentificacao
        oLog.VCH_Objeto = pNomeObjeto
        oLog.VCH_Rotina_Fonte = pNomeRotinaFonte
        oLog.VCH_Descricao = pDescricaoLog
        oLog.DTT_Geracao = Format(Now, "dd/mm/yyyy hh:mm:ss")
        oLog.VCH_Codigo_Aplic = sCodigoAplicacao
        oLog.VCH_Nome_Aplic = sNomeAplicacao
        oLog.VCH_Nome_Computador = oBasico.Geral.NomeComputador
        oLog.VCH_Nome_Usuario_Local = oBasico.Geral.UsuarioLocal
        oLog.INT_Codigo_Usuario_Aplic = lCodigoUsuarioAplicacao
        oLog.VCH_Nome_Usuario_Aplic = sNomeUsuarioAplicacao
        oLog.VCH_Conexao = cConexao.FonteConexao
        oLog.Atualizar_BD
        
        lUltimo_PK_Log_Gravado_BD = oLog.PK_Log
        
        iStatus = oLog.Status
        sMensagem = oLog.Mensagem
        
        Set oLog = Nothing
    End If
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    Set rsLog = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Grava > "
End Sub

Public Sub ExcluiVencidos()
    On Error GoTo ErroRotina

    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub

    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "ExcluiVencidos > "
End Sub

Public Sub ConsultaLogLocal()
    On Error GoTo ErroRotina

    LimpaStatus
    
    frmConsultaLOGLocal.TituloAplicacao = sTituloAplicacao
    frmConsultaLOGLocal.Show vbModal
    
    iStatus = frmConsultaLOGLocal.Status
    sMensagem = frmConsultaLOGLocal.Mensagem
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "ConsultaLogLocal > "
End Sub

Private Function VerificaPropriedadesGerais() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesGerais = True
    
    If Not cConexao Is Nothing Then
        If Not cConexao.Conectado Then
            iStatus = Erro
            sMensagem = "Se informar o objeto de conexão, o mesmo deve estar CONECTADO!"
            VerificaPropriedadesGerais = False
            Exit Function
        End If
    End If
    
    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaPropriedadesGerais = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "VerificaPropriedadesGerais > "
End Function

Private Function VerificaParametrosLog() As Boolean
    On Error GoTo ErroRotina
    
    VerificaParametrosLog = True
    
    Set oParamLog = Nothing
    Set oParamLog = CreateObject("ADRRBR_SIS_Param_Log.clsSIS_Param_Log")
    
    oParamLog.RecuperaParametrosLog
    
    If oParamLog.Status <> Sucesso Then
        iStatus = oParamLog.Status
        sMensagem = oParamLog.Mensagem
        VerificaParametrosLog = False
        Exit Function
    End If
    
    If oParamLog.GeraLogLocal = 0 And oParamLog.GeraLogBancoDados = 0 Then
        VerificaParametrosLog = False
        Exit Function
    End If
    
    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaParametrosLog = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "VerificaParametrosLog > "
End Function

Private Sub LimpaStatus()
    On Error GoTo ErroRotina
    
    iStatus = SemRequisicao
    sMensagem = Empty
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "LimpaStatus > "
End Sub

