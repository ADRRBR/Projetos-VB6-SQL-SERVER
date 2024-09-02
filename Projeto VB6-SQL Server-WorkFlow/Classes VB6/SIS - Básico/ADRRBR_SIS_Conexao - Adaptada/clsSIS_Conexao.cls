VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Conexao"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iTipoBancoDados            As ADRRBR_SIS_Basico.eTipoBancoDados
Private cConexao                   As Object
Private sBancoDados                As String
Private sUsuario                   As String
Private sSenha                     As String
Private sServidor                  As String
Private sCaminhoMDB                As String
Private sCaminhoFisicoInternet     As String
Private bVerificaAplicacaoInternet As Boolean
Private lPK_ATUBD                  As Long
Private bConectado                 As Boolean
Private bConexaoUsuarioWindows     As Boolean
Private sFonteConexao              As String
Private iStatus                    As ADRRBR_SIS_Basico.eStatus
Private sMensagem                  As String

Private oBasico                    As New ADRRBR_SIS_Basico.clsSIS_Basico
Private sCnx                       As String
Private sSQL                       As String
Private rsRegistrosTemp            As Object

Private Const sClasse = "SIS_Conexao"

Private Sub Class_Initialize()
    LimpaStatus
    bConexaoUsuarioWindows = False
    bConectado = False
    lPK_ATUBD = 0
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
End Sub

Public Property Let TipoBancoDados(ByVal vNewValue As eTipoBancoDados)
    iTipoBancoDados = vNewValue
End Property
Public Property Get TipoBancoDados() As eTipoBancoDados
    TipoBancoDados = iTipoBancoDados
End Property

Public Property Get Conexao() As Object
    Set Conexao = cConexao
End Property

Public Property Let BancoDados(ByVal vNewValue As String)
    sBancoDados = vNewValue
End Property
Public Property Get BancoDados() As String
    BancoDados = sBancoDados
End Property

Public Property Let Servidor(ByVal vNewValue As String)
    sServidor = vNewValue
End Property
Public Property Get Servidor() As String
    Servidor = sServidor
End Property

Public Property Let CaminhoMDB(ByVal vNewValue As String)
    sCaminhoMDB = vNewValue
End Property
Public Property Get CaminhoMDB() As String
    CaminhoMDB = sCaminhoMDB
End Property

Public Property Get CaminhoFisicoInternet() As String
    CaminhoFisicoInternet = sCaminhoFisicoInternet
End Property

Public Property Get VerificaAplicacaoInternet() As Boolean
    VerificaAplicacaoInternet = bVerificaAplicacaoInternet
End Property

Public Property Let Usuario(ByVal vNewValue As String)
    sUsuario = vNewValue
End Property
Public Property Get Usuario() As String
    Usuario = "*****"
End Property

Public Property Let Senha(ByVal vNewValue As String)
    sSenha = vNewValue
End Property
Public Property Get Senha() As String
    Senha = "*****"
End Property

Public Property Let ConexaoUsuarioWindows(ByVal vNewValue As Boolean)
    bConexaoUsuarioWindows = vNewValue
End Property
Public Property Get ConexaoUsuarioWindows() As Boolean
    ConexaoUsuarioWindows = bConexaoUsuarioWindows
End Property

Public Property Get Numero_AtualizacaoBD_Lote() As Long
    Numero_AtualizacaoBD_Lote = lPK_ATUBD
End Property

Public Property Get Conectado() As Boolean
    Conectado = bConectado
End Property

Public Property Get FonteConexao() As String
    FonteConexao = sFonteConexao
End Property

Public Property Get Status() As eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Sub LoginSeguranca()
    Dim iStatusAux   As eStatus
    Dim sMensagemAux As String
    
    On Error GoTo ErroRotina
        
    LimpaStatus
    
    RecuperaParametrosConexao
    
    If iStatus <> Sucesso Then Exit Sub
    
    frmSeguranca.InicializaTela
    
    Select Case iTipoBancoDados
        Case SQL_Server
            oBasico.Geral.PosicionaCombo_Conteudo frmSeguranca.cmbTipoBD, "SQL Server"
            
            If bConexaoUsuarioWindows Then
                frmSeguranca.chkConexaoUsuarioWindows.Value = 1
            Else
                frmSeguranca.chkConexaoUsuarioWindows.Value = 0
            End If
            frmSeguranca.txtServidor.Text = sServidor
            frmSeguranca.txtBancoDados.Text = sBancoDados
            frmSeguranca.txtUsuario.Text = sUsuario
            
            If Trim(sCaminhoFisicoInternet) <> Empty Then
                frmSeguranca.chkInternet.Value = 1
                frmSeguranca.txtCaminhoFisico.Text = sCaminhoFisicoInternet
            Else
                frmSeguranca.chkInternet.Value = 0
                frmSeguranca.chkInternet_Click
            End If
            
        Case Access
            oBasico.Geral.PosicionaCombo_Conteudo frmSeguranca.cmbTipoBD, "Access"
            
            frmSeguranca.txtCaminhoMDB.Text = sCaminhoMDB
    End Select
    
    frmSeguranca.txtSenha.Text = Empty
    frmSeguranca.cmbTipoBD.Locked = True
    
    frmSeguranca.Show vbModal
    
    iStatus = frmSeguranca.Status
    sMensagem = frmSeguranca.Mensagem
    
    If frmSeguranca.Status = Sucesso Then
        iStatusAux = iStatus
        sMensagemAux = sMensagem
        
        RecuperaParametrosConexao
        
        If iStatus = Sucesso Then
            iStatus = iStatusAux
            sMensagem = sMensagemAux
        End If
    End If

    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Login > "
End Sub

Public Sub RecuperaParametrosConexao()
    Dim sIdentificacao              As String
    Dim sSubIdentificacao           As String
    Dim sConexaoUsuarioWindows      As String
    Dim sColunaAux                  As String
    Dim rsArqParamConexao           As ADODB.Recordset
    
    On Error GoTo ErroRotina
        
    LimpaStatus
    
    iStatus = Sucesso
    sMensagem = "PARÂMETROS DE CONEXÃO RECUPERADOS COM SUCESSO!"

    If Not VerificaTipoBancoDados Then Exit Sub

    '************************************************
    ' Verifica Conexão da Possível Aplicação Windows
    '************************************************
    bVerificaAplicacaoInternet = False

    sIdentificacao = oBasico.Geral.EncriptarDecriptar("ADRRBR", True)
    
    Select Case iTipoBancoDados
        Case SQL_Server
            sSubIdentificacao = oBasico.Geral.EncriptarDecriptar("SEGURANCA SQL SERVER", True)
            
            sServidor = oBasico.Geral.LeRegistroWindows( _
            sIdentificacao, _
            sSubIdentificacao, _
            oBasico.Geral.EncriptarDecriptar("SERVIDOR", True))

            sBancoDados = oBasico.Geral.LeRegistroWindows( _
            sIdentificacao, _
            sSubIdentificacao, _
            oBasico.Geral.EncriptarDecriptar("BANCO DE DADOS", True))

            sConexaoUsuarioWindows = oBasico.Geral.LeRegistroWindows( _
            sIdentificacao, _
            sSubIdentificacao, _
            oBasico.Geral.EncriptarDecriptar("CONEXAO USUARIO WINDOWS", True))
            
            sUsuario = oBasico.Geral.LeRegistroWindows( _
            sIdentificacao, _
            sSubIdentificacao, _
            oBasico.Geral.EncriptarDecriptar("USUARIO", True))

            'Decriptar as Informações
            sServidor = oBasico.Geral.EncriptarDecriptar(sServidor, False)
            sBancoDados = oBasico.Geral.EncriptarDecriptar(sBancoDados, False)
            sUsuario = oBasico.Geral.EncriptarDecriptar(sUsuario, False)
            sConexaoUsuarioWindows = oBasico.Geral.EncriptarDecriptar(sConexaoUsuarioWindows, False)
            If sConexaoUsuarioWindows = "1" Then
                bConexaoUsuarioWindows = True
            Else
                bConexaoUsuarioWindows = False
            End If
            
            'If sServidor = Empty And sBancoDados = Empty Then bVerificaAplicacaoInternet = True
            
        Case Access
            sSubIdentificacao = oBasico.Geral.EncriptarDecriptar("SEGURANCA ACCESS", True)
        
            sCaminhoMDB = oBasico.Geral.LeRegistroWindows( _
            sIdentificacao, _
            sSubIdentificacao, _
            oBasico.Geral.EncriptarDecriptar("CAMINHO MDB", True))
    
            'Decriptar as Informações
            sCaminhoMDB = oBasico.Geral.EncriptarDecriptar(sCaminhoMDB, False)
    
            If sCaminhoMDB = Empty Then bVerificaAplicacaoInternet = True
    End Select
    
    sSenha = oBasico.Geral.LeRegistroWindows( _
    sIdentificacao, _
    sSubIdentificacao, _
    oBasico.Geral.EncriptarDecriptar("SENHA", True))

    sCaminhoFisicoInternet = oBasico.Geral.LeRegistroWindows( _
    sIdentificacao, _
    sSubIdentificacao, _
    oBasico.Geral.EncriptarDecriptar("CAMINHO FISICO INTERNET", True))

    'Decriptar as Informações
    sSenha = oBasico.Geral.EncriptarDecriptar(sSenha, False)
    sCaminhoFisicoInternet = oBasico.Geral.EncriptarDecriptar(sCaminhoFisicoInternet, False)
    
    '************************************************
    'Verifica Conexão da Possível Aplicação Internet
    '************************************************
    If bVerificaAplicacaoInternet Then
        If Trim(sCaminhoFisicoInternet) = Empty Then sCaminhoFisicoInternet = sCaminhoFisicoInternetTemp
        If Dir(sCaminhoFisicoInternet & "\" & sArqParamConexao) = Empty Then GoSub Cancelamento: Exit Sub
        Set rsArqParamConexao = New ADODB.Recordset
        rsArqParamConexao.Open sCaminhoFisicoInternet & "\" & sArqParamConexao
        If rsArqParamConexao.EOF Then GoSub Cancelamento: Exit Sub
                    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("TIPO SEGURANCA", True)
        
        Select Case oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
            Case "SEGURANCA SQL SERVER"
                sColunaAux = oBasico.Geral.EncriptarDecriptar("SERVIDOR", True)
                sServidor = oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
    
                sColunaAux = oBasico.Geral.EncriptarDecriptar("BANCO DE DADOS", True)
                sBancoDados = oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
    
                sColunaAux = oBasico.Geral.EncriptarDecriptar("CONEXAO USUARIO WINDOWS", True)
                sConexaoUsuarioWindows = oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
    
                sColunaAux = oBasico.Geral.EncriptarDecriptar("USUARIO", True)
                sUsuario = oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
                
                If sServidor = Empty And sBancoDados = Empty Then GoSub Cancelamento: Exit Sub
            
            Case "SEGURANCA ACCESS"
                sColunaAux = oBasico.Geral.EncriptarDecriptar("CAMINHO MDB", True)
                sCaminhoMDB = oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
        
                If sCaminhoMDB = Empty Then GoSub Cancelamento: Exit Sub
        End Select
    
        sColunaAux = oBasico.Geral.EncriptarDecriptar("SENHA", True)
        sSenha = oBasico.Geral.EncriptarDecriptar(rsArqParamConexao.Fields(sColunaAux).Value, False)
    End If
    
    Set rsArqParamConexao = Nothing
    
    Exit Sub
    
Cancelamento:
    bVerificaAplicacaoInternet = False
    Set rsArqParamConexao = Nothing
    iStatus = Cancelado
    sMensagem = "NÃO FOI POSSÍVEL RECUPERAR OS PARÂMETROS DE CONEXÃO!"
Return
    
ErroRotina:
    Set rsArqParamConexao = Nothing
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "RecuperaUltimaConexao > "
End Sub

Public Sub Conecta()
    Dim sAux()  As String
    Dim iIndice As Integer
    
    On Error GoTo ErroRotina
    
    LimpaStatus
    bConectado = False
    
    If Not VerificaTipoBancoDados Then Exit Sub
    If Not VerificaPropriedadesConexao Then Exit Sub

    Select Case iTipoBancoDados
        Case SQL_Server
            iStatus = Conecta_SQLServer
            
        Case Access
            iStatus = Conecta_Access
    End Select
    
    If iStatus = Sucesso Then
        bConectado = True
        
        sAux() = Split(cConexao, ";")
        sFonteConexao = Empty
        
        For iIndice = 0 To UBound(sAux())
            If InStr(1, UCase(sAux(iIndice)), "PWD") > 0 Then
                sFonteConexao = sFonteConexao & "PWD=*****"
            Else
                sFonteConexao = sFonteConexao & sAux(iIndice)
            End If
            If iIndice < UBound(sAux()) Then sFonteConexao = sFonteConexao & ";"
        Next iIndice
    End If
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Conecta > "
End Sub

Public Function AbreRS(pSQL As Variant, Optional pSomenteLeitura As Boolean) As Variant
    Dim oRS As ADODB.Recordset
    
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Function
    
    Set oRS = Nothing
    Set oRS = New ADODB.Recordset
    
    oRS.CursorLocation = adUseClient
    
    If pSomenteLeitura Then
        oRS.Open pSQL, cConexao, adOpenStatic, adLockReadOnly
    Else
        oRS.Open pSQL, cConexao, adOpenDynamic, adLockOptimistic
    End If

    Set AbreRS = oRS

    iStatus = Sucesso
    
    Exit Function
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "AbreRS > "
End Function

Public Sub ExecutaQuery(pSQL As Variant, Optional pRegistrosAfetados As Variant)
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    
    cConexao.Execute pSQL, pRegistrosAfetados

    iStatus = Sucesso
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "ExecutaQuery > "
End Sub

Public Sub Inicializa_AtualizacaoBD_Lote()
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    
    iStatus = Erro
    
    If lPK_ATUBD > 0 Then
        sMensagem = "O lote de atualização de banco de dados já foi inicializado!"
        Exit Sub
    End If
    
    'Executa Procedure Para Inicialização do Lote de Atualização
    sSQL = Empty
    sSQL = sSQL & "DBO.PRC_SIS_ATUALIZA_BD_LOTE "
    sSQL = sSQL & "@Tipo = 1," 'Inicializa o Lote de Atualização, Retornando o Número da Atualização Gerado - PK_ATUBD
    sSQL = sSQL & "@VCH_Nome_Computador = " & oBasico.Geral.Aspas(oBasico.Geral.NomeComputador, AspasSimples) & ","
    sSQL = sSQL & "@VCH_Nome_Usuario_Local = " & oBasico.Geral.Aspas(oBasico.Geral.UsuarioLocal, AspasSimples)
    
    Set rsRegistrosTemp = Nothing
    Set rsRegistrosTemp = AbreRS(sSQL, True)
    If rsRegistrosTemp.EOF Then
        sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SIS_ATUALIZA_BD_LOTE!"
    Else
        If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
            sMensagem = rsRegistrosTemp!Mensagem
        Else
            lPK_ATUBD = rsRegistrosTemp!PK_ATUBD
            iStatus = Sucesso
            sMensagem = "O lote de atualização de banco de dados foi inicializado com sucesso!"
        End If
    End If
    
    Exit Sub
    
ErroRotina:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Inicializa_AtualizacaoBD_Lote > "
End Sub

Public Sub Efetiva_AtualizacaoBD_Lote()
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    
    iStatus = Erro
            
    If lPK_ATUBD = 0 Then
        sMensagem = "O lote de atualização de banco de dados não foi inicializado!"
        Exit Sub
    End If
            
    'Executa Procedure Para Efetivação do Lote de Atualização
    sSQL = Empty
    sSQL = sSQL & "DBO.PRC_SIS_ATUALIZA_BD_LOTE "
    sSQL = sSQL & "@Tipo = 3," 'Efetiva o Lote de Atualização, Conforme o Número de Atualização Informado no Parâmetro < @PK_AtuBD >) e (Finaliza o Lote de Atualização, Conforme o Número de Atualização Informado no Parâmetro < @PK_AtuBD >
    sSQL = sSQL & "@PK_AtuBD = " & lPK_ATUBD
    
    Set rsRegistrosTemp = Nothing
    Set rsRegistrosTemp = AbreRS(sSQL, True)
    If rsRegistrosTemp.EOF Then
        sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SIS_ATUALIZA_BD_LOTE!"
    Else
        If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
            If InStr(1, UCase(rsRegistrosTemp.Fields(0).Value), "SUCESSO") Then
                lPK_ATUBD = 0
                iStatus = Sucesso
                sMensagem = "O lote de atualização de banco de dados foi efetivado com sucesso!"
            Else
                sMensagem = rsRegistrosTemp.Fields(0).Value
            End If
        Else
            sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SIS_ATUALIZA_BD_LOTE!"
        End If
    End If
    
    Exit Sub
    
ErroRotina:
    If Err.Number = 424 Then
        sMensagem = "Não foi possível efetivar o(s) comando(s) do lote de atualização número < " & Numero_AtualizacaoBD_Lote & " > ." & vbLf
        sMensagem = sMensagem & "Provavelmente para algum(s) registro(s) no lote de manutenção, existe(m) registro(s) correspondente(s) violado(s) na(s) tabela(s) dependente(s)!"
    Else
        sMensagem = "Ocorreu o erro: " & vbLf
        sMensagem = sMensagem & Err.Number & vbLf
        sMensagem = sMensagem & Err.Description & vbLf & vbLf
        sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Efetiva_AtualizacaoBD_Lote > "
    End If
End Sub

Public Sub Finaliza_AtualizacaoBD_Lote()
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    
    iStatus = Erro
            
    If lPK_ATUBD = 0 Then
        sMensagem = "O lote de atualização de banco de dados não foi inicializado!"
        Exit Sub
    End If
            
    'Executa Procedure Para Finalização do Lote de Atualização
    sSQL = Empty
    sSQL = sSQL & "DBO.PRC_SIS_ATUALIZA_BD_LOTE "
    sSQL = sSQL & "@Tipo = 4," 'Finaliza o Lote de Atualização, Conforme o Número de Atualização Informado no Parâmetro < @PK_AtuBD >
    sSQL = sSQL & "@PK_AtuBD = " & lPK_ATUBD
    
    Set rsRegistrosTemp = Nothing
    Set rsRegistrosTemp = AbreRS(sSQL, True)
    If rsRegistrosTemp.EOF Then
        sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SIS_ATUALIZA_BD_LOTE!"
    Else
        If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
            If InStr(1, UCase(rsRegistrosTemp.Fields(0).Value), "SUCESSO") Then
                lPK_ATUBD = 0
                iStatus = Sucesso
                sMensagem = "O lote de atualização de banco de dados foi finalizado com sucesso!"
            Else
                sMensagem = rsRegistrosTemp.Fields(0).Value
            End If
        Else
            sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SIS_ATUALIZA_BD_LOTE!"
        End If
    End If
    
    Exit Sub
    
ErroRotina:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Finaliza_AtualizacaoBD_Lote > "
End Sub

Private Function VerificaTipoBancoDados() As Boolean
    On Error GoTo ErroRotina
    
    VerificaTipoBancoDados = True
    
    If iTipoBancoDados <> SQL_Server And iTipoBancoDados <> Access Then
        iStatus = Erro
        sMensagem = "Informe o TIPO de BANCO DE DADOS!"
        VerificaTipoBancoDados = False
        Exit Function
    End If

    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaTipoBancoDados = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "VerificaTipoBancoDados > "
End Function

Private Function VerificaPropriedadesConexao() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesConexao = True
    
    Select Case iTipoBancoDados
        Case SQL_Server
            If Trim(sServidor) = Empty Then
                iStatus = Erro
                sMensagem = "Informe o SERVIDOR SQL Server!"
                VerificaPropriedadesConexao = False
                Exit Function
            End If
            If Trim(sBancoDados) = Empty Then
                iStatus = Erro
                sMensagem = "Informe o BANCO DE DADOS SQL Server"
                VerificaPropriedadesConexao = False
                Exit Function
            End If
            If bConexaoUsuarioWindows Then
                If Trim(sUsuario) <> Empty Then
                    iStatus = Erro
                    sMensagem = "Para conexão com o usuário do Windows, não informe o USUÁRIO do BANCO DE DADOS SQL Server"
                    VerificaPropriedadesConexao = False
                    Exit Function
                End If
                If Trim(sSenha) <> Empty Then
                    iStatus = Erro
                    sMensagem = "Para conexão com o usuário do Windows, não informe a SENHA do USUÁRIO do BANCO DE DADOS SQL Server"
                    VerificaPropriedadesConexao = False
                    Exit Function
                End If
            Else
                If Trim(sUsuario) = Empty Then
                    iStatus = Erro
                    sMensagem = "Informe o USUÁRIO do BANCO DE DADOS SQL Server"
                    VerificaPropriedadesConexao = False
                    Exit Function
                End If
            End If
            
        Case Access
            If bConexaoUsuarioWindows Then
                iStatus = Erro
                sMensagem = "O BANCO DE DADOS Access não deve ser conectado com o usuário do windows!"
                VerificaPropriedadesConexao = False
                Exit Function
            End If
            If Trim(sCaminhoMDB) = Empty Then
                iStatus = Erro
                sMensagem = "Informe o CAMINHO do arquivo Access (.mdb)!"
                VerificaPropriedadesConexao = False
                Exit Function
            End If
    End Select

    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaPropriedadesConexao = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "VerificaPropriedadesConexao > "
End Function

Private Function VerificaPropriedadesGerais() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesGerais = True
    
    If cConexao Is Nothing Then
        iStatus = Erro
        sMensagem = "NÃO existe CONEXÃO com banco de dados!"
        VerificaPropriedadesGerais = False
        Exit Function
    End If
    
    If Not bConectado Then
        iStatus = Erro
        sMensagem = "O objeto de conexão NÃO está CONECTADO!"
        VerificaPropriedadesGerais = False
        Exit Function
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

Private Function Conecta_SQLServer() As ADRRBR_SIS_Basico.eStatus
    On Error GoTo ErroRotina
    
    sCnx = Empty
    sCnx = sCnx & "Provider=MSDASQL; Driver=SQL Server;"
    sCnx = sCnx & "Server=" & sServidor & ";"
    sCnx = sCnx & "Database=" & sBancoDados & ";"
    sCnx = sCnx & "Uid=" & sUsuario & ";"
    sCnx = sCnx & "Pwd=" & sSenha & ";"
    
    On Error GoTo 0
    
    On Error GoTo ErroCnxSQL
        Set cConexao = Nothing
        Set cConexao = New ADODB.Connection
        cConexao.Open sCnx
    On Error GoTo 0
    
    Conecta_SQLServer = Sucesso
    
    Exit Function
    
ErroCnxSQL:
    Conecta_SQLServer = Erro
    sMensagem = "Ocorreu o erro: " & vbLf & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "ao conectar-se ao SQL SERVER:" & vbLf & vbLf
    sMensagem = sMensagem & "SERVIDOR: " & sServidor & vbLf
    sMensagem = sMensagem & "BANCO DE DADOS: " & sBancoDados & vbLf
    sMensagem = sMensagem & "USUÁRIO: " & sUsuario
        
    Exit Function
        
ErroRotina:
    Conecta_SQLServer = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "Conecta_SQLServer > "
End Function

Private Function Conecta_Access() As ADRRBR_SIS_Basico.eStatus
    On Error GoTo ErroRotina
    
    sCnx = Empty
    sCnx = sCnx & "Driver={Microsoft Access Driver (*.mdb)}"
    sCnx = sCnx & ";Dbq=" & sCaminhoMDB
    sCnx = sCnx & ";Uid=" & sUsuario
    sCnx = sCnx & ";Pwd=" & sSenha

    On Error GoTo 0

    On Error GoTo ErroCnxACC
        Set cConexao = Nothing
        Set cConexao = New ADODB.Connection
        cConexao.Open sCnx
    On Error GoTo 0

    Conecta_Access = Sucesso

    Exit Function
    
ErroCnxACC:
    Conecta_Access = Erro
    sMensagem = "Ocorreu o erro: " & vbLf & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "ao conectar-se ao ACCESS:" & vbLf & vbLf
    sMensagem = sMensagem & "CAMINHO MDB: " & sCaminhoMDB & vbLf
    sMensagem = sMensagem & "USUÁRIO: " & sUsuario
                        
    Exit Function
        
ErroRotina:
    Conecta_Access = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "Conecta_Access > "
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
