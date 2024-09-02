VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSis_Log"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lPK_LOG As Variant
Private lFK_LOG As Variant
Private sVCH_IDENTIFICACAO As Variant
Private sVCH_OBJETO As Variant
Private sVCH_ROTINA_FONTE As Variant
Private sVCH_DESCRICAO As Variant
Private dDTT_GERACAO As Variant
Private sVCH_NOME_COMPUTADOR As Variant
Private sVCH_NOME_USUARIO_LOCAL As Variant
Private sVCH_CODIGO_APLIC As Variant
Private sVCH_NOME_APLIC As Variant
Private lINT_CODIGO_USUARIO_APLIC As Variant
Private sVCH_NOME_USUARIO_APLIC As Variant
Private sVCH_CONEXAO As Variant

Private iAcao             As ADRRBR_SIS_Basico.eAcao
Private sColunasConsulta  As String
Private sOrdemConsulta    As String
Private sCondicao         As String
Private cConexao          As Object
Private rsRegistros       As Object
Private iStatus           As ADRRBR_SIS_Basico.eStatus
Private sMensagem         As String

Private oBasico           As New ADRRBR_SIS_Basico.clsSIS_Basico
Private rsRegistrosTemp   As Object
Private Estrutura_SQL     As Object
Private sSQL              As String

Dim sTipoExclusao         As String

Private Const sTabelaClasse = "Sis_Log"

Private Sub Class_Initialize()
    lPK_LOG = "NULL"
    lFK_LOG = "NULL"
    sVCH_IDENTIFICACAO = "NULL"
    sVCH_OBJETO = "NULL"
    sVCH_ROTINA_FONTE = "NULL"
    sVCH_DESCRICAO = "NULL"
    dDTT_GERACAO = "NULL"
    sVCH_NOME_COMPUTADOR = "NULL"
    sVCH_NOME_USUARIO_LOCAL = "NULL"
    sVCH_CODIGO_APLIC = "NULL"
    sVCH_NOME_APLIC = "NULL"
    lINT_CODIGO_USUARIO_APLIC = "NULL"
    sVCH_NOME_USUARIO_APLIC = "NULL"
    sVCH_CONEXAO = "NULL"
    
    LimpaStatus
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
    Set Estrutura_SQL = Nothing
End Sub

Public Property Let PK_LOG(ByVal vNewValue As Variant)
    lPK_LOG = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get PK_LOG() As Variant
    PK_LOG = lPK_LOG
End Property

Public Property Let FK_LOG(ByVal vNewValue As Variant)
    lFK_LOG = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FK_LOG() As Variant
    FK_LOG = lFK_LOG
End Property

Public Property Let VCH_IDENTIFICACAO(ByVal vNewValue As Variant)
    sVCH_IDENTIFICACAO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_IDENTIFICACAO() As Variant
    VCH_IDENTIFICACAO = sVCH_IDENTIFICACAO
End Property

Public Property Let VCH_OBJETO(ByVal vNewValue As Variant)
    sVCH_OBJETO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_OBJETO() As Variant
    VCH_OBJETO = sVCH_OBJETO
End Property

Public Property Let VCH_ROTINA_FONTE(ByVal vNewValue As Variant)
    sVCH_ROTINA_FONTE = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_ROTINA_FONTE() As Variant
    VCH_ROTINA_FONTE = sVCH_ROTINA_FONTE
End Property

Public Property Let VCH_DESCRICAO(ByVal vNewValue As Variant)
    sVCH_DESCRICAO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_DESCRICAO() As Variant
    VCH_DESCRICAO = sVCH_DESCRICAO
End Property

Public Property Let DTT_GERACAO(ByVal vNewValue As Variant)
    dDTT_GERACAO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get DTT_GERACAO() As Variant
    DTT_GERACAO = dDTT_GERACAO
End Property

Public Property Let VCH_NOME_COMPUTADOR(ByVal vNewValue As Variant)
    sVCH_NOME_COMPUTADOR = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_NOME_COMPUTADOR() As Variant
    VCH_NOME_COMPUTADOR = sVCH_NOME_COMPUTADOR
End Property

Public Property Let VCH_NOME_USUARIO_LOCAL(ByVal vNewValue As Variant)
    sVCH_NOME_USUARIO_LOCAL = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_NOME_USUARIO_LOCAL() As Variant
    VCH_NOME_USUARIO_LOCAL = sVCH_NOME_USUARIO_LOCAL
End Property

Public Property Let VCH_CODIGO_APLIC(ByVal vNewValue As Variant)
    sVCH_CODIGO_APLIC = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_CODIGO_APLIC() As Variant
    VCH_CODIGO_APLIC = sVCH_CODIGO_APLIC
End Property

Public Property Let VCH_NOME_APLIC(ByVal vNewValue As Variant)
    sVCH_NOME_APLIC = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_NOME_APLIC() As Variant
    VCH_NOME_APLIC = sVCH_NOME_APLIC
End Property

Public Property Let INT_CODIGO_USUARIO_APLIC(ByVal vNewValue As Variant)
    lINT_CODIGO_USUARIO_APLIC = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get INT_CODIGO_USUARIO_APLIC() As Variant
    INT_CODIGO_USUARIO_APLIC = lINT_CODIGO_USUARIO_APLIC
End Property

Public Property Let VCH_NOME_USUARIO_APLIC(ByVal vNewValue As Variant)
    sVCH_NOME_USUARIO_APLIC = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_NOME_USUARIO_APLIC() As Variant
    VCH_NOME_USUARIO_APLIC = sVCH_NOME_USUARIO_APLIC
End Property

Public Property Let VCH_CONEXAO(ByVal vNewValue As Variant)
    sVCH_CONEXAO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_CONEXAO() As Variant
    VCH_CONEXAO = sVCH_CONEXAO
End Property

Public Property Let Acao(ByVal vNewValue As eAcao)
    iAcao = vNewValue
End Property
Public Property Get Acao() As eAcao
    Acao = iAcao
End Property

Public Property Let ColunasConsulta(ByVal vNewValue As String)
    sColunasConsulta = vNewValue
End Property
Public Property Get ColunasConsulta() As String
    ColunasConsulta = sColunasConsulta
End Property

Public Property Let OrdemConsulta(ByVal vNewValue As String)
    sOrdemConsulta = vNewValue
End Property
Public Property Get OrdemConsulta() As String
    OrdemConsulta = sOrdemConsulta
End Property

Public Property Let Condicao(ByVal vNewValue As String)
    sCondicao = vNewValue
End Property
Public Property Get Condicao() As String
    Condicao = sCondicao
End Property

Public Property Set Conexao(ByVal vNewValue As Object)
    Set cConexao = vNewValue
End Property
Public Property Get Conexao() As Object
    Set Conexao = cConexao
End Property

Public Property Get Registros() As Object
    Set Registros = rsRegistros
End Property

Public Property Get Status() As eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Sub Atualizar_BD()
    On Error GoTo ErroRotina

    oBasico.Geral.Pausa 1

    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    If Not VerificaPropriedadesAtualizacao Then Exit Sub
    
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            If cConexao.Numero_AtualizacaoBD_Lote = 0 Then
                If iAcao <> Excluir Or (iAcao = Excluir And sTipoExclusao = "ESP") Then
                    If Not VerificaEstruturaTabela_SQL_Server Then Exit Sub
                End If
            End If
            
            iStatus = Atualiza_SQL_Server
            
        Case Access
            iStatus = Atualiza_Access
    End Select
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O M�TODO < cls" & sTabelaClasse & "." & "Atualizar_BD > "
End Sub

Public Sub Consultar_BD()
    On Error GoTo ErroRotina

    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    If Not VerificaPropriedadesConsulta Then Exit Sub
    
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            iStatus = Consulta_SQL_Server
            
        Case Access
            iStatus = Consulta_Access
    End Select
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O M�TODO < cls" & sTabelaClasse & "." & "Consultar_BD > "
End Sub

Private Function VerificaPropriedadesGerais() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesGerais = True
    
    If cConexao Is Nothing Then
        iStatus = Erro
        sMensagem = "Informe o objeto de CONEX�O!"
        VerificaPropriedadesGerais = False
        Exit Function
    End If
    
    If Not cConexao.conectado Then
        iStatus = Erro
        sMensagem = "O objeto de conex�o informado N�O est� CONECTADO!"
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "VerificaPropriedadesGerais > "
End Function

Private Function VerificaPropriedadesAtualizacao() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesAtualizacao = True
    
    If iAcao <> Incluir And iAcao <> Alterar And iAcao <> Excluir Then
        iStatus = Erro
        sMensagem = "Informe a A��O para atualiza��o dos dados!"
        VerificaPropriedadesAtualizacao = False
        Exit Function
    End If
    
    If iAcao = Excluir Then
        If Trim(UCase(lPK_LOG)) <> "NULL" And Trim(lPK_LOG) <> Empty And Trim(lPK_LOG) <> "0" Then
            If sCondicao <> Empty Then
                iStatus = Erro
                sMensagem = "Para a a��o EXCLUIR, se informar a CHAVE DO REGISTRO n�o informar a CONDI��O, para atualiza��o dos dados!"
                VerificaPropriedadesAtualizacao = False
                Exit Function
            End If
            sTipoExclusao = "ESP"
        Else
            sTipoExclusao = "CON"
        End If
    End If
    
    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaPropriedadesAtualizacao = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "VerificaPropriedadesAtualizacao > "
End Function

Private Function VerificaPropriedadesConsulta() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesConsulta = True
    
    If iAcao <> Consultar Then
        iStatus = Erro
        sMensagem = "Informe a A��O para consulta dos dados!"
        VerificaPropriedadesConsulta = False
        Exit Function
    End If
    
    Exit Function
    
ErroRotina:
    iStatus = Erro
    VerificaPropriedadesConsulta = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "VerificaPropriedadesConsulta > "
End Function

Private Function VerificaEstruturaTabela_SQL_Server() As Boolean
    Dim iColuna   As Integer
    Dim sConteudo As String
    
    On Error GoTo ErroRotina
    
    VerificaEstruturaTabela_SQL_Server = True
    
    Set Estrutura_SQL = CreateObject("ADRRBR_SQL_Estrutura.clsSQL_Estrutura")
    
    With Estrutura_SQL
        Set .Conexao = cConexao
        .tabela = sTabelaClasse
        .CarregaEstruturaTabela
        GoSub VerificaStatus
        
        .AcaoValidar = iAcao
        
        .Coluna = "PK_LOG"
        .Conteudo = lPK_LOG
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "FK_LOG"
        .Conteudo = lFK_LOG
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_IDENTIFICACAO"
        .Conteudo = sVCH_IDENTIFICACAO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_OBJETO"
        .Conteudo = sVCH_OBJETO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_ROTINA_FONTE"
        .Conteudo = sVCH_ROTINA_FONTE
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_DESCRICAO"
        .Conteudo = sVCH_DESCRICAO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "DTT_GERACAO"
        .Conteudo = dDTT_GERACAO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_NOME_COMPUTADOR"
        .Conteudo = sVCH_NOME_COMPUTADOR
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_NOME_USUARIO_LOCAL"
        .Conteudo = sVCH_NOME_USUARIO_LOCAL
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_CODIGO_APLIC"
        .Conteudo = sVCH_CODIGO_APLIC
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_NOME_APLIC"
        .Conteudo = sVCH_NOME_APLIC
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "INT_CODIGO_USUARIO_APLIC"
        .Conteudo = lINT_CODIGO_USUARIO_APLIC
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_NOME_USUARIO_APLIC"
        .Conteudo = sVCH_NOME_USUARIO_APLIC
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_CONEXAO"
        .Conteudo = sVCH_CONEXAO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .ValidaIntegridade
        GoSub VerificaStatus
    End With

    Exit Function
    
VerificaStatus:
    iStatus = Estrutura_SQL.Status
    If iStatus = Erro Then
        sMensagem = Estrutura_SQL.Mensagem
        VerificaEstruturaTabela_SQL_Server = False
        Exit Function
    End If
Return

ErroRotina:
    iStatus = Erro
    VerificaEstruturaTabela_SQL_Server = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "VerificaEstruturaTabela_SQL_Server > "
End Function

Private Function Atualiza_SQL_Server() As eStatus
    Dim sCamposConteudos As String

    On Error GoTo ErroRotina
    
    Atualiza_SQL_Server = Erro
    
    Select Case iAcao
        Case Incluir
            'Prepara Campos Para Inclus�o
            sCamposConteudos = Empty
            If Trim(UCase(lFK_LOG)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "FK_LOG=" & lFK_LOG & ","
            If Trim(UCase(sVCH_IDENTIFICACAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_IDENTIFICACAO=" & sVCH_IDENTIFICACAO & ","
            If Trim(UCase(sVCH_OBJETO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_OBJETO=" & sVCH_OBJETO & ","
            If Trim(UCase(sVCH_ROTINA_FONTE)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_ROTINA_FONTE=" & sVCH_ROTINA_FONTE & ","
            If Trim(UCase(sVCH_DESCRICAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_DESCRICAO=" & sVCH_DESCRICAO & ","
            If Trim(UCase(dDTT_GERACAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "DTT_GERACAO=" & oBasico.DataHora.FormataDataGravar(dDTT_GERACAO) & ","
            If Trim(UCase(sVCH_NOME_COMPUTADOR)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_COMPUTADOR=" & sVCH_NOME_COMPUTADOR & ","
            If Trim(UCase(sVCH_NOME_USUARIO_LOCAL)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_USUARIO_LOCAL=" & sVCH_NOME_USUARIO_LOCAL & ","
            If Trim(UCase(sVCH_CODIGO_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_CODIGO_APLIC=" & sVCH_CODIGO_APLIC & ","
            If Trim(UCase(sVCH_NOME_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_APLIC=" & sVCH_NOME_APLIC & ","
            If Trim(UCase(lINT_CODIGO_USUARIO_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "INT_CODIGO_USUARIO_APLIC=" & lINT_CODIGO_USUARIO_APLIC & ","
            If Trim(UCase(sVCH_NOME_USUARIO_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_USUARIO_APLIC=" & sVCH_NOME_USUARIO_APLIC & ","
            If Trim(UCase(sVCH_CONEXAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_CONEXAO=" & sVCH_CONEXAO & " "
            
            'Executa Procedure Para Inclus�o
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_GRAVA_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Operacao='INC',"
            sSQL = sSQL & "@CamposConteudos=" & oBasico.Geral.Aspas(sCamposConteudos, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & " "
            
            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO N�O IDENTIFICADO NA EXECU��O DA PROCEDURE DBO.PRC_SQL_GRAVA_REGISTRO!"
                Exit Function
            Else
                If cConexao.Numero_AtualizacaoBD_Lote > 0 Then
                    Atualiza_SQL_Server = Sucesso
                    sMensagem = "O comando para atualiza��o em lote foi inserido com sucesso!"
                    Exit Function
                Else
                    If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
                        sMensagem = rsRegistrosTemp!Mensagem
                        Exit Function
                    Else
                        lPK_LOG = rsRegistrosTemp!PK_LOG
                        Atualiza_SQL_Server = Sucesso
                        sMensagem = "O registro foi inserido com sucesso!"
                        Exit Function
                    End If
                End If
            End If
        
        Case Alterar
            'Prepara Campos Para Altera��o
            sCamposConteudos = Empty
            If Trim(UCase(lFK_LOG)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "FK_LOG=" & lFK_LOG & ","
            If Trim(UCase(sVCH_IDENTIFICACAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_IDENTIFICACAO=" & sVCH_IDENTIFICACAO & ","
            If Trim(UCase(sVCH_OBJETO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_OBJETO=" & sVCH_OBJETO & ","
            If Trim(UCase(sVCH_ROTINA_FONTE)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_ROTINA_FONTE=" & sVCH_ROTINA_FONTE & ","
            If Trim(UCase(sVCH_DESCRICAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_DESCRICAO=" & sVCH_DESCRICAO & ","
            If Trim(UCase(dDTT_GERACAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "DTT_GERACAO=" & oBasico.DataHora.FormataDataGravar(dDTT_GERACAO) & ","
            If Trim(UCase(sVCH_NOME_COMPUTADOR)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_COMPUTADOR=" & sVCH_NOME_COMPUTADOR & ","
            If Trim(UCase(sVCH_NOME_USUARIO_LOCAL)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_USUARIO_LOCAL=" & sVCH_NOME_USUARIO_LOCAL & ","
            If Trim(UCase(sVCH_CODIGO_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_CODIGO_APLIC=" & sVCH_CODIGO_APLIC & ","
            If Trim(UCase(sVCH_NOME_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_APLIC=" & sVCH_NOME_APLIC & ","
            If Trim(UCase(lINT_CODIGO_USUARIO_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "INT_CODIGO_USUARIO_APLIC=" & lINT_CODIGO_USUARIO_APLIC & ","
            If Trim(UCase(sVCH_NOME_USUARIO_APLIC)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_USUARIO_APLIC=" & sVCH_NOME_USUARIO_APLIC & ","
            If Trim(UCase(sVCH_CONEXAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_CONEXAO=" & sVCH_CONEXAO & " "
            
            'Executa Procedure Para Altera��o
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_GRAVA_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Operacao='ALT',"
            sSQL = sSQL & "@CamposConteudos=" & oBasico.Geral.Aspas(sCamposConteudos, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & ","
            sSQL = sSQL & "@IDRegistro=" & lPK_LOG 

            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO N�O IDENTIFICADO NA EXECU��O DA PROCEDURE DBO.PRC_SQL_GRAVA_REGISTRO!"
                Exit Function
            Else
                If cConexao.Numero_AtualizacaoBD_Lote > 0 Then
                    Atualiza_SQL_Server = Sucesso
                    sMensagem = "O comando para atualiza��o em lote foi inserido com sucesso!"
                    Exit Function
                Else
                    If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
                        sMensagem = rsRegistrosTemp!Mensagem
                        Exit Function
                    Else
                        Atualiza_SQL_Server = Sucesso
                        sMensagem = "O registro foi alterado com sucesso!"
                        Exit Function
                    End If
                End If
            End If
            
        Case Excluir
            'Executa Procedure Para Exclus�o
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_EXCLUI_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Tipo=" & oBasico.Geral.Aspas(sTipoExclusao, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & ","
            If sTipoExclusao = "ESP" Then
                sSQL = sSQL & "@IDRegistro=" & lPK_LOG 
            ElseIf sTipoExclusao = "CON" Then
                sSQL = sSQL & "@Condicoes=" & oBasico.Geral.Aspas(sCondicao, AspasSimples) & " "
            End If
            
            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO N�O IDENTIFICADO NA EXECU��O DA PROCEDURE DBO.PRC_SQL_EXCLUI_REGISTRO!"
                Exit Function
            Else
                If cConexao.Numero_AtualizacaoBD_Lote > 0 Then
                    Atualiza_SQL_Server = Sucesso
                    sMensagem = "O comando para atualiza��o em lote foi inserido com sucesso!"
                    Exit Function
                Else
                    If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
                        If InStr(1, UCase(rsRegistrosTemp.Fields(0).Value), "SUCESSO") Then
                            Atualiza_SQL_Server = Sucesso
                            If sTipoExclusao = "ESP" Then
                                sMensagem = "O registro foi exclu�do com sucesso!"
                            Else
                                sMensagem = "O(s) registro(s) foi(ram) exclu�do(s) com sucesso!"
                            End If
                        Else
                            sMensagem = rsRegistrosTemp.Fields(0).Value
                        End If
                        Exit Function
                    Else
                        sMensagem = "OCORREU UM ERRO N�O IDENTIFICADO NA EXECU��O DA PROCEDURE DBO.PRC_SQL_EXCLUI_REGISTRO!"
                        Exit Function
                    End If
                End If
            End If
    End Select
    
    Exit Function
    
ErroRotina:
    If Err.Number = 424 Then
        sMensagem = "N�o foi poss�vel atualizar o(s) registro(s) da tabela < " & sTabelaClasse & " > ." & vbLf
        sMensagem = sMensagem & "Provavelmente para algum(s) registro(s) a ser(em) atualizado(s), existe(m) registro(s) correspondente(s) violado(s) na(s) tabela(s) dependente(s)!"
    Else
        sMensagem = "Ocorreu o erro: " & vbLf
        sMensagem = sMensagem & Err.Number & vbLf
        sMensagem = sMensagem & Err.Description & vbLf & vbLf
        sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "Atualiza_SQL_Server > "
    End If
End Function

Private Function Atualiza_Access() As eStatus
    On Error GoTo ErroRotina
    
    Atualiza_Access = Erro

    Exit Function
    
ErroRotina:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "Atualiza_Access > "
End Function

Private Function Consulta_SQL_Server() As eStatus
    On Error GoTo ErroRotina
    
    Consulta_SQL_Server = Erro
    
    sSQL = Empty
    sSQL = sSQL & "DBO.PRC_SQL_PESQUISA_REGISTROS "
    sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ", "
    sSQL = sSQL & "@Colunas=" & oBasico.Geral.Aspas(sColunasConsulta, AspasSimples) & ", "
    sSQL = sSQL & "@ColunasOrdem=" & oBasico.Geral.Aspas(sOrdemConsulta, AspasSimples) & ", "
    sSQL = sSQL & "@Condicoes=" & oBasico.Geral.Aspas(sCondicao, AspasSimples) & " "
    
    Set rsRegistros = Nothing
    Set rsRegistros = cConexao.AbreRS(sSQL, True)
    
    If rsRegistros.EOF Then
        Consulta_SQL_Server = NaoEncontrado
        sMensagem = "Nenhum registro foi encontrado com os crit�rios indicados!"
        Exit Function
    Else
        If Trim(UCase(rsRegistros.Fields(0).Name)) = "MENSAGEM" Then
            sMensagem = rsRegistros!Mensagem
            Exit Function
        End If
    End If
    
    Consulta_SQL_Server = Encontrado
    
    Exit Function
    
ErroRotina:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "Consulta_SQL_Server > "
End Function

Private Function Consulta_Access() As eStatus
    On Error GoTo ErroRotina

    Consulta_Access = Erro

    Exit Function
    
ErroRotina:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "Consulta_Access > "
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sTabelaClasse & "." & "LimpaStatus > "
End Sub