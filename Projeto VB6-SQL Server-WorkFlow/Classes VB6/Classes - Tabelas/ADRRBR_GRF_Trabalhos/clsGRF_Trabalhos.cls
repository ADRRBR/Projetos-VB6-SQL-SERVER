VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGRF_Trabalhos"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lpk_trabalho As Variant
Private lfk_trabalho_alt_seq As Variant
Private lfk_trabalho_tipo As Variant
Private lfk_produto As Variant
Private lfk_cliente As Variant
Private lfk_representante As Variant
Private lfk_aprovacao_tipo As Variant
Private snum_pedido As Variant
Private snum_pedido_antigo As Variant
Private snum_pedido_novo As Variant
Private slead_time_programado As Variant
Private ddt_inclusao As Variant
Private ddt_alteracao As Variant
Private ddt_exclusao As Variant
Private lfk_filial As Variant

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

Private Const sTabelaClasse = "GRF_Trabalhos"

Private Sub Class_Initialize()
    lpk_trabalho = "NULL"
    lfk_trabalho_alt_seq = "NULL"
    lfk_trabalho_tipo = "NULL"
    lfk_produto = "NULL"
    lfk_cliente = "NULL"
    lfk_representante = "NULL"
    lfk_aprovacao_tipo = "NULL"
    snum_pedido = "NULL"
    snum_pedido_antigo = "NULL"
    snum_pedido_novo = "NULL"
    slead_time_programado = "NULL"
    ddt_inclusao = "NULL"
    ddt_alteracao = "NULL"
    ddt_exclusao = "NULL"
    lfk_filial = "NULL"
    
    LimpaStatus
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
    Set Estrutura_SQL = Nothing
End Sub

Public Property Let pk_trabalho(ByVal vNewValue As Variant)
    lpk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get pk_trabalho() As Variant
    pk_trabalho = lpk_trabalho
End Property

Public Property Let fk_trabalho_alt_seq(ByVal vNewValue As Variant)
    lfk_trabalho_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_trabalho_alt_seq() As Variant
    fk_trabalho_alt_seq = lfk_trabalho_alt_seq
End Property

Public Property Let fk_trabalho_tipo(ByVal vNewValue As Variant)
    lfk_trabalho_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_trabalho_tipo() As Variant
    fk_trabalho_tipo = lfk_trabalho_tipo
End Property

Public Property Let fk_produto(ByVal vNewValue As Variant)
    lfk_produto = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_produto() As Variant
    fk_produto = lfk_produto
End Property

Public Property Let fk_cliente(ByVal vNewValue As Variant)
    lfk_cliente = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_cliente() As Variant
    fk_cliente = lfk_cliente
End Property

Public Property Let fk_representante(ByVal vNewValue As Variant)
    lfk_representante = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_representante() As Variant
    fk_representante = lfk_representante
End Property

Public Property Let fk_aprovacao_tipo(ByVal vNewValue As Variant)
    lfk_aprovacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_aprovacao_tipo() As Variant
    fk_aprovacao_tipo = lfk_aprovacao_tipo
End Property

Public Property Let num_pedido(ByVal vNewValue As Variant)
    snum_pedido = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get num_pedido() As Variant
    num_pedido = snum_pedido
End Property

Public Property Let num_pedido_antigo(ByVal vNewValue As Variant)
    snum_pedido_antigo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get num_pedido_antigo() As Variant
    num_pedido_antigo = snum_pedido_antigo
End Property

Public Property Let num_pedido_novo(ByVal vNewValue As Variant)
    snum_pedido_novo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get num_pedido_novo() As Variant
    num_pedido_novo = snum_pedido_novo
End Property

Public Property Let lead_time_programado(ByVal vNewValue As Variant)
    slead_time_programado = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get lead_time_programado() As Variant
    lead_time_programado = slead_time_programado
End Property

Public Property Let dt_inclusao(ByVal vNewValue As Variant)
    ddt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get dt_inclusao() As Variant
    dt_inclusao = ddt_inclusao
End Property

Public Property Let dt_alteracao(ByVal vNewValue As Variant)
    ddt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get dt_alteracao() As Variant
    dt_alteracao = ddt_alteracao
End Property

Public Property Let dt_exclusao(ByVal vNewValue As Variant)
    ddt_exclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get dt_exclusao() As Variant
    dt_exclusao = ddt_exclusao
End Property

Public Property Let fk_filial(ByVal vNewValue As Variant)
    lfk_filial = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get fk_filial() As Variant
    fk_filial = lfk_filial
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
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sTabelaClasse & "." & "Atualizar_BD > "
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
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sTabelaClasse & "." & "Consultar_BD > "
End Sub

Private Function VerificaPropriedadesGerais() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesGerais = True
    
    If cConexao Is Nothing Then
        iStatus = Erro
        sMensagem = "Informe o objeto de CONEXÃO!"
        VerificaPropriedadesGerais = False
        Exit Function
    End If
    
    If Not cConexao.conectado Then
        iStatus = Erro
        sMensagem = "O objeto de conexão informado NÃO está CONECTADO!"
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
        sMensagem = "Informe a AÇÃO para atualização dos dados!"
        VerificaPropriedadesAtualizacao = False
        Exit Function
    End If
    
    If iAcao = Excluir Then
        If Trim(UCase(lpk_trabalho)) <> "NULL" And Trim(lpk_trabalho) <> Empty And Trim(lpk_trabalho) <> "0" Then
            If sCondicao <> Empty Then
                iStatus = Erro
                sMensagem = "Para a ação EXCLUIR, se informar a CHAVE DO REGISTRO não informar a CONDIÇÃO, para atualização dos dados!"
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
        sMensagem = "Informe a AÇÃO para consulta dos dados!"
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
        
        .Coluna = "pk_trabalho"
        .Conteudo = lpk_trabalho
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_trabalho_alt_seq"
        .Conteudo = lfk_trabalho_alt_seq
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_trabalho_tipo"
        .Conteudo = lfk_trabalho_tipo
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_produto"
        .Conteudo = lfk_produto
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_cliente"
        .Conteudo = lfk_cliente
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_representante"
        .Conteudo = lfk_representante
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_aprovacao_tipo"
        .Conteudo = lfk_aprovacao_tipo
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "num_pedido"
        .Conteudo = snum_pedido
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "num_pedido_antigo"
        .Conteudo = snum_pedido_antigo
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "num_pedido_novo"
        .Conteudo = snum_pedido_novo
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "lead_time_programado"
        .Conteudo = slead_time_programado
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "dt_inclusao"
        .Conteudo = ddt_inclusao
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "dt_alteracao"
        .Conteudo = ddt_alteracao
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "dt_exclusao"
        .Conteudo = ddt_exclusao
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "fk_filial"
        .Conteudo = lfk_filial
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
            'Prepara Campos Para Inclusão
            sCamposConteudos = Empty
            If Trim(UCase(lfk_trabalho_alt_seq)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_trabalho_alt_seq=" & lfk_trabalho_alt_seq & ","
            If Trim(UCase(lfk_trabalho_tipo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_trabalho_tipo=" & lfk_trabalho_tipo & ","
            If Trim(UCase(lfk_produto)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_produto=" & lfk_produto & ","
            If Trim(UCase(lfk_cliente)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_cliente=" & lfk_cliente & ","
            If Trim(UCase(lfk_representante)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_representante=" & lfk_representante & ","
            If Trim(UCase(lfk_aprovacao_tipo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_aprovacao_tipo=" & lfk_aprovacao_tipo & ","
            If Trim(UCase(snum_pedido)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "num_pedido=" & snum_pedido & ","
            If Trim(UCase(snum_pedido_antigo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "num_pedido_antigo=" & snum_pedido_antigo & ","
            If Trim(UCase(snum_pedido_novo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "num_pedido_novo=" & snum_pedido_novo & ","
            If Trim(UCase(slead_time_programado)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "lead_time_programado=" & slead_time_programado & ","
            If Trim(UCase(ddt_inclusao)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "dt_inclusao=" & oBasico.DataHora.FormataDataGravar(ddt_inclusao) & ","
            If Trim(UCase(ddt_alteracao)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "dt_alteracao=" & oBasico.DataHora.FormataDataGravar(ddt_alteracao) & ","
            If Trim(UCase(ddt_exclusao)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "dt_exclusao=" & oBasico.DataHora.FormataDataGravar(ddt_exclusao) & ","
            If Trim(UCase(lfk_filial)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_filial=" & lfk_filial & " "
            
            'Executa Procedure Para Inclusão
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_GRAVA_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Operacao='INC',"
            sSQL = sSQL & "@CamposConteudos=" & oBasico.Geral.Aspas(sCamposConteudos, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & " "
            
            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SQL_GRAVA_REGISTRO!"
                Exit Function
            Else
                If cConexao.Numero_AtualizacaoBD_Lote > 0 Then
                    Atualiza_SQL_Server = Sucesso
                    sMensagem = "O comando para atualização em lote foi inserido com sucesso!"
                    Exit Function
                Else
                    If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
                        sMensagem = rsRegistrosTemp!Mensagem
                        Exit Function
                    Else
                        lpk_trabalho = rsRegistrosTemp!pk_trabalho
                        Atualiza_SQL_Server = Sucesso
                        sMensagem = "O registro foi inserido com sucesso!"
                        Exit Function
                    End If
                End If
            End If
        
        Case Alterar
            'Prepara Campos Para Alteração
            sCamposConteudos = Empty
            If Trim(UCase(lfk_trabalho_alt_seq)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_trabalho_alt_seq=" & lfk_trabalho_alt_seq & ","
            If Trim(UCase(lfk_trabalho_tipo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_trabalho_tipo=" & lfk_trabalho_tipo & ","
            If Trim(UCase(lfk_produto)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_produto=" & lfk_produto & ","
            If Trim(UCase(lfk_cliente)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_cliente=" & lfk_cliente & ","
            If Trim(UCase(lfk_representante)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_representante=" & lfk_representante & ","
            If Trim(UCase(lfk_aprovacao_tipo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_aprovacao_tipo=" & lfk_aprovacao_tipo & ","
            If Trim(UCase(snum_pedido)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "num_pedido=" & snum_pedido & ","
            If Trim(UCase(snum_pedido_antigo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "num_pedido_antigo=" & snum_pedido_antigo & ","
            If Trim(UCase(snum_pedido_novo)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "num_pedido_novo=" & snum_pedido_novo & ","
            If Trim(UCase(slead_time_programado)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "lead_time_programado=" & slead_time_programado & ","
            If Trim(UCase(ddt_inclusao)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "dt_inclusao=" & oBasico.DataHora.FormataDataGravar(ddt_inclusao) & ","
            If Trim(UCase(ddt_alteracao)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "dt_alteracao=" & oBasico.DataHora.FormataDataGravar(ddt_alteracao) & ","
            If Trim(UCase(ddt_exclusao)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "dt_exclusao=" & oBasico.DataHora.FormataDataGravar(ddt_exclusao) & ","
            If Trim(UCase(lfk_filial)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "fk_filial=" & lfk_filial & " "
            
            'Executa Procedure Para Alteração
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_GRAVA_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Operacao='ALT',"
            sSQL = sSQL & "@CamposConteudos=" & oBasico.Geral.Aspas(sCamposConteudos, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & ","
            sSQL = sSQL & "@IDRegistro=" & lpk_trabalho 

            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SQL_GRAVA_REGISTRO!"
                Exit Function
            Else
                If cConexao.Numero_AtualizacaoBD_Lote > 0 Then
                    Atualiza_SQL_Server = Sucesso
                    sMensagem = "O comando para atualização em lote foi inserido com sucesso!"
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
            'Executa Procedure Para Exclusão
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_EXCLUI_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Tipo=" & oBasico.Geral.Aspas(sTipoExclusao, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & ","
            If sTipoExclusao = "ESP" Then
                sSQL = sSQL & "@IDRegistro=" & lpk_trabalho 
            ElseIf sTipoExclusao = "CON" Then
                sSQL = sSQL & "@Condicoes=" & oBasico.Geral.Aspas(sCondicao, AspasSimples) & " "
            End If
            
            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SQL_EXCLUI_REGISTRO!"
                Exit Function
            Else
                If cConexao.Numero_AtualizacaoBD_Lote > 0 Then
                    Atualiza_SQL_Server = Sucesso
                    sMensagem = "O comando para atualização em lote foi inserido com sucesso!"
                    Exit Function
                Else
                    If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
                        If InStr(1, UCase(rsRegistrosTemp.Fields(0).Value), "SUCESSO") Then
                            Atualiza_SQL_Server = Sucesso
                            If sTipoExclusao = "ESP" Then
                                sMensagem = "O registro foi excluído com sucesso!"
                            Else
                                sMensagem = "O(s) registro(s) foi(ram) excluído(s) com sucesso!"
                            End If
                        Else
                            sMensagem = rsRegistrosTemp.Fields(0).Value
                        End If
                        Exit Function
                    Else
                        sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE DBO.PRC_SQL_EXCLUI_REGISTRO!"
                        Exit Function
                    End If
                End If
            End If
    End Select
    
    Exit Function
    
ErroRotina:
    If Err.Number = 424 Then
        sMensagem = "Não foi possível atualizar o(s) registro(s) da tabela < " & sTabelaClasse & " > ." & vbLf
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
        sMensagem = "Nenhum registro foi encontrado com os critérios indicados!"
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
