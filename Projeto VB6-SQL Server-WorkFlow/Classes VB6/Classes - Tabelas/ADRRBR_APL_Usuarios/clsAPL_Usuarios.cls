VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPL_Usuarios"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lPK_USUARIO As Variant
Private lINT_CODIGO As Variant
Private sVCH_NOME As Variant
Private sVCH_NOME_COMPL As Variant
Private sVCH_SENHA As Variant
Private iTINT_ATIVO As Variant

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

Private Const sTabelaClasse = "APL_Usuarios"

Private Sub Class_Initialize()
    lPK_USUARIO = "NULL"
    lINT_CODIGO = "NULL"
    sVCH_NOME = "NULL"
    sVCH_NOME_COMPL = "NULL"
    sVCH_SENHA = "NULL"
    iTINT_ATIVO = "NULL"
    
    LimpaStatus
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
    Set Estrutura_SQL = Nothing
End Sub

Public Property Let PK_USUARIO(ByVal vNewValue As Variant)
    lPK_USUARIO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get PK_USUARIO() As Variant
    PK_USUARIO = lPK_USUARIO
End Property

Public Property Let INT_CODIGO(ByVal vNewValue As Variant)
    lINT_CODIGO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get INT_CODIGO() As Variant
    INT_CODIGO = lINT_CODIGO
End Property

Public Property Let VCH_NOME(ByVal vNewValue As Variant)
    sVCH_NOME = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_NOME() As Variant
    VCH_NOME = sVCH_NOME
End Property

Public Property Let VCH_NOME_COMPL(ByVal vNewValue As Variant)
    sVCH_NOME_COMPL = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_NOME_COMPL() As Variant
    VCH_NOME_COMPL = sVCH_NOME_COMPL
End Property

Public Property Let VCH_SENHA(ByVal vNewValue As Variant)
    sVCH_SENHA = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_SENHA() As Variant
    VCH_SENHA = sVCH_SENHA
End Property

Public Property Let TINT_ATIVO(ByVal vNewValue As Variant)
    iTINT_ATIVO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TINT_ATIVO() As Variant
    TINT_ATIVO = iTINT_ATIVO
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
        If Trim(UCase(lPK_USUARIO)) <> "NULL" And Trim(lPK_USUARIO) <> Empty And Trim(lPK_USUARIO) <> "0" Then
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
        
        .Coluna = "PK_USUARIO"
        .Conteudo = lPK_USUARIO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "INT_CODIGO"
        .Conteudo = lINT_CODIGO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_NOME"
        .Conteudo = sVCH_NOME
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_NOME_COMPL"
        .Conteudo = sVCH_NOME_COMPL
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_SENHA"
        .Conteudo = sVCH_SENHA
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "TINT_ATIVO"
        .Conteudo = iTINT_ATIVO
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
            If Trim(UCase(lINT_CODIGO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "INT_CODIGO=" & lINT_CODIGO & ","
            If Trim(UCase(sVCH_NOME)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME=" & sVCH_NOME & ","
            If Trim(UCase(sVCH_NOME_COMPL)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_COMPL=" & sVCH_NOME_COMPL & ","
            If Trim(UCase(sVCH_SENHA)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_SENHA=" & sVCH_SENHA & ","
            If Trim(UCase(iTINT_ATIVO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "TINT_ATIVO=" & iTINT_ATIVO & " "
            
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
                        lPK_USUARIO = rsRegistrosTemp!PK_USUARIO
                        Atualiza_SQL_Server = Sucesso
                        sMensagem = "O registro foi inserido com sucesso!"
                        Exit Function
                    End If
                End If
            End If
        
        Case Alterar
            'Prepara Campos Para Alteração
            sCamposConteudos = Empty
            If Trim(UCase(lINT_CODIGO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "INT_CODIGO=" & lINT_CODIGO & ","
            If Trim(UCase(sVCH_NOME)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME=" & sVCH_NOME & ","
            If Trim(UCase(sVCH_NOME_COMPL)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_NOME_COMPL=" & sVCH_NOME_COMPL & ","
            If Trim(UCase(sVCH_SENHA)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_SENHA=" & sVCH_SENHA & ","
            If Trim(UCase(iTINT_ATIVO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "TINT_ATIVO=" & iTINT_ATIVO & " "
            
            'Executa Procedure Para Alteração
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_GRAVA_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Operacao='ALT',"
            sSQL = sSQL & "@CamposConteudos=" & oBasico.Geral.Aspas(sCamposConteudos, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & ","
            sSQL = sSQL & "@IDRegistro=" & lPK_USUARIO 

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
                sSQL = sSQL & "@IDRegistro=" & lPK_USUARIO 
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
