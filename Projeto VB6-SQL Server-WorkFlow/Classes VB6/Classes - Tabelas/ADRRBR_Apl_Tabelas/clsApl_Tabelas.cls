VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApl_Tabelas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lPK_TABELA As Variant
Private lFK_SISTEMA As Variant
Private sVCH_TABELA As Variant
Private sTXT_DESCRICAO As Variant
Private iTINT_GRAVA_MANUT As Variant

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

Private Const sTabelaClasse = "Apl_Tabelas"

Private Sub Class_Initialize()
    lPK_TABELA = "NULL"
    lFK_SISTEMA = "NULL"
    sVCH_TABELA = "NULL"
    sTXT_DESCRICAO = "NULL"
    iTINT_GRAVA_MANUT = "NULL"
    
    LimpaStatus
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
    Set Estrutura_SQL = Nothing
End Sub

Public Property Let PK_TABELA(ByVal vNewValue As Variant)
    lPK_TABELA = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get PK_TABELA() As Variant
    PK_TABELA = lPK_TABELA
End Property

Public Property Let FK_SISTEMA(ByVal vNewValue As Variant)
    lFK_SISTEMA = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FK_SISTEMA() As Variant
    FK_SISTEMA = lFK_SISTEMA
End Property

Public Property Let VCH_TABELA(ByVal vNewValue As Variant)
    sVCH_TABELA = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get VCH_TABELA() As Variant
    VCH_TABELA = sVCH_TABELA
End Property

Public Property Let TXT_DESCRICAO(ByVal vNewValue As Variant)
    sTXT_DESCRICAO = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TXT_DESCRICAO() As Variant
    TXT_DESCRICAO = sTXT_DESCRICAO
End Property

Public Property Let TINT_GRAVA_MANUT(ByVal vNewValue As Variant)
    iTINT_GRAVA_MANUT = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TINT_GRAVA_MANUT() As Variant
    TINT_GRAVA_MANUT = iTINT_GRAVA_MANUT
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
        If Trim(UCase(lPK_TABELA)) <> "NULL" And Trim(lPK_TABELA) <> Empty And Trim(lPK_TABELA) <> "0" Then
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
        
        .Coluna = "PK_TABELA"
        .Conteudo = lPK_TABELA
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "FK_SISTEMA"
        .Conteudo = lFK_SISTEMA
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "VCH_TABELA"
        .Conteudo = sVCH_TABELA
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "TXT_DESCRICAO"
        .Conteudo = sTXT_DESCRICAO
        .ValidaConteudoColuna
        Gosub VerificaStatus

        .Coluna = "TINT_GRAVA_MANUT"
        .Conteudo = iTINT_GRAVA_MANUT
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
            If Trim(UCase(lFK_SISTEMA)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "FK_SISTEMA=" & lFK_SISTEMA & ","
            If Trim(UCase(sVCH_TABELA)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_TABELA=" & sVCH_TABELA & ","
            If Trim(UCase(sTXT_DESCRICAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "TXT_DESCRICAO=" & sTXT_DESCRICAO & ","
            If Trim(UCase(iTINT_GRAVA_MANUT)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "TINT_GRAVA_MANUT=" & iTINT_GRAVA_MANUT & " "
            
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
                        lPK_TABELA = rsRegistrosTemp!PK_TABELA
                        Atualiza_SQL_Server = Sucesso
                        sMensagem = "O registro foi inserido com sucesso!"
                        Exit Function
                    End If
                End If
            End If
        
        Case Alterar
            'Prepara Campos Para Altera��o
            sCamposConteudos = Empty
            If Trim(UCase(lFK_SISTEMA)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "FK_SISTEMA=" & lFK_SISTEMA & ","
            If Trim(UCase(sVCH_TABELA)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "VCH_TABELA=" & sVCH_TABELA & ","
            If Trim(UCase(sTXT_DESCRICAO)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "TXT_DESCRICAO=" & sTXT_DESCRICAO & ","
            If Trim(UCase(iTINT_GRAVA_MANUT)) <> "NULL" Then sCamposConteudos = sCamposConteudos & "TINT_GRAVA_MANUT=" & iTINT_GRAVA_MANUT & " "
            
            'Executa Procedure Para Altera��o
            sSQL = Empty
            sSQL = sSQL & "DBO.PRC_SQL_GRAVA_REGISTRO "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sTabelaClasse, AspasSimples) & ","
            sSQL = sSQL & "@Operacao='ALT',"
            sSQL = sSQL & "@CamposConteudos=" & oBasico.Geral.Aspas(sCamposConteudos, AspasSimples) & ","
            sSQL = sSQL & "@NumAtuBD_Lote=" & cConexao.Numero_AtualizacaoBD_Lote & ","
            sSQL = sSQL & "@IDRegistro=" & lPK_TABELA 

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
                sSQL = sSQL & "@IDRegistro=" & lPK_TABELA 
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