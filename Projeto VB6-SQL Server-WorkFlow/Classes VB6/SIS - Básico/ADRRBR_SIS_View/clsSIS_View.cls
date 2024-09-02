VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_View"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sViewConsulta     As String
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

Const sObjetoClasse = "View"

Private Sub Class_Initialize()
    LimpaStatus
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
    Set Estrutura_SQL = Nothing
End Sub

Public Property Let ViewConsulta(ByVal vNewValue As String)
    sViewConsulta = vNewValue
End Property
Public Property Get ViewConsulta() As String
    ViewConsulta = sViewConsulta
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

Public Sub Consultar_BD()
    On Error GoTo ErroRotina

    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    If Not VerificaPropriedadesConsulta Then Exit Sub

    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            iStatus = Consulta_SQL_Server
    End Select
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sObjetoClasse & "." & "Consultar_BD > "
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sObjetoClasse & "." & "VerificaPropriedadesGerais > "
End Function

Private Function VerificaPropriedadesConsulta() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesConsulta = True
    
    If sViewConsulta = Empty Then
        iStatus = Erro
        sMensagem = "Informe a VIEW para consulta dos dados!"
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sObjetoClasse & "." & "VerificaPropriedadesConsulta > "
End Function

Private Function Consulta_SQL_Server() As eStatus
    On Error GoTo ErroRotina
    
    Consulta_SQL_Server = Erro
    
    sSQL = Empty
    sSQL = sSQL & "DBO.PRC_SQL_PESQUISA_REGISTROS "
    sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sViewConsulta, AspasSimples) & ", "
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sObjetoClasse & "." & "Consulta_SQL_Server > "
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sObjetoClasse & "." & "LimpaStatus > "
End Sub
