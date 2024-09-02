VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Pesquisa"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private bView             As Boolean
Private sTituloPesquisa   As String
Private sTabelaPesquisar  As String
Private sObjetoPesquisar  As String
Private sColunasExibir    As String
Private sCondicao         As String
Private bSelecionado      As Boolean
Private rsRegistroSelecao As Object
Private cConexao          As Object
Private iStatus           As ADRRBR_SIS_Basico.eStatus
Private sMensagem         As String

Private oBasico           As New ADRRBR_SIS_Basico.clsSIS_Basico

Private Const sNomeClasse = "SIS_Pesquisa"

Private Sub Class_Initialize()
    LimpaStatus
    bSelecionado = False
End Sub

Private Sub Class_Terminate()
    Set rsRegistroSelecao = Nothing
    Set cConexao = Nothing
    Set oBasico = Nothing
End Sub

Public Property Let View(ByVal vNewValue As Boolean)
    bView = vNewValue
End Property
Public Property Get View() As Boolean
    View = bView
End Property

Public Property Let TituloPesquisa(ByVal vNewValue As String)
    sTituloPesquisa = vNewValue
End Property
Public Property Get TituloPesquisa() As String
    TituloPesquisa = sTituloPesquisa
End Property

Public Property Let TabelaPesquisar(ByVal vNewValue As String)
    sTabelaPesquisar = vNewValue
End Property
Public Property Get TabelaPesquisar() As String
    TabelaPesquisar = sTabelaPesquisar
End Property

Public Property Let ObjetoPesquisar(ByVal vNewValue As String)
    sObjetoPesquisar = vNewValue
End Property
Public Property Get ObjetoPesquisar() As String
    ObjetoPesquisar = sObjetoPesquisar
End Property

Public Property Let ColunasExibir(ByVal vNewValue As String)
    sColunasExibir = vNewValue
End Property
Public Property Get ColunasExibir() As String
    ColunasExibir = sColunasExibir
End Property

Public Property Let Condicao(ByVal vNewValue As String)
    sCondicao = vNewValue
End Property
Public Property Get Condicao() As String
    Condicao = sCondicao
End Property

Public Property Get Selecionado() As Boolean
    Selecionado = bSelecionado
End Property

Public Property Get RegistroSelecao() As Object
    Set RegistroSelecao = rsRegistroSelecao
End Property

Public Property Set Conexao(ByVal vNewValue As Object)
    Set cConexao = vNewValue
End Property
Public Property Get Conexao() As Object
    Set Conexao = cConexao
End Property

Public Property Get Status() As eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Sub Pesquisar()
    On Error GoTo ErroRotina

    LimpaStatus
    
    bSelecionado = False
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    
    Set frmPesquisa.Conexao = cConexao
    frmPesquisa.View = bView
    frmPesquisa.TituloPesquisa = sTituloPesquisa
    frmPesquisa.TabelaPesquisar = sTabelaPesquisar
    frmPesquisa.ObjetoPesquisar = sObjetoPesquisar
    frmPesquisa.ColunasExibir = sColunasExibir
    frmPesquisa.Condicao = sCondicao
    frmPesquisa.Show vbModal
    
    iStatus = frmPesquisa.Status
    sMensagem = frmPesquisa.Mensagem
    
    If iStatus <> Encontrado Then
        Set frmPesquisa = Nothing
        Exit Sub
    End If
    
    Set rsRegistroSelecao = frmPesquisa.RegistroSelecao
    Set frmPesquisa = Nothing
    
    iStatus = Sucesso
    sMensagem = Empty
    
    bSelecionado = True
        
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sNomeClasse & "." & "Pesquisar > "
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
    
    If Not cConexao.Conectado Then
        iStatus = Erro
        sMensagem = "O objeto de conexão informado NÃO está CONECTADO!"
        VerificaPropriedadesGerais = False
        Exit Function
    End If
    
    If Trim(sTabelaPesquisar) = Empty Then
        iStatus = Erro
        sMensagem = "Informe obrigatoriamente a tabela a ser pesquisada!"
        VerificaPropriedadesGerais = False
        Exit Function
    End If
    
    If Trim(sObjetoPesquisar) = Empty Then sObjetoPesquisar = sTabelaPesquisar
    
    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaPropriedadesGerais = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sNomeClasse & "." & "VerificaPropriedadesGerais > "
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sNomeClasse & "." & "LimpaStatus > "
End Sub
