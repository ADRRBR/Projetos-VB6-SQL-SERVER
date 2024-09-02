VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_BD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iTipoBancoDados As ADRRBR_SIS_Basico.eTipoBancoDados
Private iAcao           As ADRRBR_SIS_Basico.eAcao
Private iStatus         As ADRRBR_SIS_Basico.eStatus
Private sMensagem       As String

Enum eTipoBancoDados
    SQL_Server = 1
    Access = 2
End Enum

Enum eAcao
    Incluir = 1
    Alterar = 2
    Excluir = 3
    Consultar = 4
End Enum

Enum eStatus
    SemRequisicao = 0
    Encontrado = 1
    NaoEncontrado = 2
    Sucesso = 3
    Erro = 4
    Confirmado = 5
    Cancelado = 6
End Enum

Private Const sClasse = "SIS_BD"

Public Property Let TipoBancoDados(ByVal vNewValue As ADRRBR_SIS_Basico.eTipoBancoDados)
    iTipoBancoDados = vNewValue
End Property
Public Property Get TipoBancoDados() As ADRRBR_SIS_Basico.eTipoBancoDados
    TipoBancoDados = iTipoBancoDados
End Property

Public Property Let Acao(ByVal vNewValue As ADRRBR_SIS_Basico.eAcao)
    iAcao = vNewValue
End Property
Public Property Get Acao() As ADRRBR_SIS_Basico.eAcao
    Acao = iAcao
End Property

Public Property Let Status(ByVal vNewValue As ADRRBR_SIS_Basico.eStatus)
    iStatus = vNewValue
End Property
Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Function ConexaoSeguranca() As Object
    On Error GoTo ErroConexaoSeguranca
    
    LimpaStatus
    
    If Not VerificaTipoBancoDados Then Exit Function

    Set ConexaoSeguranca = Nothing
    Set ConexaoSeguranca = CreateObject("ADRRBR_SIS_Conexao.clsSIS_Conexao")
    
    ConexaoSeguranca.TipoBancoDados = iTipoBancoDados
    ConexaoSeguranca.RecuperaParametrosConexao
    ConexaoSeguranca.Conecta
    
    If ConexaoSeguranca.Conectado Then Exit Function
    
    While Not ConexaoSeguranca.Conectado And ConexaoSeguranca.Status <> Cancelado
        ConexaoSeguranca.LoginSeguranca
        MsgBox ConexaoSeguranca.Mensagem
        If ConexaoSeguranca.Status = Sucesso Then ConexaoSeguranca.Conecta
    Wend
    
    Exit Function
    
ErroConexaoSeguranca:
    iStatus = Erro
    Set ConexaoSeguranca = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "Não foi possível efetuar a conexão de segurança!"
End Function

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

