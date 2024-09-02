VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Word"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private clTextosSubstituir As Collection
Private iStatus            As ADRRBR_SIS_Basico.eStatus
Private sMensagem          As String

Dim WordApp                As Word.Application
Dim oWord                  As Word.Selection

Private Const sClasse = "SIS_WORD"

Private Sub Class_Initialize()
    Set clTextosSubstituir = New Collection
End Sub

Private Sub Class_Terminate()
    If Not WordApp Is Nothing Then WordApp.Quit
    Set WordApp = Nothing
    Set oWord = Nothing
    Set clTextosSubstituir = Nothing
End Sub

Public Property Get TextosSubstituir() As Collection
    Set TextosSubstituir = clTextosSubstituir
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

Public Function Gera(pCaminhoArquivoModelo As String, pCaminhoArquivoGerar As String, pExcluirExistente As Boolean, pAbreWord As Boolean) As Boolean
    Dim iTexto   As Integer
    Dim sTexto() As String
    
    On Error GoTo Erro_Gera

    LimpaStatus
    
    Gera = True
    iStatus = Sucesso

    Set WordApp = New Word.Application
    
    If Dir(pCaminhoArquivoModelo) = Empty Then
        Gera = False
        iStatus = Erro
        sMensagem = "O caminho/arquivo modelo informado, não existe!"
        Exit Function
    End If
    
    If Dir(pCaminhoArquivoGerar) <> Empty Then
        If Not pExcluirExistente Then
            Gera = False
            iStatus = Erro
            sMensagem = "O caminho/arquivo gerar informado, já existe!"
            Exit Function
        End If
        Kill pCaminhoArquivoGerar
    End If
    
    FileCopy pCaminhoArquivoModelo, pCaminhoArquivoGerar
    WordApp.Documents.Open pCaminhoArquivoGerar
    
    If clTextosSubstituir.Count > 0 Then
        For iTexto = 1 To clTextosSubstituir.Count
            If InStr(1, clTextosSubstituir(iTexto), ";") > 0 Then
                sTexto() = Split(clTextosSubstituir(iTexto), ";")
                
                Set oWord = WordApp.Selection
                oWord.Find.Forward = True
                oWord.Find.ClearFormatting
                oWord.Find.MatchWholeWord = True
                oWord.Find.MatchCase = False
                oWord.Find.Wrap = wdFindContinue
                
                oWord.Find.Execute FindText:=sTexto(0)
                Clipboard.SetText sTexto(1)
                oWord.Text = Clipboard.GetText
            End If
        Next iTexto
    End If
    
    WordApp.Visible = pAbreWord

    Exit Function

Erro_Gera:
    iStatus = Erro
    Gera = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Gera > "
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


