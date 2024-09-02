VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iStatus            As ADRRBR_SIS_Basico.eStatus
Private sMensagem          As String
Private oPlanilha          As Excel.Worksheet

Dim oExcel                 As Excel.Application
Dim oBook                  As Excel.Workbook

Dim sCaminhoArquivoGerar   As String

Private Const sClasse = "SIS_EXCEL"

Private Sub Class_Terminate()
    If Not oExcel Is Nothing Then oExcel.Quit
    Set oPlanilha = Nothing
    Set oBook = Nothing
    Set oExcel = Nothing
End Sub

Public Property Get Planilha() As Excel.Worksheet
    Set Planilha = oPlanilha
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

Public Function Inicializa(pCaminhoArquivoModelo As String, pCaminhoArquivoGerar As String, pExcluirExistente As Boolean)
    On Error GoTo Erro_Inicializa

    LimpaStatus
    
    Inicializa = True
    iStatus = Sucesso
    
    Set oExcel = Nothing
    Set oExcel = New Excel.Application

    If Dir(pCaminhoArquivoModelo) = Empty Then
        Inicializa = False
        iStatus = Erro
        sMensagem = "O caminho/arquivo modelo informado, não existe!"
        Exit Function
    End If
    
    If Dir(pCaminhoArquivoGerar) <> Empty Then
        If Not pExcluirExistente Then
            Inicializa = False
            iStatus = Erro
            sMensagem = "O caminho/arquivo gerar informado, já existe!"
            Exit Function
        End If
        Kill pCaminhoArquivoGerar
    End If

    sCaminhoArquivoGerar = pCaminhoArquivoGerar
    
    Set oBook = Nothing
    Set oBook = oExcel.Workbooks.Open(pCaminhoArquivoModelo, False)
    
    Set oPlanilha = Nothing
    Set oPlanilha = oBook.Worksheets(1)
    
    Exit Function

Erro_Inicializa:
    iStatus = Erro
    Inicializa = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Inicializa > "
End Function

Public Function Aplica(pAbreExcel As Boolean) As Boolean
    On Error GoTo Erro_Aplica

    LimpaStatus
    
    Aplica = True
    iStatus = Sucesso

    If oExcel Is Nothing Or oBook Is Nothing Or oPlanilha Is Nothing Then
        Aplica = False
        iStatus = Erro
        sMensagem = "O objeto excel ainda não foi inicializado!"
        Exit Function
    End If

    oBook.SaveAs sCaminhoArquivoGerar
    oBook.Close False
    
    If pAbreExcel Then
        Set oBook = Nothing
        Set oBook = oExcel.Workbooks.Open(sCaminhoArquivoGerar, False)
        oExcel.Visible = True
    End If
    
    Exit Function
    
Erro_Aplica:
    iStatus = Erro
    Aplica = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Aplica > "
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

Public Function GeracaoAutomatica(pRSExcel As ADODB.Recordset) As Boolean
    Dim oExcel  As Excel.Application
    Dim oBook   As Excel.Workbook
    Dim lLinha  As Long
    Dim iColuna As Integer
    Dim lLinhas As Long
    
    On Error GoTo Erro_GeracaoAutomatica
    
    LimpaStatus
    
    GeracaoAutomatica = True

    Set oExcel = Nothing
    Set oExcel = New Excel.Application
    
    Set oBook = Nothing
    Set oBook = oExcel.Workbooks.Add
    
    pRSExcel.MoveFirst
    
    While Not pRSExcel.EOF
        lLinhas = lLinhas + 1
        pRSExcel.MoveNext
    Wend
    
    For iColuna = 1 To pRSExcel.Fields.Count
        oBook.ActiveSheet.Cells(1, iColuna).Value = pRSExcel.Fields(iColuna - 1).Name
    Next iColuna
    
    pRSExcel.MoveFirst
    
    For lLinha = 1 To lLinhas
        For iColuna = 0 To pRSExcel.Fields.Count - 1
            oBook.ActiveSheet.Cells(lLinha + 1, iColuna + 1).Value = pRSExcel.Fields(iColuna).Value
        Next iColuna
        
        pRSExcel.MoveNext
    Next lLinha
    
    oExcel.Cells.Select
    oExcel.Cells.EntireColumn.AutoFit
    oExcel.Rows("1:1").Select
    oExcel.Selection.Font.Bold = True
    
    oExcel.Visible = True
    Set oBook = Nothing
    Set oExcel = Nothing

    Exit Function
    
Erro_GeracaoAutomatica:
    iStatus = Erro
    GeracaoAutomatica = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "GeracaoAutomatica > "
End Function

