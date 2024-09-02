VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_MDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iStatus         As ADRRBR_SIS_Basico.eStatus
Private sMensagem       As String

Private Const sClasse = "SIS_MDB"

Public Property Let Status(ByVal vNewValue As ADRRBR_SIS_Basico.eStatus)
    iStatus = vNewValue
End Property
Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Function Cria(pCaminhoMDB As String, pExcluirExistente As Boolean) As Boolean
    Dim oMDB    As ADOX.Catalog
    Dim oCnxMDB As ADODB.Connection

    On Error GoTo Erro_Cria

    LimpaStatus
    iStatus = Sucesso
    
    Cria = True

    If Dir(pCaminhoMDB) <> Empty Then
        Set oCnxMDB = Nothing
        Set oCnxMDB = Abre(pCaminhoMDB)
        If iStatus <> Sucesso Then
            Cria = False
            Set oCnxMDB = Nothing
            Exit Function
        End If
        Set oCnxMDB = Nothing
        
        If pExcluirExistente Then
            Kill pCaminhoMDB
        Else
            sMensagem = "O arquivo MDB já existe e não será recriado!"
            Exit Function
        End If
    End If
    
    Set oMDB = Nothing
    Set oMDB = New ADOX.Catalog
    oMDB.Create ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pCaminhoMDB & ";")
    oMDB.ActiveConnection = Nothing
    Set oMDB = Nothing
    
    Exit Function
    
Erro_Cria:
    iStatus = Erro
    Cria = False
    Set oMDB = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Cria > "
End Function

Public Function Abre(pCaminhoMDB As String) As ADODB.Connection
    On Error GoTo Erro_Abre
    
    LimpaStatus
    iStatus = Sucesso
    
    Set Abre = Nothing
    Set Abre = New Connection
    Abre.Provider = "Microsoft.Jet.OLEDB.4.0"
    Abre.ConnectionString = "Data Source=" & pCaminhoMDB
    Abre.Open

    Exit Function
    
Erro_Abre:
    iStatus = Erro
    Set Abre = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Abre > "
End Function

Public Function CriaTabela(pNomeTabela As String, pRSTabela As ADODB.Recordset, pExcluirExistente As Boolean, pGravaRegistros As Boolean, pMantemRegistrosAtuais As Boolean, pConexaoMDB As ADODB.Connection) As Boolean
    Dim oMDB  As ADOX.Catalog
    Dim rsMDB As ADODB.Recordset
    Dim sSQL  As String
    Dim iAux  As Integer
    Dim bCria As Boolean
    
    On Error GoTo Erro_CriaTabela
    
    LimpaStatus
    iStatus = Sucesso
    
    CriaTabela = True
    
    'Tabela
    Set oMDB = Nothing
    Set oMDB = New ADOX.Catalog
    oMDB.ActiveConnection = pConexaoMDB
    
    bCria = True
    For iAux = 0 To oMDB.Tables.Count - 1
        If oMDB.Tables(iAux).Name = pNomeTabela Then
            bCria = False
            Exit For
        End If
    Next iAux
    
    If Not bCria And pExcluirExistente Then
        oMDB.Tables.Delete pNomeTabela
        bCria = True
    End If
    
    If bCria Then
        sSQL = Empty
        sSQL = sSQL & vbCrLf & Space(1) & "CREATE TABLE " & pNomeTabela & " ("
        
        'Colunas
        For iAux = 0 To pRSTabela.Fields.Count - 1
            sSQL = sSQL & vbCrLf & Space(1) & pRSTabela.Fields(iAux).Name & " " & TipoColunaMDB(pRSTabela.Fields(iAux).Type, pRSTabela.Fields(iAux).DefinedSize)
            sSQL = sSQL & IIf(iAux < pRSTabela.Fields.Count - 1, ",", ")")
        Next iAux
        
        GoSub Executa
    
        Set oMDB = Nothing
        Set oMDB = New ADOX.Catalog
        oMDB.ActiveConnection = pConexaoMDB
    End If
    
    'Indices (Opcional)
    bCria = True
    For iAux = 0 To oMDB.Tables(pNomeTabela).Indexes.Count - 1
        If oMDB.Tables(pNomeTabela).Indexes(iAux).Name = "IND_" & pNomeTabela Then
            bCria = False
            Exit For
        End If
    Next iAux
    
    If bCria Then
        sSQL = Empty
        sSQL = sSQL & vbCrLf & Space(1) & "CREATE INDEX " & "IND_" & pNomeTabela
        sSQL = sSQL & vbCrLf & Space(1) & "ON"
        sSQL = sSQL & vbCrLf & Space(1) & pNomeTabela & " (" & pRSTabela.Fields(0).Name & ")"
        
        GoSub Executa
    End If
    
    If Not pGravaRegistros Then Exit Function
    
    If Not pMantemRegistrosAtuais Then
        sSQL = Empty
        sSQL = sSQL & vbCrLf & Space(1) & "DELETE FROM " & pNomeTabela
        
        GoSub Executa
    End If
    
    If pRSTabela.EOF Then
        Set rsMDB = Nothing
        Set oMDB = Nothing
        Exit Function
    End If
    
    sSQL = Empty
    sSQL = sSQL & vbCrLf & Space(1) & "SELECT TOP 1 *"
    sSQL = sSQL & vbCrLf & Space(1) & "FROM " & pNomeTabela
    
    Set rsMDB = Nothing
    Set rsMDB = New ADODB.Recordset
    rsMDB.Open sSQL, pConexaoMDB, adOpenDynamic, adLockOptimistic

    pRSTabela.MoveFirst
    
    While Not pRSTabela.EOF
        rsMDB.AddNew
        
        For iAux = 0 To pRSTabela.Fields.Count - 1
            rsMDB.Fields(iAux).Value = pRSTabela.Fields(iAux).Value
        Next iAux
        
        rsMDB.Update
        
        pRSTabela.MoveNext
    Wend
    
    Set rsMDB = Nothing
    Set oMDB = Nothing
    
    Exit Function
    
Executa:
    pConexaoMDB.BeginTrans
        pConexaoMDB.Execute sSQL
    pConexaoMDB.CommitTrans
Return

Erro_CriaTabela:
    iStatus = Erro
    CriaTabela = False
    Set rsMDB = Nothing
    Set oMDB = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "CriaTabela > "
End Function

Private Function TipoColunaMDB(pTipoColunaRecordSet As Long, pTamanhoColunaRecordSet As Integer) As String
    '*** Número do Tipo de Coluna
    '  2 - adInteger
    '  3 - adInteger
    '  4 - adSingle
    '  5 - adDouble
    '  6 - adCurrency
    '  7 - adDate
    ' 14 - adDecimal
    ' 16 - adTinyInt
    ' 20 - adBigInt
    '129 - adChar
    '131 - adCurrency
    '133 - adDate
    '134 - adDate
    '200 - adVarChar
    '***

    Select Case pTipoColunaRecordSet
        Case 2, 3, 20
            TipoColunaMDB = "INTEGER"
            
        Case 4, 5, 6, 14, 131
            TipoColunaMDB = "CURRENCY"
            
        Case 129
            TipoColunaMDB = "CHAR (" & CStr(pTamanhoColunaRecordSet) & ")"
            
        Case 16, 200
            TipoColunaMDB = "VARCHAR (" & CStr(pTamanhoColunaRecordSet) & ")"
            
        Case 7, 133, 134
            TipoColunaMDB = "DATE"
            
        Case Else
            TipoColunaMDB = "VARCHAR (" & CStr(pTamanhoColunaRecordSet) & ")"
    End Select
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
