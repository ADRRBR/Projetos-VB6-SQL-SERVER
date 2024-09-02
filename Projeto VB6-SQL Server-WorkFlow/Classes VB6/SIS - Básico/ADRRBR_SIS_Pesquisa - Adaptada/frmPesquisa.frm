VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPesquisa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa..."
   ClientHeight    =   6300
   ClientLeft      =   3390
   ClientTop       =   2265
   ClientWidth     =   8040
   Icon            =   "frmPesquisa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboOrdem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   3105
   End
   Begin VB.TextBox txtPesquisa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   660
      Width           =   7845
   End
   Begin MSComctlLib.ListView lvwPesquisa 
      Height          =   4650
      Left            =   90
      TabIndex        =   2
      Top             =   1305
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   8202
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList3"
      SmallIcons      =   "ImageList3"
      ForeColor       =   -2147483635
      BackColor       =   15133154
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraMensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   2745
      TabIndex        =   4
      Top             =   2400
      Width           =   2475
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D2F4FD&
         Height          =   225
         Left            =   795
         TabIndex        =   5
         Top             =   135
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00D2F4FD&
         Height          =   465
         Left            =   30
         Top             =   30
         Width           =   2415
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Registros da Tabela <......>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   90
      TabIndex        =   7
      Tag             =   "SELECAO"
      Top             =   1035
      Width           =   7860
   End
   Begin VB.Label lblQtdRegistros 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "lblQtdReg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   210
      Left            =   6990
      TabIndex        =   3
      Tag             =   "LISTA"
      Top             =   6020
      Width           =   915
   End
   Begin VB.Label lblInforma 
      Caption         =   "Seleção/Ordem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   105
      TabIndex        =   6
      Top             =   75
      Width           =   1260
   End
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cConexao          As Object
Private iStatus           As ADRRBR_SIS_Basico.eStatus
Private sMensagem         As String
Private bView             As Boolean
Private sTituloPesquisa   As String
Private sTabelaPesquisar  As String
Private sObjetoPesquisar  As String
Private sColunasExibir    As String
Private sCondicao         As String
Private rsRegistroSelecao As Object

Dim Estrutura_SQL         As Object
Dim rsColunasExibir       As ADODB.Recordset
Dim rsPesquisa            As ADODB.Recordset
 
Const sNomePadrao = "Pesquisa"
 
Public Property Set Conexao(ByVal vNewValue As Object)
    Set cConexao = vNewValue
End Property
Public Property Get Conexao() As Object
    Set Conexao = cConexao
End Property

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
 
Public Property Get RegistroSelecao() As Object
    Set RegistroSelecao = rsRegistroSelecao
End Property
 
Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property
 
Private Sub Form_Load()
    Select Case cConexao.TipoBancoDados
        Case SQL_Server:  Me.Caption = sNomePadrao & Space(8) & "(" & cConexao.Servidor & "." & cConexao.BancoDados & ")"
        Case Access: Me.Caption = sNomePadrao & Space(8) & "(" & cConexao.CaminhoMDB & ")"
    End Select

    If Trim(sTituloPesquisa) = Empty Then
        lblTitulo.Caption = "Registros da tabela  < " & sTabelaPesquisar & " >"
    Else
        lblTitulo.Caption = sTituloPesquisa
    End If
    
    InicializaTela
End Sub

Private Sub Form_Activate()
    If rsPesquisa Is Nothing Then
        Retorna
        Exit Sub
    End If

    oBasico.Geral.PosicionaTela Screen, Me
    
    CriaRecordSetColunasExibir_SQL_Server
    CriaColunasListView_SQL_Server
    EncheListaPesquisa_SQL_Server
    
    cboOrdem.ListIndex = 0
End Sub

Private Sub cboOrdem_Click()
    txtPesquisa.Text = Empty

    'Reordena a Lista
    With lvwPesquisa
       .Sorted = True
       .SortOrder = lvwAscending
       .SortKey = cboOrdem.ListIndex
       .Sorted = False
    End With
    
    txtPesquisa.SetFocus
End Sub

Private Sub lvwPesquisa_DblClick()
    'Retorna ao Formulário de Chamada Com o Registro Selecionado.
    
    If rsPesquisa Is Nothing Then Exit Sub
    
    LimpaStatus
    
    rsPesquisa.Filter = ""
    rsPesquisa.MoveFirst
    rsPesquisa.Filter = CStr(rsPesquisa.Fields(0).Name) & " = " & lvwPesquisa.ListItems(lvwPesquisa.SelectedItem.Index).Text
    
    Set rsRegistroSelecao = rsPesquisa
    
    If rsRegistroSelecao.EOF Then
        iStatus = NaoEncontrado
        sMensagem = "Registro selecionado, não localizado!"
    Else
        iStatus = Encontrado
        sMensagem = "Registro selecionado!"
    End If
    
    Retorna
End Sub

Private Sub lvwPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lvwPesquisa_DblClick
End Sub

Private Sub lvwPesquisa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwPesquisa
       .Sorted = True
       If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
       Else
          .SortOrder = lvwAscending
       End If
       .SortKey = ColumnHeader.Index - 1
       .Sorted = False
    End With
End Sub

Private Sub txtPesquisa_Change()
    Dim sColuna            As String
    Dim sCondicoesPesquisa As String
    
    If Trim(txtPesquisa.Text) = Empty Then
        Set rsPesquisa = Pesquisa(sTabelaPesquisar, sObjetoPesquisar, sCondicao)
        
        Select Case cConexao.TipoBancoDados
            Case SQL_Server
                EncheListaPesquisa_SQL_Server
            
            Case Access
            
        End Select
        
        Exit Sub
    End If
    
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            sColuna = Estrutura_SQL.Estrutura!Coluna
            
            If Not VerificaTipoColuna_SQL_Server(cboOrdem.List(cboOrdem.ListIndex)) Then Exit Sub
                
            Select Case Estrutura_SQL.Estrutura!Tipo_Referencia
                Case "TEXTO"
                    sCondicoesPesquisa = sColuna & " LIKE ^" & txtPesquisa.Text & "%^"
                    
                Case "NUMERO"
                    sCondicoesPesquisa = sColuna & " = " & txtPesquisa.Text & " "
                
                Case "DATA"
                    sCondicoesPesquisa = sColuna & " LIKE ^" & txtPesquisa.Text & "%^"
                
                Case "VALOR"
                    sCondicoesPesquisa = sColuna & " LIKE ^" & txtPesquisa.Text & "%^"
            End Select
                
            If Trim(sCondicao) <> Empty Then sCondicoesPesquisa = sCondicoesPesquisa & " AND " & sCondicao
                
            Set rsPesquisa = Pesquisa(sTabelaPesquisar, sObjetoPesquisar, sCondicoesPesquisa)
            
            EncheListaPesquisa_SQL_Server
        
        Case Access
        
    End Select
End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}": Exit Sub
        
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            If Not VerificaTipoColuna_SQL_Server(cboOrdem.List(cboOrdem.ListIndex)) Then Exit Sub
        
            Select Case Estrutura_SQL.Estrutura!Tipo_Referencia
                Case "NUMERO"
                    KeyAscii = oBasico.Geral.TrataNumeros(KeyAscii)
                Case Else
                    KeyAscii = oBasico.Geral.Maiuscula(KeyAscii)
            End Select
        
        Case Access
        
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Retorna ao Formulário de Chamada sem Nenhum Registro Selecionado.
    If KeyCode = vbKeyEscape Then
        Retorna
        Exit Sub
    End If
End Sub

Sub InicializaTela()
    LimpaStatus
    
    MousePointer = vbHourglass
    
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            If CarregaEstruturaTabela_SQL_Server Then Set rsPesquisa = Pesquisa(sTabelaPesquisar, sObjetoPesquisar, sCondicao)
            
        Case Access
        
    End Select
    
    MousePointer = vbDefault
End Sub

Private Function CarregaEstruturaTabela_SQL_Server() As Boolean
    On Error GoTo ErroRotina
    
    CarregaEstruturaTabela_SQL_Server = False
    
    Set Estrutura_SQL = CreateObject("ADRRBR_SQL_Estrutura.clsSQL_Estrutura")
    
    Set Estrutura_SQL.Conexao = cConexao
    Estrutura_SQL.Tabela = sTabelaPesquisar
    Estrutura_SQL.CarregaEstruturaTabela
    
    iStatus = Estrutura_SQL.Status
    sMensagem = Estrutura_SQL.Mensagem
    
    If iStatus = Erro Then
        CarregaEstruturaTabela_SQL_Server = False
        Exit Function
    End If
    
    CarregaEstruturaTabela_SQL_Server = True
    
    Exit Function
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A FUNÇÃO PRIVADA < " & sNomePadrao & "." & "CarregaEstruturaTabela_SQL_Server > "
End Function

Sub CriaColunasListView_SQL_Server()
    Dim sColuna        As String
    Dim lTamanhoColuna As Long
    Dim iAlinhamento   As ADRRBR_SIS_Basico.eAlinhamentoLV
    
    cboOrdem.Clear
    
    Estrutura_SQL.Estrutura.Filter = ""
    Estrutura_SQL.Estrutura.MoveFirst
    
    While Not Estrutura_SQL.Estrutura.EOF
        If ColunaInformada(Estrutura_SQL.Estrutura!Coluna) Then
            sColuna = LocalizaApelidoColuna(Estrutura_SQL.Estrutura!Coluna)
            
            Select Case Estrutura_SQL.Estrutura!Tipo_Referencia
                Case "TEXTO"
                    If Estrutura_SQL.Estrutura!Tamanho <= 100 Then
                        lTamanhoColuna = 3500
                    
                    ElseIf Estrutura_SQL.Estrutura!Tamanho >= 100 Then
                        lTamanhoColuna = 5000
                    Else
                        lTamanhoColuna = 6500
                    End If
                    iAlinhamento = Esquerda_LV
                    
                Case "NUMERO"
                    If Estrutura_SQL.Estrutura!Tamanho <= 4 Then
                        lTamanhoColuna = 1400
                    Else
                        lTamanhoColuna = 2200
                    End If
                    iAlinhamento = Esquerda_LV
                
                Case "DATA"
                    lTamanhoColuna = 2000
                    iAlinhamento = Centro_LV
                
                Case "VALOR"
                    lTamanhoColuna = 2200
                    iAlinhamento = Direita_LV
            End Select
            
            If LocalizaApelidoColuna(Estrutura_SQL.Estrutura!Coluna) = "Chave Primária" Then
                lTamanhoColuna = 0
            Else
                'Adiciona Campos Para Possibilitar Ordenação
                cboOrdem.AddItem sColuna
            End If
            
            oBasico.LV.AddCol lvwPesquisa, sColuna, lTamanhoColuna, iAlinhamento
        End If
    
        Estrutura_SQL.Estrutura.MoveNext
    Wend
End Sub

Function VerificaTipoColuna_SQL_Server(pColuna As String) As Boolean
    Dim sColuna As String
    
    VerificaTipoColuna_SQL_Server = True
    
    sColuna = LocalizaColuna(cboOrdem.List(cboOrdem.ListIndex))
    
    Estrutura_SQL.Estrutura.Filter = ""
    Estrutura_SQL.Estrutura.MoveFirst
    Estrutura_SQL.Estrutura.Filter = "COLUNA='" & sColuna & "'"
    
    If Estrutura_SQL.Estrutura.EOF Then
        VerificaTipoColuna_SQL_Server = False
        MsgBox "A coluna " & pColuna & " não existe na estrutura!", vbCritical, "Atenção"
        Exit Function
    End If
End Function

Sub EncheListaPesquisa_SQL_Server()
    Dim iCampo      As Integer
    Dim iCampoLista As Integer
    Dim sConteudo   As String
    
    lvwPesquisa.ListItems.Clear
    
    If rsPesquisa Is Nothing Then GoTo ExibicaoTotalRegistros
   
    If Not rsPesquisa.EOF Then rsPesquisa.MoveFirst
    
    While Not rsPesquisa.EOF
        lvwPesquisa.ListItems.Add
        
        iCampo = 0
        iCampoLista = 0
        Estrutura_SQL.Estrutura.Filter = ""
        Estrutura_SQL.Estrutura.MoveFirst
        
        While Not Estrutura_SQL.Estrutura.EOF
            If ColunaInformada(Estrutura_SQL.Estrutura!Coluna) Then
                Select Case Estrutura_SQL.Estrutura!Tipo_Referencia
                    Case "TEXTO"
                        sConteudo = oBasico.Geral.TrocaNuLL(rsPesquisa.Fields(iCampo).Value, Empty)
                        GoSub AtualizaConteudoColuna
                            
                    Case "NUMERO"
                        sConteudo = oBasico.Geral.TrocaNuLL(rsPesquisa.Fields(iCampo).Value, 0)
                        GoSub AtualizaConteudoColuna
                    
                    Case "DATA"
                        sConteudo = oBasico.Geral.TrocaNuLL(rsPesquisa.Fields(iCampo).Value, Empty)
                        GoSub AtualizaConteudoColuna
                    
                    Case "VALOR"
                        sConteudo = Format(oBasico.Geral.TrocaNuLL(rsPesquisa.Fields(iCampo).Value, 0), "###,###,###,##0.00")
                        GoSub AtualizaConteudoColuna
                End Select
            
                iCampoLista = iCampoLista + 1
            End If
            
            iCampo = iCampo + 1
            
            Estrutura_SQL.Estrutura.MoveNext
        Wend
                        
        rsPesquisa.MoveNext
    Wend
    
    GoTo ExibicaoTotalRegistros
    
    Exit Sub
    
AtualizaConteudoColuna:
    If iCampoLista = 0 Then
        lvwPesquisa.ListItems(lvwPesquisa.ListItems.Count).Text = sConteudo
    Else
        lvwPesquisa.ListItems(lvwPesquisa.ListItems.Count).SubItems(iCampoLista) = sConteudo
    End If
Return

ExibicaoTotalRegistros:
    If lvwPesquisa.ListItems.Count = 1 Then
        lblQtdRegistros.Caption = lvwPesquisa.ListItems.Count & " registro"
    Else
        lblQtdRegistros.Caption = lvwPesquisa.ListItems.Count & " registros"
    End If
End Sub

Private Function Pesquisa(pTabelaPesquisar As String, pObjetoPesquisar As String, Optional pCondicoesPesquisa As String) As ADODB.Recordset
    Dim oTabela As Object
    
    On Error GoTo ErroRotina
    
    Set oTabela = Nothing
    
    If bView Then
        Set oTabela = CreateObject("ADRRBR_SIS_View.clsSIS_View")
        oTabela.ViewConsulta = pTabelaPesquisar
    Else
        Set oTabela = CreateObject("ADRRBR_" & pObjetoPesquisar & "." & "cls" & pObjetoPesquisar)
        oTabela.Acao = Consultar
    End If
    
    Set oTabela.Conexao = cConexao
    
    oTabela.Condicao = Replace(pCondicoesPesquisa, "'", "^")
    oTabela.Consultar_BD
    
    If oTabela.Status <> Encontrado Then
        MsgBox oTabela.Mensagem, vbExclamation, "Atenção"
        Exit Function
    End If
    
    Set Pesquisa = oTabela.Registros
    Set oTabela = Nothing
    
    Exit Function
    
ErroRotina:
    iStatus = Erro
    Set oTabela = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A FUNÇÃO PRIVADA < " & sNomePadrao & "." & "Pesquisa > "
End Function

Sub CriaRecordSetColunasExibir_SQL_Server()
    Dim sAux1() As String
    Dim sAux2() As String
    Dim iAux    As Integer

    Set rsColunasExibir = Nothing
    Set rsColunasExibir = New ADODB.Recordset
    rsColunasExibir.Fields.Append "Coluna", adVarChar, "50"
    rsColunasExibir.Fields.Append "Apelido", adVarChar, "50"
    rsColunasExibir.Open
    
    rsColunasExibir.AddNew
    rsColunasExibir!Coluna = CStr(rsPesquisa.Fields(0).Name)
    rsColunasExibir!Apelido = "Chave Primária"
    rsColunasExibir.Update
    
    If InStr(1, sColunasExibir, ",") = 0 Then
        If InStr(1, sColunasExibir, ";", vbTextCompare) > 0 Then
            sAux2() = Split(sColunasExibir, ";")
        Else
            ReDim Preserve sAux2(0)
            ReDim Preserve sAux2(1)
            sAux2(0) = sColunasExibir
            sAux2(1) = sColunasExibir
        End If
        GoSub ArmazenaColunaApelido
    Else
        sAux1() = Split(sColunasExibir, ",")
        
        For iAux = LBound(sAux1()) To UBound(sAux1())
            If InStr(1, sAux1(iAux), ";", vbTextCompare) > 0 Then
                sAux2() = Split(sAux1(iAux), ";")
            Else
                ReDim Preserve sAux2(0)
                ReDim Preserve sAux2(1)
                sAux2(0) = sAux1(iAux)
                sAux2(1) = sAux1(iAux)
            End If
            GoSub ArmazenaColunaApelido
        Next iAux
    End If
    
    Exit Sub
    
ArmazenaColunaApelido:
    Estrutura_SQL.Estrutura.Filter = ""
    Estrutura_SQL.Estrutura.MoveFirst
    Estrutura_SQL.Estrutura.Filter = "COLUNA='" & sAux2(0) & "'"
    
    If Not Estrutura_SQL.Estrutura.EOF Then
        rsColunasExibir.Filter = ""
        rsColunasExibir.MoveFirst
        rsColunasExibir.Filter = "Coluna = '" & sAux2(0) & "'"
        
        If rsColunasExibir.EOF Then
            rsColunasExibir.AddNew
            rsColunasExibir!Coluna = sAux2(0)
            rsColunasExibir!Apelido = sAux2(1)
            rsColunasExibir.Update
        End If
    End If
Return
End Sub

Function ColunaInformada(pColuna As String) As Boolean
    rsColunasExibir.Filter = ""
    
    If rsColunasExibir.RecordCount = 1 Then
        ColunaInformada = True
        Exit Function
    End If
    
    rsColunasExibir.MoveFirst
    rsColunasExibir.Filter = "Coluna = '" & pColuna & "'"
    If rsColunasExibir.EOF Then
        ColunaInformada = False
    Else
        ColunaInformada = True
    End If
End Function

Function LocalizaColuna(pApelidoColuna As String) As String
    rsColunasExibir.Filter = ""
    rsColunasExibir.MoveFirst
    rsColunasExibir.Filter = "Apelido = '" & pApelidoColuna & "'"
    If rsColunasExibir.EOF Then
        LocalizaColuna = pApelidoColuna
    Else
        LocalizaColuna = rsColunasExibir!Coluna
    End If
End Function

Function LocalizaApelidoColuna(pColuna As String) As String
    rsColunasExibir.Filter = ""
    rsColunasExibir.MoveFirst
    rsColunasExibir.Filter = "Coluna = '" & pColuna & "'"
    If rsColunasExibir.EOF Then
        LocalizaApelidoColuna = pColuna
    Else
        LocalizaApelidoColuna = rsColunasExibir!Apelido
    End If
End Function

Sub Retorna()
    Set rsPesquisa = Nothing
    Unload Me
End Sub

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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sNomePadrao & "." & "LimpaStatus > "
End Sub
