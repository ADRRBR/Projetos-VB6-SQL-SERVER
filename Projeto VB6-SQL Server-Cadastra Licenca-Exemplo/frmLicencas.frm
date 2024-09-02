VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLicencas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção Licenças"
   ClientHeight    =   6420
   ClientLeft      =   3945
   ClientTop       =   2115
   ClientWidth     =   12570
   Icon            =   "frmLicencas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   12570
   Begin VB.Frame fraCliente 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8445
      Left            =   -30
      TabIndex        =   6
      Top             =   0
      Width           =   12585
      Begin VB.TextBox txtDiasExpirar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         MaxLength       =   3
         TabIndex        =   15
         Top             =   5625
         Width           =   585
      End
      Begin VB.TextBox txtSerial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         MaxLength       =   1000
         TabIndex        =   12
         Top             =   4905
         Width           =   4365
      End
      Begin VB.CommandButton BotaoManutencao 
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   90
         Picture         =   "frmLicencas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " < F7 > Excluir"
         Top             =   2385
         Width           =   600
      End
      Begin VB.TextBox txtNomeSoftware 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         MaxLength       =   100
         TabIndex        =   4
         Top             =   4185
         Width           =   4365
      End
      Begin VB.ComboBox cboTiposSoftware 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmLicencas.frx":074C
         Left            =   5085
         List            =   "frmLicencas.frx":074E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4185
         Width           =   4155
      End
      Begin VB.CommandButton BotaoManutencao 
         Height          =   390
         Index           =   0
         Left            =   135
         Picture         =   "frmLicencas.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "<  F5 > Habilitar Inclusão"
         Top             =   3330
         Width           =   420
      End
      Begin VB.CommandButton BotaoManutencao 
         Height          =   390
         Index           =   1
         Left            =   630
         Picture         =   "frmLicencas.frx":0ADF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "< F6 > Atualizar"
         Top             =   3330
         Width           =   420
      End
      Begin MSComctlLib.ListView lvwLicencas 
         Height          =   2085
         Left            =   90
         TabIndex        =   0
         Tag             =   "LISTA"
         ToolTipText     =   "Lista de licenças cadastradas"
         Top             =   270
         Width           =   12450
         _ExtentX        =   21960
         _ExtentY        =   3678
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   4210752
         BackColor       =   15133154
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblInforma 
         AutoSize        =   -1  'True
         Caption         =   "Dias expirar a partir de hoje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Tag             =   "CAMPOS"
         Top             =   5355
         Width           =   2460
      End
      Begin VB.Label lblInforma 
         AutoSize        =   -1  'True
         Caption         =   "Serial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Tag             =   "CAMPOS"
         Top             =   4590
         Width           =   525
      End
      Begin VB.Label lblInforma 
         AutoSize        =   -1  'True
         Caption         =   "Serial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Tag             =   "CAMPOS"
         Top             =   4320
         Width           =   525
      End
      Begin VB.Label lblInforma 
         AutoSize        =   -1  'True
         Caption         =   "Nome Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   11
         Tag             =   "CAMPOS"
         Top             =   3915
         Width           =   1380
      End
      Begin VB.Label lblInforma 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   74
         Left            =   5055
         TabIndex        =   10
         Tag             =   "CAMPOS"
         Top             =   3915
         Width           =   1245
      End
      Begin VB.Label lblQtdeLicencas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "99 Registro(s) de Licenças"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   10065
         TabIndex        =   9
         Tag             =   "LISTA"
         Top             =   2520
         Width           =   2385
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Licenças Cadastradas"
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
         Left            =   0
         TabIndex        =   8
         Tag             =   "LISTA"
         Top             =   0
         Width           =   12600
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Manutenção Licenças"
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
         Left            =   45
         TabIndex        =   7
         Tag             =   "SELECAO"
         Top             =   2970
         Width           =   12510
      End
   End
End
Attribute VB_Name = "frmLicencas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oConexao     As ADRRBR_SIS_Conexao.clsSIS_Conexao
Dim bIniciado    As Boolean 'Indica se o Programa já está na Memória
 
Dim sTipoManut   As String
Dim lID_Software As Long
 
Private Sub Form_Load()
    InicializaTela
End Sub

Private Sub Form_Activate()
    If bIniciado Then Exit Sub
    bIniciado = True
    
    If Not ConectaSQLServer Then
        Unload Me
        Exit Sub
    End If
    
    Geral.PosicionaTela Screen, Me

    Consulta
End Sub

Private Sub txtDiasExpirar_KeyPress(KeyAscii As Integer)
    KeyAscii = Geral.TrataNumeros(KeyAscii)
End Sub

'Teclas de Atalho
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case F5 'Inclui
            If BotaoManutencao(0).Enabled = True Then
                BotaoManutencao_Click 0
            End If
        Case F6 'Grava
            If BotaoManutencao(1).Enabled = True Then
                BotaoManutencao_Click 1
            End If
        Case F7 'Exclui
            If BotaoManutencao(2).Enabled = True Then
                BotaoManutencao_Click 2
            End If
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bIniciado = False
    Set frmLicencas = Nothing
End Sub

Private Sub lvwLicencas_GotFocus()
    If lvwLicencas.ListItems.Count = 0 Then Exit Sub
    
    If lvwLicencas.SelectedItem.Index = 0 Then LV.PosicionaRow_Indice lvwLicencas, 1, True
    lvwLicencas_ItemClick lvwLicencas.ListItems.Item(lvwLicencas.SelectedItem.Index)
End Sub

Private Sub lvwLicencas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwLicencas.ListItems.Count = 0 Then Exit Sub
    
    LimpaCampos
    EncheCampos

    STATUS_Operacao Consultar
End Sub

Private Sub lvwLicencas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwLicencas
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

Private Sub BotaoManutencao_Click(Index As Integer)
    Select Case Index
        Case 0 'Inclui
            LimpaCampos
            sTipoManut = "INC"
            STATUS_Operacao Incluir
            txtNomeSoftware.SetFocus
            Exit Sub
            
        Case 1 'Grava
            If Not VerificaCampos Then Exit Sub
            If sTipoManut = Empty Then sTipoManut = "ALT"
            If Not AtualizaRegistro(sTipoManut) Then Exit Sub
            sTipoManut = Empty
            LimpaCampos
            Consulta
            Exit Sub
        
        Case 2 'Exclui
            If MsgBox("Deseja realmente excluir a licença do software:" & vbLf & vbLf & Trim(txtNomeSoftware.Text) & " do tipo " & cboTiposSoftware.List(cboTiposSoftware.ListIndex) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbNo Then
                Exit Sub
            End If
            If Not AtualizaRegistro("EXC") Then Exit Sub
            sTipoManut = Empty
            LimpaCampos
            Consulta
            Exit Sub
    End Select
End Sub

Sub LimpaCampos()
    cboTiposSoftware.ListIndex = -1
    lID_Software = 0
    txtNomeSoftware.Text = Empty
    txtSerial.Text = Empty
    txtDiasExpirar.Text = Empty
End Sub

Sub EncheCampos()
    Dim dDataExpira As Date
    
    dDataExpira = lvwLicencas.ListItems(lvwLicencas.SelectedItem.Index).SubItems(4)
    
    lID_Software = lvwLicencas.ListItems(lvwLicencas.SelectedItem.Index).Text
    txtNomeSoftware.Text = lvwLicencas.ListItems(lvwLicencas.SelectedItem.Index).SubItems(1)
    Geral.PosicionaCombo_Conteudo cboTiposSoftware, lvwLicencas.ListItems(lvwLicencas.SelectedItem.Index).SubItems(2)
    txtSerial.Text = lvwLicencas.ListItems(lvwLicencas.SelectedItem.Index).SubItems(3)
    
    txtDiasExpirar.Text = DateDiff("d", Now(), dDataExpira)
End Sub

Function VerificaCampos() As Boolean
    Dim sCondicaoPesquisa As String
    
    VerificaCampos = False
    
    'Não Validei os Campos Aqu. A Procedure faz a Validação.
   
    VerificaCampos = True
End Function

Function EncheComboTiposSoftware() As Boolean
    EncheComboTiposSoftware = False
    
    cboTiposSoftware.Clear
    
    cboTiposSoftware.AddItem "SO"
    cboTiposSoftware.ItemData(cboTiposSoftware.NewIndex) = 1
        
    cboTiposSoftware.AddItem "OFFICE"
    cboTiposSoftware.ItemData(cboTiposSoftware.NewIndex) = 2
        
    cboTiposSoftware.AddItem "UTILITARIO"
    cboTiposSoftware.ItemData(cboTiposSoftware.NewIndex) = 3

    EncheComboTiposSoftware = True
End Function

Function AtualizaRegistro(pTipoAtualizacao As String) As Boolean
    Dim rsPesquisa As ADODB.Recordset
    Dim sSQL       As String
    
    AtualizaRegistro = False
    
    Screen.MousePointer = vbHourglass

    sSQL = Empty
    sSQL = sSQL & "PRC_Atualiza_Licencas "
    sSQL = sSQL & lID_Software & ", "                                                         '@ID_software"
    sSQL = sSQL & "'" & txtNomeSoftware.Text & "', "                                          '@nome_software"
    sSQL = sSQL & "'" & cboTiposSoftware.Text & "', "                                         '@tipo_software  SO / OFFICE / UTILITARIO"
    sSQL = sSQL & "'" & txtSerial.Text & "', "                                                '@serial"
    sSQL = sSQL & "'" & DataHora.FormataDataGravar(Now() + Val(txtDiasExpirar.Text)) & "', "  '@data_expiracao"
    sSQL = sSQL & "'" & Geral.UsuarioLocal & "', "                                            '@nome_usuario_ult_manut"
    sSQL = sSQL & "'" & pTipoAtualizacao & "'"                                                '@tipo_manut     INC / ALT / EXC"

    Set rsPesquisa = Nothing
    Set rsPesquisa = oConexao.AbreRS(sSQL)

    MsgBox rsPesquisa!Mensagem, vbExclamation, "Atenção"
    
    Screen.MousePointer = vbDefault
    
    Set rsPesquisa = Nothing

    AtualizaRegistro = True
End Function

Sub CriaColLicencas()
    LV.AddCol lvwLicencas, "ID_software", "0", Esquerda_LV
    LV.AddCol lvwLicencas, "Nome Software", "2000", Esquerda_LV
    LV.AddCol lvwLicencas, "Tipo Software", "1800", Esquerda_LV
    LV.AddCol lvwLicencas, "Serial", "1950", Centro_LV
    LV.AddCol lvwLicencas, "Data Expiração", "2200", Centro_LV
    LV.AddCol lvwLicencas, "Data Última Manutenção", "2200", Centro_LV
    LV.AddCol lvwLicencas, "Usuário Última Manutenção", "2200", Centro_LV
End Sub

Function MontaListaLicencas() As Boolean
    Dim rsPesquisa As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    MontaListaLicencas = True
    lblQtdeLicencas.Caption = Empty
    
    lvwLicencas.ListItems.Clear
    
    Set rsPesquisa = Nothing
    Set rsPesquisa = oConexao.AbreRS("SELECT * FROM TAB_LICENCAS")
    If rsPesquisa Is Nothing Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    While Not rsPesquisa.EOF
        lvwLicencas.ListItems.Add
        
        With lvwLicencas.ListItems(lvwLicencas.ListItems.Count)
            .Text = Geral.TrocaNuLL(rsPesquisa!ID_software, Empty)
            .SubItems(1) = Geral.TrocaNuLL(rsPesquisa!nome_software, Empty)
            .SubItems(2) = Geral.TrocaNuLL(rsPesquisa!tipo_software, Empty)
            .SubItems(3) = Geral.TrocaNuLL(rsPesquisa!serial, Empty)
            .SubItems(4) = Format(Geral.TrocaNuLL(rsPesquisa!data_expiracao, Empty), "dd/mm/yyyy hh:mm:ss")
            .SubItems(5) = Format(Geral.TrocaNuLL(rsPesquisa!data_ult_manut, Empty), "dd/mm/yyyy hh:mm:ss")
            .SubItems(6) = Format(Geral.TrocaNuLL(rsPesquisa!nome_usuario_ult_manut, Empty), "dd/mm/yyyy hh:mm:ss")
        End With
        
        rsPesquisa.MoveNext
    Wend

    If lvwLicencas.ListItems.Count = 1 Then
        lblQtdeLicencas.Caption = lvwLicencas.ListItems.Count & " Registro"
    ElseIf lvwLicencas.ListItems.Count > 1 Then
        lblQtdeLicencas.Caption = lvwLicencas.ListItems.Count & " Registros"
    Else
        MontaListaLicencas = False
    End If
    
    Set rsPesquisa = Nothing
    
    Screen.MousePointer = vbDefault
End Function

Sub InicializaTela()
    MousePointer = vbHourglass
    
    CriaColLicencas
    EncheComboTiposSoftware
        
    MousePointer = vbDefault
End Sub

Sub Consulta()
    If MontaListaLicencas Then
        STATUS_Operacao Consultar
        lvwLicencas_GotFocus
        lvwLicencas.SetFocus
        BotaoManutencao(2).Enabled = True
    Else
        BotaoManutencao(2).Enabled = False
        BotaoManutencao_Click 0 'Incluir
        Exit Sub
    End If
End Sub

Sub HabilitaCampos(pHabilita As Boolean)
    Dim bTrava    As Boolean
    Dim CorCampos As Long
    
    If pHabilita Then
        bTrava = False
        CorCampos = Branco
        lvwLicencas.TabStop = False
    Else
        bTrava = True
        CorCampos = AzulGelo
        lvwLicencas.TabStop = True
    End If
    
    txtNomeSoftware.Locked = bTrava: txtNomeSoftware.BackColor = CorCampos
    cboTiposSoftware.Locked = bTrava: cboTiposSoftware.BackColor = CorCampos
    txtSerial.Locked = bTrava: txtSerial.BackColor = CorCampos
End Sub

Sub STATUS_Operacao(pOperacao As ADRRBR_SIS_Basico.eAcao)
    'Determina o STATUS dos Botões Para Manutenção
    Select Case pOperacao
        Case Incluir
            BotaoManutencao(0).Enabled = False
            BotaoManutencao(1).Enabled = True
            HabilitaCampos True
        
        Case Consultar
            BotaoManutencao(0).Enabled = True
            BotaoManutencao(1).Enabled = True
            HabilitaCampos True
    End Select
End Sub

Function ConectaSQLServer() As Boolean
    ConectaSQLServer = True
    
    Screen.MousePointer = vbHourglass
    
    Set oConexao = New ADRRBR_SIS_Conexao.clsSIS_Conexao
    
    oConexao.TipoBancoDados = SQL_Server
    oConexao.Servidor = "CrossDev"
    oConexao.BancoDados = "TesteDB"
    'oConexao.Usuario = "sa"
    'oConexao.senha = "admin001"
    oConexao.Conecta
    
    If oConexao.Status <> Sucesso Then
        MsgBox oConexao.Mensagem, vbCritical, "Conexão Banco de Dados"
        ConectaSQLServer = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Screen.MousePointer = vbDefault
End Function


