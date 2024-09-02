VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConsultaLOGLocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta LOG Local"
   ClientHeight    =   8160
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11835
   ControlBox      =   0   'False
   Icon            =   "frmConsultaLOGLocal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11835
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
      Height          =   1200
      Left            =   15
      TabIndex        =   8
      Top             =   675
      Width           =   11775
      Begin VB.Frame fraTipoPesquisa 
         Caption         =   "Tipo de Pesquisa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   780
         Left            =   4545
         TabIndex        =   17
         Tag             =   "SELECAO"
         Top             =   315
         Width           =   3165
         Begin VB.OptionButton optTipo 
            Caption         =   "Período"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1620
            TabIndex        =   2
            Tag             =   "SELECAO"
            Top             =   360
            Width           =   1185
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Geral"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   450
            TabIndex        =   1
            Tag             =   "SELECAO"
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame fraFornecedor 
         Caption         =   "Tipo de Aplicação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   780
         Left            =   180
         TabIndex        =   16
         Tag             =   "SELECAO"
         Top             =   315
         Width           =   4155
         Begin VB.ComboBox cboTipoAplicacao 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Tag             =   "SELECAO"
            Top             =   270
            Width           =   3840
         End
      End
      Begin VB.Frame fraPeriodo 
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   780
         Left            =   7920
         TabIndex        =   13
         Tag             =   "SELECAO"
         Top             =   315
         Width           =   3705
         Begin MSMask.MaskEdBox mskDtLog 
            Height          =   360
            Index           =   0
            Left            =   585
            TabIndex        =   3
            Tag             =   "SELECAO"
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDtLog 
            Height          =   360
            Index           =   1
            Left            =   2385
            TabIndex        =   4
            Tag             =   "SELECAO"
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label lblInforma 
            AutoSize        =   -1  'True
            Caption         =   "Até"
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
            Left            =   1980
            TabIndex        =   15
            Tag             =   "SELECAO"
            Top             =   315
            Width           =   300
         End
         Begin VB.Label lblInforma 
            AutoSize        =   -1  'True
            Caption         =   "De"
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
            TabIndex        =   14
            Tag             =   "SELECAO"
            Top             =   315
            Width           =   270
         End
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         TabIndex        =   12
         Top             =   0
         Width           =   11775
      End
   End
   Begin VB.Frame fraNF 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6300
      Left            =   15
      TabIndex        =   7
      Tag             =   "NF"
      Top             =   1875
      Width           =   11775
      Begin RichTextLib.RichTextBox rtbDescricao 
         Height          =   1230
         Left            =   0
         TabIndex        =   6
         Tag             =   "LISTA"
         Top             =   4545
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   2170
         _Version        =   393217
         BackColor       =   15133154
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmConsultaLOGLocal.frx":030A
      End
      Begin MSComctlLib.ListView lvwLog 
         Height          =   3975
         Left            =   135
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "LISTA"
         ToolTipText     =   "Lista de logs gerados"
         Top             =   270
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   7011
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
         Alignment       =   2  'Center
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Tag             =   "LISTA"
         Top             =   4275
         Width           =   11775
      End
      Begin VB.Label lblQtdeLog 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "99 Registro(s) de LOG"
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
         Height          =   240
         Left            =   9630
         TabIndex        =   11
         Tag             =   "LISTA"
         Top             =   5895
         Width           =   1980
      End
      Begin VB.Label lblInforma 
         Alignment       =   2  'Center
         Caption         =   "LOGs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Tag             =   "LISTA"
         Top             =   0
         Width           =   11775
      End
   End
   Begin MSComctlLib.Toolbar tbrManut 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Limpar"
            Object.ToolTipText     =   "< F7 > Limpar a Tela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pesquisar"
            Object.ToolTipText     =   "< F3 > Efetuar Pesquisa "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   500,001
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "< F10 > Sair do Programa"
            ImageIndex      =   7
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   75
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":038C
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":049E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":21AA
            Key             =   "Gravar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":22BC
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":270E
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":2820
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLOGLocal.frx":2B3C
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaLOGLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTituloAplicacao  As String
Private iStatus           As ADRRBR_SIS_Basico.eStatus
Private sMensagem         As String

Dim oParamLog             As Object

Dim rsLog                 As ADODB.Recordset

Dim bIniciado             As Boolean 'Indica se o Programa já está na Memória

Dim sCaminhoArquivoLog(1) As String
Const sArquivoLog = "\ADRRBR_Log.Log"

Enum Colunas_lvwLog
    LV_TipoAplic = 0
    LV_Identif = 1
    LV_Objeto = 2
    LV_Rotina = 3
    LV_Descricao = 4
    LV_DtGeracao = 5
    LV_CodAplic = 6
    LV_NomeAplic = 7
    LV_CodUsuAplic = 8
    LV_NomeUsuAplic = 9
    LV_NomeUsuLocal = 10
    LV_NomeComputador = 11
    LV_FonteConexao = 12
End Enum

Const sNomePadrao = "Consulta LOG Local"

Public Property Let TituloAplicacao(ByVal vNewValue As String)
    sTituloAplicacao = vNewValue
End Property
Public Property Get TituloAplicacao() As String
    TituloAplicacao = sTituloAplicacao
End Property

Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Private Sub Form_Load()
    Me.Left = 30
    Me.Top = 15

    Me.Caption = sTituloAplicacao
    lblTitulo.Caption = sNomePadrao
End Sub

Private Sub Form_Activate()
    If bIniciado Then Exit Sub

    bIniciado = True
    
    If Not EncheComboTiposAplicacao Then
        MsgBox "Não existem tipos de aplicação informadas!", vbExclamation, "Atenção"
        tbrManut_ButtonClick tbrManut.Buttons("Sair")
        Exit Sub
    End If
    
    If Not VerificaParametrosLog Then Exit Sub
    
    CriaColLog
    InicializaTela
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oParamLog = Nothing
    Set rsLog = Nothing
End Sub

Private Sub optTipo_Click(Index As Integer)
    If Index = 0 Then
        fraPeriodo.Visible = False
        
        mskDtLog(0).Mask = Empty
        mskDtLog(0).Text = Empty
        mskDtLog(0).Mask = "##/##/####"
        
        mskDtLog(1).Mask = Empty
        mskDtLog(1).Text = Empty
        mskDtLog(1).Mask = "##/##/####"
    Else
        fraPeriodo.Visible = True
    End If
End Sub

Private Sub lvwLog_GotFocus()
    If lvwLog.ListItems.Count = 0 Then Exit Sub
    
    If lvwLog.SelectedItem.Index = 0 Then oBasico.LV.PosicionaRow_Indice lvwLog, 1, True
    lvwLog_ItemClick lvwLog.ListItems.Item(lvwLog.SelectedItem.Index)
End Sub

Private Sub lvwLog_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwLog.ListItems.Count = 0 Then Exit Sub
    
    rtbDescricao.Text = lvwLog.ListItems(lvwLog.SelectedItem.Index).SubItems(LV_Descricao)
End Sub

Private Sub lvwLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwLog
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

'**** Controle Para Digitação de Datas
Private Sub mskDtLog_Validate(Index As Integer, Cancel As Boolean)
    If Not oBasico.DataHora.ValidarData(mskDtLog(Index).ClipText, mskDtLog(Index).Text) Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub mskDtLog_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
'****

'Navegação Entre os Campos
Private Sub cboTipoAplicacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Teclas de Atalho
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case F3
            If tbrManut.Buttons("Pesquisar").Enabled = True Then
                tbrManut_ButtonClick tbrManut.Buttons("Pesquisar")
            End If
        Case F7
            If tbrManut.Buttons("Limpar").Enabled = True Then
                tbrManut_ButtonClick tbrManut.Buttons("Limpar")
            End If
        Case F10
            If tbrManut.Buttons("Sair").Enabled = True Then
                tbrManut_ButtonClick tbrManut.Buttons("Sair")
            End If
    End Select
End Sub

Private Sub tbrManut_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sAux                As String
    Dim sMsg                As String
    Dim sLinhaSelecionada() As String
    
    Select Case Button.Key
        Case "Limpar"
            InicializaTela
            cboTipoAplicacao.SetFocus
            Exit Sub
        
        Case "Pesquisar"
            If Not VerificaSelecao Then Exit Sub
            If Not EncheListaLogs Then Exit Sub
            
            HabilitaLista True
            oBasico.LV.PosicionaRow_Indice lvwLog, lvwLog.SelectedItem.Index, True

        Case "Sair"
            bIniciado = False
            Unload Me
            Set frmConsultaLOGLocal = Nothing
            Set rsLog = Nothing
            Exit Sub
    End Select
End Sub

'************************
'  Funções/Sub-Rotinas
'************************
Sub InicializaTela()
    sCaminhoArquivoLog(0) = Empty
    sCaminhoArquivoLog(1) = Empty
    
    lvwLog.ListItems.Clear
    cboTipoAplicacao.ListIndex = -1
    
    rtbDescricao.Text = Empty
    
    optTipo(0).Value = True
    optTipo_Click 0
    
    mskDtLog(0).Mask = Empty
    mskDtLog(0).Text = Empty
    mskDtLog(0).Mask = "##/##/####"
    
    mskDtLog(1).Mask = Empty
    mskDtLog(1).Text = Empty
    mskDtLog(1).Mask = "##/##/####"
    
    lblQtdeLog.Caption = Empty
    
    HabilitaLista False
End Sub

Sub CriaColLog()
    oBasico.LV.AddCol lvwLog, "Tipo Aplicação", "2000", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Identificação", "2000", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Objeto", "2000", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Rotina", "2000", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Descrição", "0", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Data", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Cód.Aplic.", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Nome Aplic.", "2000", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Cód.Usu.Aplic.", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Nome Usu.Aplic.", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Usu.Local", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Nome Computador", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwLog, "Conexão BD", "3000", Esquerda_LV
End Sub

Function EncheComboTiposAplicacao() As Boolean
    EncheComboTiposAplicacao = True
    
    cboTipoAplicacao.Clear
    cboTipoAplicacao.AddItem "TODOS"
    cboTipoAplicacao.ItemData(cboTipoAplicacao.NewIndex) = 0
        
    cboTipoAplicacao.AddItem "Windows"
    cboTipoAplicacao.ItemData(cboTipoAplicacao.NewIndex) = 1
    
    cboTipoAplicacao.AddItem "Internet"
    cboTipoAplicacao.ItemData(cboTipoAplicacao.NewIndex) = 2
    
    cboTipoAplicacao.ListIndex = 0
End Function

Function ValidaPeriodo() As Boolean
    Dim iDia  As Integer
    Dim iMes  As Integer
    Dim lAno  As Long
    Dim sHora As String

    ValidaPeriodo = True
    
    If mskDtLog(0).ClipText = Empty Then
        MsgBox "Informe a data inicial do período!", vbExclamation, "Atenção"
        ValidaPeriodo = False
        mskDtLog(0).SetFocus
        Exit Function
    End If
    If Not oBasico.DataHora.DataValida(mskDtLog(0).Text, iDia, iMes, lAno, sHora) Then
        MsgBox "Informe uma data inicial do período válida!", vbExclamation, "Atenção"
        ValidaPeriodo = False
        mskDtLog(0).SetFocus
        Exit Function
    End If
    If mskDtLog(1).ClipText = Empty Then
        MsgBox "Informe a data final do período!", vbExclamation, "Atenção"
        ValidaPeriodo = False
        mskDtLog(1).SetFocus
        Exit Function
    End If
    If Not oBasico.DataHora.DataValida(mskDtLog(1).Text, iDia, iMes, lAno, sHora) Then
        MsgBox "Informe uma data final do período válida!", vbExclamation, "Atenção"
        ValidaPeriodo = False
        mskDtLog(1).SetFocus
        Exit Function
    End If
    If CDate(mskDtLog(1).Text) < CDate(mskDtLog(0).Text) Then
        MsgBox "A data final do período não pode ser menor que a data inicial!", vbExclamation, "Atenção"
        ValidaPeriodo = False
        mskDtLog(1).SetFocus
        Exit Function
    End If
End Function

Function EncheListaLogs() As Boolean
    Dim bPeriodoValido As Boolean
    Dim sColunaAux     As String
    Dim sConteudoAux   As String
    Dim iArquivo       As Integer
    
    EncheListaLogs = True
    
    Screen.MousePointer = vbHourglass
    
    lvwLog.ListItems.Clear
    
    For iArquivo = 0 To 1
        If Trim(sCaminhoArquivoLog(iArquivo)) <> Empty And Dir(sCaminhoArquivoLog(iArquivo)) <> Empty Then
            Set rsLog = Nothing
            Set rsLog = New ADODB.Recordset
            rsLog.Open sCaminhoArquivoLog(iArquivo)
            
            While Not rsLog.EOF
                bPeriodoValido = True
                
                'Seleção Por Período
                If optTipo(1).Value = True Then
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Data Geração", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    
                    If CDate(Left(sConteudoAux, 10)) < CDate(mskDtLog(0).Text) Or CDate(Left(sConteudoAux, 10)) > CDate(mskDtLog(1).Text) Then
                        bPeriodoValido = False
                    End If
                End If
                
                If bPeriodoValido Then
                    lvwLog.ListItems.Add
                    
                    If iArquivo = 0 Then
                        lvwLog.ListItems(lvwLog.ListItems.Count).Text = "Windows"
                    ElseIf iArquivo = 1 Then
                        lvwLog.ListItems(lvwLog.ListItems.Count).Text = "Internet"
                    Else
                        lvwLog.ListItems(lvwLog.ListItems.Count).Text = Empty
                    End If
                    
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Identificação", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_Identif) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Objeto", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_Objeto) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Rotina Fonte", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_Rotina) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Descrição", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_Descricao) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Data Geração", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_DtGeracao) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Código Aplicação", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_CodAplic) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Aplicação", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_NomeAplic) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Código Usuário Aplicação", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_CodUsuAplic) = sConteudoAux
            
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Usuário Aplicação", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_NomeUsuAplic) = sConteudoAux
                    
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Usuário Local", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_NomeUsuLocal) = sConteudoAux
                    
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Nome Computador", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_NomeComputador) = sConteudoAux
                    
                    sColunaAux = oBasico.Geral.EncriptarDecriptar("Fonte Conexão BD", True)
                    sConteudoAux = oBasico.Geral.EncriptarDecriptar(rsLog.Fields(sColunaAux).Value, False)
                    lvwLog.ListItems(lvwLog.ListItems.Count).SubItems(LV_FonteConexao) = sConteudoAux
                End If
                
                rsLog.MoveNext
            Wend
        End If
    Next iArquivo
    
    Screen.MousePointer = vbDefault
    
    Set rsLog = Nothing
    
    If lvwLog.ListItems.Count = 1 Then
        lblQtdeLog.Caption = lvwLog.ListItems.Count & " registro de log"
    ElseIf lvwLog.ListItems.Count > 1 Then
        lblQtdeLog.Caption = lvwLog.ListItems.Count & " registros de log"
    Else
        lblQtdeLog.Caption = Empty
    End If
    
    If lvwLog.ListItems.Count = 0 Then
        EncheListaLogs = False
        MsgBox "Não existem registros de log para esta seleção!", vbExclamation, "Atenção"
    End If
End Function

Function VerificaSelecao() As Boolean
    Dim sMensagem As String
    
    VerificaSelecao = True
    
    If cboTipoAplicacao.ListIndex = -1 Then
        MsgBox "Informe um tipo de aplicação ou TODOS para consulta!", vbExclamation, "Atenção"
        cboTipoAplicacao.SetFocus
        VerificaSelecao = False
        Exit Function
    End If
    
    sCaminhoArquivoLog(0) = Empty
    sCaminhoArquivoLog(1) = Empty
    
    Select Case UCase(cboTipoAplicacao.List(cboTipoAplicacao.ListIndex))
        Case "WINDOWS"
            sCaminhoArquivoLog(0) = oParamLog.CaminhoLogLocal & sArquivoLog
            If Dir(sCaminhoArquivoLog(0)) = Empty Then
                VerificaSelecao = False
                sMensagem = "Não existe o arquivo de log referente à aplicação < " & cboTipoAplicacao.List(cboTipoAplicacao.ListIndex) & " >!"
            End If
            
        Case "INTERNET"
            sCaminhoArquivoLog(1) = oParamLog.CaminhoFisicoInternet & sArquivoLog
            If Dir(sCaminhoArquivoLog(1)) = Empty Then
                VerificaSelecao = False
                sMensagem = "Não existe o arquivo de log referente à aplicação < " & cboTipoAplicacao.List(cboTipoAplicacao.ListIndex) & " >!"
            End If
        
        Case "TODOS"
            sCaminhoArquivoLog(0) = oParamLog.CaminhoLogLocal & sArquivoLog
            sCaminhoArquivoLog(1) = oParamLog.CaminhoFisicoInternet & sArquivoLog
            
            If Dir(sCaminhoArquivoLog(0)) = Empty Then VerificaSelecao = False
            If Not VerificaSelecao And Dir(sCaminhoArquivoLog(1)) <> Empty Then VerificaSelecao = True
            If Not VerificaSelecao Then sMensagem = "Não existem os arquivos de log referente às aplicações!"
    End Select
    
    If Not VerificaSelecao Then
        MsgBox sMensagem, vbExclamation, "Atenção"
        cboTipoAplicacao.SetFocus
        Exit Function
    End If

    If optTipo(1).Value = True Then 'Seleção Por Período?
        If Not ValidaPeriodo Then
            VerificaSelecao = False
            Exit Function
        End If
    End If
End Function

Function VerificaParametrosLog() As Boolean
    On Error GoTo ErroRotina
    
    VerificaParametrosLog = True
    
    Set oParamLog = Nothing
    Set oParamLog = CreateObject("ADRRBR_SIS_Param_Log.clsSIS_Param_Log")
    
    oParamLog.RecuperaParametrosLog
    
    If oParamLog.Status <> Sucesso Then
        iStatus = oParamLog.Status
        sMensagem = oParamLog.Mensagem
        VerificaParametrosLog = False
        Exit Function
    End If
    
    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaParametrosLog = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A FUNÇÃO < frmConsultaLOGLocal.VerificaParametrosLog >"
End Function

Sub HabilitaLista(pHabilita As Boolean)
    'Esta Rotina Habilita ou Desabilita todos os Componentes da tela com a Propriedade
    'TAG = "SELECAO"/"LISTA"
     
    Dim oObjeto As Object
    
    For Each oObjeto In Me
        Select Case Trim(UCase(oObjeto.Tag))
            Case "LISTA"
                oObjeto.Visible = pHabilita
                            
            Case "SELECAO"
                If pHabilita Then
                    oObjeto.Enabled = False
                Else
                    oObjeto.Enabled = True
                End If
        End Select
    Next oObjeto
    
    If pHabilita Then
        tbrManut.Buttons("Pesquisar").Enabled = False
        lvwLog.BackColor = AzulGelo
    Else
        tbrManut.Buttons("Pesquisar").Enabled = True
        lvwLog.BackColor = CorBotao
    End If
End Sub

