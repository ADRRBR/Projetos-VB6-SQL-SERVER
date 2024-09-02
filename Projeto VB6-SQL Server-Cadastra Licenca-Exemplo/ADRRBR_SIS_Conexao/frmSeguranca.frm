VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSeguranca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segurança para Login no Banco de Dados"
   ClientHeight    =   4095
   ClientLeft      =   5535
   ClientTop       =   3690
   ClientWidth     =   4620
   ControlBox      =   0   'False
   HelpContextID   =   17
   Icon            =   "frmSeguranca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgAbrir 
      Left            =   45
      Top             =   3555
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabSeguranca 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros de &Conexão"
      TabPicture(0)   =   "frmSeguranca.frx":27A2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraAccess"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSenha"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraTipoBancoDados"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraSQLServer"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Aplicação &Internet"
      TabPicture(1)   =   "frmSeguranca.frx":27BE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraCaminhoFisico"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraCaminhoFisico 
         ForeColor       =   &H00404000&
         Height          =   1365
         Left            =   90
         TabIndex        =   25
         Top             =   375
         Width           =   4440
         Begin VB.CheckBox chkInternet 
            Caption         =   "Utilizar os Parâmetros para Aplicação Internet?"
            Height          =   240
            Left            =   135
            TabIndex        =   9
            Top             =   225
            Width           =   3615
         End
         Begin VB.CommandButton cmdCaminhoFisico 
            Caption         =   "..."
            Height          =   315
            Left            =   4035
            TabIndex        =   15
            ToolTipText     =   "Localiza o Caminho..."
            Top             =   855
            Width           =   315
         End
         Begin VB.TextBox txtCaminhoFisico 
            Height          =   315
            Left            =   120
            MaxLength       =   500
            TabIndex        =   10
            Top             =   855
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   "Caminho Físico"
            Height          =   165
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1155
         End
      End
      Begin VB.Frame fraSQLServer 
         ForeColor       =   &H00404000&
         Height          =   1410
         Left            =   -74930
         TabIndex        =   21
         Top             =   1365
         Width           =   4440
         Begin VB.TextBox txtServidor 
            Height          =   315
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   3
            Top             =   225
            Width           =   2865
         End
         Begin VB.TextBox txtBancoDados 
            Height          =   315
            Left            =   1410
            MaxLength       =   40
            TabIndex        =   4
            Top             =   600
            Width           =   2865
         End
         Begin VB.TextBox txtUsuario 
            Height          =   315
            Left            =   1410
            MaxLength       =   40
            TabIndex        =   5
            Top             =   975
            Width           =   2865
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Servidor"
            Height          =   165
            Index           =   0
            Left            =   690
            TabIndex        =   24
            Top             =   285
            Width           =   615
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Banco de Dados"
            Height          =   165
            Index           =   1
            Left            =   60
            TabIndex        =   23
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuário"
            Height          =   165
            Index           =   2
            Left            =   690
            TabIndex        =   22
            Top             =   1020
            Width           =   615
         End
      End
      Begin VB.Frame fraTipoBancoDados 
         Caption         =   "Tipo de Banco de Dados"
         Height          =   990
         Left            =   -74930
         TabIndex        =   20
         Top             =   375
         Width           =   4440
         Begin VB.ComboBox cmbTipoBD 
            Height          =   315
            Left            =   105
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   300
            Width           =   4215
         End
         Begin VB.CheckBox chkConexaoUsuarioWindows 
            Caption         =   "Conexão utilizando o usuário do Windows?"
            Height          =   240
            Left            =   960
            TabIndex        =   2
            Top             =   675
            Width           =   3390
         End
      End
      Begin VB.Frame fraSenha 
         Height          =   690
         Left            =   -74930
         TabIndex        =   18
         Top             =   2805
         Width           =   4440
         Begin VB.TextBox txtSenha 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2895
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   235
            Width           =   1365
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Senha"
            Height          =   165
            Index           =   3
            Left            =   2190
            TabIndex        =   19
            Top             =   285
            Width           =   615
         End
      End
      Begin VB.Frame fraAccess 
         ForeColor       =   &H00404000&
         Height          =   1410
         Left            =   -74930
         TabIndex        =   16
         Top             =   1365
         Width           =   4440
         Begin VB.TextBox txtCaminhoMDB 
            Height          =   315
            Left            =   120
            MaxLength       =   500
            TabIndex        =   6
            Top             =   675
            Width           =   3855
         End
         Begin VB.CommandButton cmdCaminhoMDB 
            Caption         =   "..."
            Height          =   315
            Left            =   4035
            TabIndex        =   7
            ToolTipText     =   "Localiza o Arquivo MDB..."
            Top             =   675
            Width           =   315
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Caption         =   "Caminho MDB"
            Height          =   165
            Index           =   4
            Left            =   75
            TabIndex        =   17
            Top             =   420
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3405
      TabIndex        =   12
      Top             =   3720
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2190
      TabIndex        =   11
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Frame fraPastas 
      Height          =   1770
      Left            =   135
      TabIndex        =   27
      Top             =   1620
      Width           =   4335
      Begin VB.DirListBox Pastas 
         Height          =   990
         Left            =   45
         TabIndex        =   14
         Top             =   585
         Width           =   4245
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   45
         TabIndex        =   13
         Top             =   180
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSeguranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iStatus   As ADRRBR_SIS_Basico.eStatus
Private sMensagem As String

Dim oBasico  As New ADRRBR_SIS_Basico.clsSIS_Basico
Dim oConexao As Object

Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Private Sub Form_Unload(Cancel As Integer)
    Set oConexao = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
    
Private Sub cmbTipoBD_GotFocus()
    tabSeguranca.Tab = 0
End Sub

Private Sub txtSenha_GotFocus()
    tabSeguranca.Tab = 0
End Sub

Private Sub cmdCaminhoMDB_Click()
    txtCaminhoMDB.Text = Empty
    
    dlgAbrir.DialogTitle = "Arquivo MDB (Access)"
    dlgAbrir.Filter = "Microsoft Access|*.mdb"
    dlgAbrir.FileName = Empty
    dlgAbrir.ShowOpen

    If dlgAbrir.FileName <> Empty Then
        If Dir(dlgAbrir.FileName, vbDirectory) <> Empty Then txtCaminhoMDB.Text = dlgAbrir.FileName
    End If
End Sub

Private Sub txtCaminhoMDB_Validate(Cancel As Boolean)
    If Trim(txtCaminhoMDB.Text) = Empty Then Exit Sub
    
    On Error GoTo Erro_Rotina
    
    If Dir(txtCaminhoMDB.Text) = Empty Then GoTo Erro_Rotina
    
    Exit Sub
    
Erro_Rotina:
    Cancel = True
    MsgBox "Arquivo não encontrado!", vbExclamation, "Atenção"
    txtCaminhoMDB.SetFocus
End Sub

Private Sub chkInternet_GotFocus()
    tabSeguranca.Tab = 1
End Sub

Private Sub cmdCaminhoFisico_GotFocus()
    tabSeguranca.Tab = 1
End Sub

Private Sub cmbTipoBD_Click()
    LimpaCampos
    
    Select Case cmbTipoBD.ItemData(cmbTipoBD.ListIndex)
        Case SQL_Server
            chkConexaoUsuarioWindows.Visible = True
            fraSQLServer.Visible = True
            fraAccess.Visible = False
        
        Case Access
            chkConexaoUsuarioWindows.Visible = False
            fraSQLServer.Visible = False
            fraAccess.Visible = True
            txtSenha.Enabled = True: txtSenha.BackColor = Branco
    End Select
    
    fraSenha.Visible = True
    cmdOK.Visible = True
End Sub
    
Private Sub chkConexaoUsuarioWindows_Click()
    If chkConexaoUsuarioWindows.Value = 1 Then
        txtUsuario.Text = Empty
        txtSenha.Text = Empty
        txtUsuario.Enabled = False: txtUsuario.BackColor = CinzaBotao
        txtSenha.Enabled = False: txtSenha.BackColor = CinzaBotao
    Else
        txtUsuario.Enabled = True: txtUsuario.BackColor = Branco
        txtSenha.Enabled = True: txtSenha.BackColor = Branco
    End If
End Sub
    
Friend Sub chkInternet_Click()
    If chkInternet.Value = 1 Then
        txtCaminhoFisico.Enabled = True: txtCaminhoFisico.BackColor = Branco
        cmdCaminhoFisico.Enabled = True
    Else
        txtCaminhoFisico.Text = Empty
        txtCaminhoFisico.Enabled = False: txtCaminhoFisico.BackColor = CinzaBotao
        cmdCaminhoFisico.Enabled = False
    End If
End Sub
    
Private Sub cmdCaminhoFisico_Click()
    fraPastas.ZOrder 0
    fraPastas.Visible = True
    Drive.SetFocus
End Sub

Private Sub txtCaminhoFisico_Validate(Cancel As Boolean)
    If Trim(txtCaminhoFisico.Text) = Empty Then Exit Sub
    
    On Error GoTo Erro_Rotina
    
    If Dir(txtCaminhoFisico.Text, vbDirectory) = Empty Then GoTo Erro_Rotina
    
    Exit Sub
    
Erro_Rotina:
    Cancel = True
    MsgBox "Caminho inválido!", vbExclamation, "Atenção"
    txtCaminhoFisico.SetFocus
End Sub
    
Private Sub cmdOK_Click()
    If Not VerificaCampos Then Exit Sub
    If Not ValidaConexao Then Exit Sub
    
    If Not GravaParametrosConexao(True) Then
        MsgBox sMensagem, vbCritical, "Atenção"
        cmdOK.SetFocus
        Exit Sub
    End If
    
    If Not GravaParametrosConexao Then
        MsgBox sMensagem, vbCritical, "Atenção"
        cmdOK.SetFocus
        Exit Sub
    End If
    
    iStatus = Sucesso
    sMensagem = "Segurança gerada com sucesso!"
    
    Unload Me
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    LimpaStatus
    
    iStatus = Cancelado
    sMensagem = "Cancelamento do Login de Segurança para acesso ao banco de dados!"
    
    Unload Me
    Exit Sub
End Sub
    
Private Sub Drive_Change()
    On Error GoTo Erro
    
    Pastas.Path = Drive
    txtCaminhoFisico.Text = Pastas.Path
    
    Exit Sub
    
Erro:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Atenção"
End Sub
Private Sub Drive_LostFocus()
    If UCase(Me.ActiveControl.Name) <> "PASTAS" Then fraPastas.Visible = False
End Sub

Private Sub Pastas_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Pastas.Path = Pastas.List(Pastas.ListIndex)
End Sub
Private Sub Pastas_Change()
    Pastas.Path = Drive.Drive
    txtCaminhoFisico.Text = Pastas.List(Pastas.ListIndex)
End Sub
Private Sub Pastas_DblClick()
    txtCaminhoFisico.Text = Pastas.List(Pastas.ListIndex)
    fraPastas.Visible = False
End Sub
Private Sub Pastas_LostFocus()
    If UCase(Me.ActiveControl.Name) <> "DRIVE" Then fraPastas.Visible = False
End Sub

Sub InicializaTela()
    Set oConexao = Nothing
    Set oConexao = CreateObject("ADRRBR_SIS_Conexao.clsSIS_Conexao")
    
    LimpaStatus
    
    oBasico.BD.Status = Sucesso
    EncheComboTipoBancoDados
    LimpaCampos
    chkConexaoUsuarioWindows.Visible = False
    fraSQLServer.Visible = False
    fraAccess.Visible = False
    fraSenha.Visible = False
    cmdOK.Visible = False
    
    tabSeguranca.Tab = 0
    
    tabSeguranca.TabVisible(1) = False
End Sub
    
Sub LimpaCampos()
    chkConexaoUsuarioWindows.Value = 0
    chkConexaoUsuarioWindows_Click
    
    txtServidor.Text = Empty
    txtBancoDados.Text = Empty
    txtUsuario.Text = Empty
    txtCaminhoMDB.Text = Empty
    txtSenha.Text = Empty
    
    chkInternet.Value = 0
    chkInternet_Click
    
    txtCaminhoFisico.Text = Empty
End Sub
    
Sub EncheComboTipoBancoDados()
    cmbTipoBD.Clear
    cmbTipoBD.AddItem "SQL Server"
    cmbTipoBD.ItemData(cmbTipoBD.NewIndex) = SQL_Server
    cmbTipoBD.AddItem "Access"
    cmbTipoBD.ItemData(cmbTipoBD.NewIndex) = Access
End Sub

Function VerificaCampos() As Boolean
    VerificaCampos = True
    
    'Parâmetros de Conexão
    If cmbTipoBD.ListIndex = -1 Then
        VerificaCampos = False
        tabSeguranca.Tab = 0
        MsgBox "Informe o tipo de banco de dados!", vbExclamation, "Atenção"
        cmbTipoBD.SetFocus
        Exit Function
    End If
    
    Select Case cmbTipoBD.ItemData(cmbTipoBD.ListIndex)
        Case SQL_Server
            If txtServidor.Text = Empty Then
                VerificaCampos = False
                tabSeguranca.Tab = 0
                MsgBox "Informe o servidor de banco de dados!", vbExclamation, "Atenção"
                txtServidor.SetFocus
                Exit Function
            End If
            If txtBancoDados.Text = Empty Then
                VerificaCampos = False
                tabSeguranca.Tab = 0
                MsgBox "Informe o banco de dados!", vbExclamation, "Atenção"
                txtBancoDados.SetFocus
                Exit Function
            End If
            If chkConexaoUsuarioWindows.Value = 0 Then
                If txtUsuario.Text = Empty Then
                    VerificaCampos = False
                    tabSeguranca.Tab = 0
                    MsgBox "Informe o usuário do banco de dados!", vbExclamation, "Atenção"
                    txtUsuario.SetFocus
                    Exit Function
                End If
            End If
            
        Case Access
            If txtCaminhoMDB.Text = Empty Then
                VerificaCampos = False
                tabSeguranca.Tab = 0
                MsgBox "Informe o caminho do arquivo MDB (Access)!", vbExclamation, "Atenção"
                txtCaminhoMDB.SetFocus
                Exit Function
            End If
    End Select
    
    'Aplicação Internet
    If chkInternet.Value = 1 Then
        If txtCaminhoFisico.Text = Empty Then
            VerificaCampos = False
            tabSeguranca.Tab = 1
            MsgBox "Informe o caminho físco da Aplicação Internet!", vbExclamation, "Atenção"
            txtCaminhoFisico.SetFocus
            Exit Function
        End If
    End If
End Function

Function ValidaConexao() As Boolean
    oConexao.TipoBancoDados = cmbTipoBD.ItemData(cmbTipoBD.ListIndex)
    
    Select Case cmbTipoBD.ItemData(cmbTipoBD.ListIndex)
        Case SQL_Server
            If chkConexaoUsuarioWindows.Value = 1 Then
                oConexao.ConexaoUsuarioWindows = True
            Else
                oConexao.ConexaoUsuarioWindows = False
            End If
            oConexao.Servidor = txtServidor.Text
            oConexao.BancoDados = txtBancoDados.Text
            oConexao.Usuario = txtUsuario.Text
            oConexao.Senha = txtSenha.Text
        
        Case Access
            oConexao.CaminhoMDB = txtCaminhoMDB.Text
            oConexao.Senha = txtSenha.Text
    End Select
    
    oConexao.Conecta
    ValidaConexao = oConexao.Conectado
    
    If Not oConexao.Status = oBasico.BD.Status Then
        MsgBox oConexao.Mensagem, vbCritical, "Atenção"
        Exit Function
    End If
End Function

Function GravaParametrosConexao(Optional pLimparAtuais As Boolean) As Boolean
    Dim sIdentificacao              As String
    Dim sSubIdentificacao           As String
    Dim sDescricaoInformacao        As String
    Dim sConteudoInformacao         As String
    Dim sColunaAux                  As String
    Dim sCaminhoConfConexaoInternet As String
    Dim rsAux                       As ADODB.Recordset
    Dim rsArqParamConexao           As ADODB.Recordset
        
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    GravaParametrosConexao = True
    
    sCaminhoConfConexaoInternet = txtCaminhoFisico.Text & sArqParamConexao
    
    Set rsAux = Nothing
    Set rsAux = New ADODB.Recordset
    rsAux.Fields.Append "DescricaoInformacao", adVarChar, "50"
    rsAux.Fields.Append "ConteudoInformacao", adVarChar, "50"
    rsAux.Open
    
    sIdentificacao = oBasico.Geral.EncriptarDecriptar("ADRRBR", True)
    
    Select Case cmbTipoBD.ItemData(cmbTipoBD.ListIndex)
        Case SQL_Server
            sSubIdentificacao = oBasico.Geral.EncriptarDecriptar("SEGURANCA SQL SERVER", True)
            
            sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("SERVIDOR", True)
            sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", txtServidor.Text), True)
            GoSub GravaInformacao
            
            sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("BANCO DE DADOS", True)
            sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", txtBancoDados.Text), True)
            GoSub GravaInformacao

            sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("CONEXAO USUARIO WINDOWS", True)
            sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", chkConexaoUsuarioWindows.Value), True)
            GoSub GravaInformacao

            sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("USUARIO", True)
            sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", txtUsuario.Text), True)
            GoSub GravaInformacao

        Case Access
            sSubIdentificacao = oBasico.Geral.EncriptarDecriptar("SEGURANCA ACCESS", True)
        
            sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("CAMINHO MDB", True)
            sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", txtCaminhoMDB.Text), True)
            GoSub GravaInformacao
            
        Case Else
            GravaParametrosConexao = False
            iStatus = Erro
            sMensagem = "Tipo de Banco de Dados Inválido!"
            Set rsAux = Nothing
            Exit Function
    End Select

    sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("SENHA", True)
    sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", txtSenha.Text), True)
    GoSub GravaInformacao

    sDescricaoInformacao = oBasico.Geral.EncriptarDecriptar("CAMINHO FISICO INTERNET", True)
    sConteudoInformacao = oBasico.Geral.EncriptarDecriptar(IIf(pLimparAtuais, "", txtCaminhoFisico.Text), True)
    GoSub GravaInformacao
    If pLimparAtuais And Dir(sCaminhoConfConexaoInternet) <> Empty Then Kill sCaminhoConfConexaoInternet
    
    If Not pLimparAtuais And chkInternet.Value = 1 Then
        Set rsArqParamConexao = Nothing
        Set rsArqParamConexao = New ADODB.Recordset
        
        sColunaAux = oBasico.Geral.EncriptarDecriptar("TIPO SEGURANCA", True)
        rsArqParamConexao.Fields.Append sColunaAux, adVarChar, "50"
        rsAux.MoveFirst
        While Not rsAux.EOF
            rsArqParamConexao.Fields.Append rsAux!DescricaoInformacao, adVarChar, "50"
            rsAux.MoveNext
        Wend
        
        rsArqParamConexao.Open
        rsArqParamConexao.AddNew
        rsArqParamConexao.Fields(sColunaAux).Value = sSubIdentificacao
        rsAux.MoveFirst
        While Not rsAux.EOF
            sColunaAux = rsAux!DescricaoInformacao
            rsArqParamConexao.Fields(sColunaAux).Value = rsAux!ConteudoInformacao
            rsAux.MoveNext
        Wend
        
        rsArqParamConexao.Update
                
        If Dir(sCaminhoConfConexaoInternet) <> Empty Then Kill sCaminhoConfConexaoInternet
        rsArqParamConexao.Save sCaminhoConfConexaoInternet
    End If
    
    Set rsAux = Nothing
    Set rsArqParamConexao = Nothing

    Exit Function
    
GravaInformacao:
    oBasico.Geral.GravaRegistroWindows sIdentificacao, sSubIdentificacao, sDescricaoInformacao, sConteudoInformacao
    
    If Not pLimparAtuais Then
        rsAux.AddNew
        rsAux!DescricaoInformacao = sDescricaoInformacao
        rsAux!ConteudoInformacao = sConteudoInformacao
        rsAux.Update
    End If
Return

ErroRotina:
    iStatus = Erro
    GravaParametrosConexao = False
    Set rsAux = Nothing
    Set rsArqParamConexao = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A FUNÇÃO < frmSeguranca.GravaParametrosConexao >"
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
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < frmSeguranca.LimpaStatus >"
End Sub

