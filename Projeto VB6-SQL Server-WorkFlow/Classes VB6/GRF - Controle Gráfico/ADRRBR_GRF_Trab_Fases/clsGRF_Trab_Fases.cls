VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGRF_Trab_Fases"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lTB_pk_trabalho As Variant
Private lTB_fk_trabalho_alt_seq As Variant
Private lTB_fk_trabalho_tipo As Variant
Private sTB_R_TT_tipo_trabalho As Variant
Private lTB_fk_produto As Variant
Private sTB_R_P_produto As Variant
Private lTB_fk_cliente As Variant
Private sTB_R_C_cliente As Variant
Private lTB_fk_representante As Variant
Private sTB_R_R_representante As Variant
Private lTB_fk_aprovacao_tipo As Variant
Private sTB_R_AT_aprovacao_tipo As Variant
Private lTB_fk_filial As Variant
Private sTB_R_F_filial As Variant
Private sTB_num_pedido As Variant
Private sTB_num_pedido_antigo As Variant
Private sTB_num_pedido_novo As Variant
Private sTB_lead_time_programado As Variant
Private dTB_dt_inclusao As Variant
Private dTB_dt_alteracao As Variant
Private dTB_dt_exclusao As Variant
Private lFM1_pk_trabalho_fase_manut1 As Variant
Private lFM1_fk_trabalho_fase_manut1_alt_seq As Variant
Private lFM1_fk_trabalho As Variant
Private lFM1_fk_fase As Variant
Private sFM1_R_F_fase As Variant
Private lFM1_fk_operador As Variant
Private sFM1_R_O_operador As Variant
Private lFM1_cores_alteradas As Variant
Private lFM1_circ_cilindro As Variant
Private sFM1_observacoes As Variant
Private dFM1_dt_inclusao As Variant
Private dFM1_dt_alteracao As Variant
Private lFM2_pk_trabalho_fase_manut2 As Variant
Private lFM2_fk_trabalho_fase_manut2_alt_seq As Variant
Private lFM2_fk_trabalho As Variant
Private lFM2_fk_fase As Variant
Private sFM2_R_F_fase As Variant
Private lFM2_fk_operador As Variant
Private sFM2_R_O_operador As Variant
Private sFM2_observacoes As Variant
Private dFM2_dt_inclusao As Variant
Private dFM2_dt_alteracao As Variant
Private lFM3_pk_trabalho_fase_manut3 As Variant
Private lFM3_fk_trabalho_fase_manut3_alt_seq As Variant
Private lFM3_fk_trabalho As Variant
Private lFM3_fk_fase As Variant
Private sFM3_R_F_fase As Variant
Private lFM3_fk_operador As Variant
Private sFM3_R_O_operador As Variant
Private sFM3_observacoes As Variant
Private dFM3_dt_inclusao As Variant
Private dFM3_dt_alteracao As Variant
Private lFM4_pk_trabalho_fase_manut4 As Variant
Private lFM4_fk_trabalho_fase_manut4_alt_seq As Variant
Private lFM4_fk_trabalho As Variant
Private lFM4_fk_fase As Variant
Private sFM4_R_F_fase As Variant
Private lFM4_fk_operador As Variant
Private sFM4_R_O_operador As Variant
Private sFM4_observacoes As Variant
Private dFM4_dt_inclusao As Variant
Private dFM4_dt_alteracao As Variant
Private lFM5_pk_trabalho_fase_manut5 As Variant
Private lFM5_fk_trabalho_fase_manut5_alt_seq As Variant
Private lFM5_fk_trabalho As Variant
Private lFM5_fk_fase As Variant
Private sFM5_R_F_fase As Variant
Private lFM5_fk_operador As Variant
Private sFM5_R_O_operador As Variant
Private sFM5_observacoes As Variant
Private dFM5_dt_inclusao As Variant
Private dFM5_dt_alteracao As Variant
Private lFM6_pk_trabalho_fase_manut6 As Variant
Private lFM6_fk_trabalho_fase_manut6_alt_seq As Variant
Private lFM6_fk_trabalho As Variant
Private lFM6_fk_fase As Variant
Private sFM6_R_F_fase As Variant
Private lFM6_fk_operador As Variant
Private sFM6_R_O_operador As Variant
Private sFM6_observacoes As Variant
Private dFM6_dt_inclusao As Variant
Private dFM6_dt_alteracao As Variant
Private lFM7_pk_trabalho_fase_manut7 As Variant
Private lFM7_fk_trabalho_fase_manut7_alt_seq As Variant
Private lFM7_fk_trabalho As Variant
Private lFM7_fk_fase As Variant
Private sFM7_R_F_fase As Variant
Private lFM7_fk_operador As Variant
Private sFM7_R_O_operador As Variant
Private sFM7_observacoes As Variant
Private dFM7_dt_inclusao As Variant
Private dFM7_dt_alteracao As Variant
Private lFM8_pk_trabalho_fase_manut8 As Variant
Private lFM8_fk_trabalho_fase_manut8_alt_seq As Variant
Private lFM8_fk_trabalho As Variant
Private lFM8_fk_fase As Variant
Private sFM8_R_F_fase As Variant
Private lFM8_fk_operador As Variant
Private sFM8_R_O_operador As Variant
Private sFM8_observacoes As Variant
Private dFM8_dt_inclusao As Variant
Private dFM8_dt_alteracao As Variant
Private lFM9_pk_trabalho_fase_manut9 As Variant
Private lFM9_fk_trabalho_fase_manut9_alt_seq As Variant
Private lFM9_fk_trabalho As Variant
Private lFM9_fk_fase As Variant
Private sFM9_R_F_fase As Variant
Private lFM9_fk_operador As Variant
Private sFM9_R_O_operador As Variant
Private lFM9_fk_avaliacao_tipo As Variant
Private sFM9_R_AT_avaliacao_tipo As Variant
Private dFM9_dt_Avaliacao As Variant
Private dFM9_dt_envio_cliente As Variant
Private dFM9_dt_envio_padroes_CQ As Variant
Private sFM9_observacoes As Variant
Private dFM9_dt_inclusao As Variant
Private dFM9_dt_alteracao As Variant
Private lFM10_pk_trabalho_fase_manut10 As Variant
Private lFM10_fk_trabalho_fase_manut10_alt_seq As Variant
Private lFM10_fk_trabalho As Variant
Private lFM10_fk_fase As Variant
Private sFM10_R_F_fase As Variant
Private lFM10_fk_operador As Variant
Private sFM10_R_O_operador As Variant
Private lFM10_total_cores As Variant
Private sFM10_observacoes As Variant
Private dFM10_dt_inclusao As Variant
Private dFM10_dt_alteracao As Variant
Private lFM11_pk_trabalho_fase_manut11 As Variant
Private lFM11_fk_trabalho_fase_manut11_alt_seq As Variant
Private lFM11_fk_trabalho As Variant
Private lFM11_fk_fase As Variant
Private sFM11_R_F_fase As Variant
Private lFM11_fk_operador As Variant
Private sFM11_R_O_operador As Variant
Private sFM11_observacoes As Variant
Private dFM11_dt_inclusao As Variant
Private dFM11_dt_alteracao As Variant
Private lFM12_pk_trabalho_fase_manut12 As Variant
Private lFM12_fk_trabalho_fase_manut12_alt_seq As Variant
Private lFM12_fk_trabalho As Variant
Private lFM12_fk_fase As Variant
Private sFM12_R_F_fase As Variant
Private lFM12_fk_operador As Variant
Private sFM12_R_O_operador As Variant
Private lFM12_fk_avaliacao_tipo As Variant
Private sFM12_R_AT_avaliacao_tipo As Variant
Private dFM12_dt_Avaliacao As Variant
Private dFM12_dt_receb_rolinho As Variant
Private dFM12_dt_envio_laminacao As Variant
Private dFM12_dt_receb_laminacao As Variant
Private dFM12_dt_envio_padroes_cliente As Variant
Private dFM12_dt_envio_padroes_CQ As Variant
Private sFM12_observacoes As Variant
Private dFM12_dt_inclusao As Variant
Private dFM12_dt_alteracao As Variant
Private lGEN_pk_trabalho_fase_manut_gen As Variant
Private lGEN_fk_trabalho_fase_manut_gen_alt_seq As Variant
Private lGEN_fk_trabalho As Variant
Private lGEN_fk_fase As Variant
Private sGEN_R_F_fase As Variant
Private lGEN_fk_operador As Variant
Private sGEN_R_O_operador As Variant
Private dGEN_dt_liberacao_gravacao As Variant
Private dGEN_dt_inclusao As Variant
Private dGEN_dt_alteracao As Variant
Private lTFSO_pk_trabalho_fase_status_operador As Variant
Private lTFSO_fk_trabalho_fase_status_tipo_operador As Variant
Private sTFSO_tipo_status_operador As Variant
Private lTFSO_pk_fase_operador As Variant
Private lTFSO_codigo_fase_operador As Variant
Private sTFSO_nome_fase_operador As Variant
Private lTFSO_fk_operador_status As Variant
Private sTFSO_nome_operador_status As Variant
Private sTFSO_observacao_status As Variant
Private dTFSO_data_inclusao_status As Variant
Private lTFSA_pk_trabalho_fase_status_atual As Variant
Private lTFSA_fk_trabalho_fase_status_tipo_atual As Variant
Private sTFSA_tipo_status_atual As Variant
Private lTFSA_pk_fase_atual As Variant
Private lTFSA_codigo_fase_atual As Variant
Private sTFSA_nome_fase_atual As Variant

Private iAcaoFasesManut   As eAcaoFasesManut
Private sColunasConsulta  As String
Private sOrdemConsulta    As String
Private sCondicao         As String
Private cConexao          As Object
Private rsRegistros       As Object
Private iStatus           As ADRRBR_SIS_Basico.eStatus
Private sMensagem         As String

Private oBasico           As New ADRRBR_SIS_Basico.clsSIS_Basico
Private rsRegistrosTemp   As Object
Private sSQL              As String

Enum eAcaoFasesManut
    INC_Trab_Manut_1 = 1
    ALT_Trab_Manut_1 = 2
    INC_Manut_2 = 3
    ALT_Manut_2 = 4
    INC_Manut_3 = 5
    ALT_Manut_3 = 6
    INC_Manut_4 = 7
    ALT_Manut_4 = 8
    INC_Manut_5 = 9
    ALT_Manut_5 = 10
    INC_Manut_6 = 11
    ALT_Manut_6 = 12
    INC_Manut_7 = 13
    ALT_Manut_7 = 14
    INC_Manut_8 = 15
    ALT_Manut_8 = 16
    INC_Manut_9 = 17
    ALT_Manut_9 = 18
    INC_Manut_10 = 19
    ALT_Manut_10 = 20
    INC_Manut_11 = 21
    ALT_Manut_11 = 22
    INC_Manut_12 = 23
    ALT_Manut_12 = 24
    INC_Gener_DtLibGrav = 25
    ALT_Gener_DtLibGrav = 26
    ATU_Muda_Fase = 27
    CON_Trab_Fases_Manut = 28
    CON_Fase_Atual = 29
    CON_GER_Tipo_1 = 30
    CON_GER_Tipo_2 = 31
    CON_GER_Tipo_3 = 32
    CON_GER_Tipo_4 = 33
    CON_GER_Tipo_5 = 34
    CON_GER_Tipo_6 = 35
    CON_GER_Tipo_7 = 36
    CON_GER_Tipo_8 = 37
    CON_GER_Tipo_9 = 38
    CON_GER_Tipo_10 = 39
    CON_GER_Tipo_11 = 40
    CON_GER_Tipo_12 = 41
    CON_Resumo_Trab = 42
    CON_Status_Fase_Hist = 43
End Enum

Dim sProcedure As String
Dim sConsulta  As String

Private Const sObjetoClasse = "GRF_Trab_Fases"

Private Sub Class_Initialize()
    lTB_pk_trabalho = "NULL"
    lTB_fk_trabalho_alt_seq = "NULL"
    lTB_fk_trabalho_tipo = "NULL"
    sTB_R_TT_tipo_trabalho = "NULL"
    lTB_fk_produto = "NULL"
    sTB_R_P_produto = "NULL"
    lTB_fk_cliente = "NULL"
    sTB_R_C_cliente = "NULL"
    lTB_fk_representante = "NULL"
    sTB_R_R_representante = "NULL"
    lTB_fk_aprovacao_tipo = "NULL"
    sTB_R_AT_aprovacao_tipo = "NULL"
    lTB_fk_filial = "NULL"
    sTB_R_F_filial = "NULL"
    sTB_num_pedido = "NULL"
    sTB_num_pedido_antigo = "NULL"
    sTB_num_pedido_novo = "NULL"
    sTB_lead_time_programado = "NULL"
    dTB_dt_inclusao = "NULL"
    dTB_dt_alteracao = "NULL"
    dTB_dt_exclusao = "NULL"
    lFM1_pk_trabalho_fase_manut1 = "NULL"
    lFM1_fk_trabalho_fase_manut1_alt_seq = "NULL"
    lFM1_fk_trabalho = "NULL"
    lFM1_fk_fase = "NULL"
    sFM1_R_F_fase = "NULL"
    lFM1_fk_operador = "NULL"
    sFM1_R_O_operador = "NULL"
    lFM1_cores_alteradas = "NULL"
    lFM1_circ_cilindro = "NULL"
    sFM1_observacoes = "NULL"
    dFM1_dt_inclusao = "NULL"
    dFM1_dt_alteracao = "NULL"
    lFM2_pk_trabalho_fase_manut2 = "NULL"
    lFM2_fk_trabalho_fase_manut2_alt_seq = "NULL"
    lFM2_fk_trabalho = "NULL"
    lFM2_fk_fase = "NULL"
    sFM2_R_F_fase = "NULL"
    lFM2_fk_operador = "NULL"
    sFM2_R_O_operador = "NULL"
    sFM2_observacoes = "NULL"
    dFM2_dt_inclusao = "NULL"
    dFM2_dt_alteracao = "NULL"
    lFM3_pk_trabalho_fase_manut3 = "NULL"
    lFM3_fk_trabalho_fase_manut3_alt_seq = "NULL"
    lFM3_fk_trabalho = "NULL"
    lFM3_fk_fase = "NULL"
    sFM3_R_F_fase = "NULL"
    lFM3_fk_operador = "NULL"
    sFM3_R_O_operador = "NULL"
    sFM3_observacoes = "NULL"
    dFM3_dt_inclusao = "NULL"
    dFM3_dt_alteracao = "NULL"
    lFM4_pk_trabalho_fase_manut4 = "NULL"
    lFM4_fk_trabalho_fase_manut4_alt_seq = "NULL"
    lFM4_fk_trabalho = "NULL"
    lFM4_fk_fase = "NULL"
    sFM4_R_F_fase = "NULL"
    lFM4_fk_operador = "NULL"
    sFM4_R_O_operador = "NULL"
    sFM4_observacoes = "NULL"
    dFM4_dt_inclusao = "NULL"
    dFM4_dt_alteracao = "NULL"
    lFM5_pk_trabalho_fase_manut5 = "NULL"
    lFM5_fk_trabalho_fase_manut5_alt_seq = "NULL"
    lFM5_fk_trabalho = "NULL"
    lFM5_fk_fase = "NULL"
    sFM5_R_F_fase = "NULL"
    lFM5_fk_operador = "NULL"
    sFM5_R_O_operador = "NULL"
    sFM5_observacoes = "NULL"
    dFM5_dt_inclusao = "NULL"
    dFM5_dt_alteracao = "NULL"
    lFM6_pk_trabalho_fase_manut6 = "NULL"
    lFM6_fk_trabalho_fase_manut6_alt_seq = "NULL"
    lFM6_fk_trabalho = "NULL"
    lFM6_fk_fase = "NULL"
    sFM6_R_F_fase = "NULL"
    lFM6_fk_operador = "NULL"
    sFM6_R_O_operador = "NULL"
    sFM6_observacoes = "NULL"
    dFM6_dt_inclusao = "NULL"
    dFM6_dt_alteracao = "NULL"
    lFM7_pk_trabalho_fase_manut7 = "NULL"
    lFM7_fk_trabalho_fase_manut7_alt_seq = "NULL"
    lFM7_fk_trabalho = "NULL"
    lFM7_fk_fase = "NULL"
    sFM7_R_F_fase = "NULL"
    lFM7_fk_operador = "NULL"
    sFM7_R_O_operador = "NULL"
    sFM7_observacoes = "NULL"
    dFM7_dt_inclusao = "NULL"
    dFM7_dt_alteracao = "NULL"
    lFM8_pk_trabalho_fase_manut8 = "NULL"
    lFM8_fk_trabalho_fase_manut8_alt_seq = "NULL"
    lFM8_fk_trabalho = "NULL"
    lFM8_fk_fase = "NULL"
    sFM8_R_F_fase = "NULL"
    lFM8_fk_operador = "NULL"
    sFM8_R_O_operador = "NULL"
    sFM8_observacoes = "NULL"
    dFM8_dt_inclusao = "NULL"
    dFM8_dt_alteracao = "NULL"
    lFM9_pk_trabalho_fase_manut9 = "NULL"
    lFM9_fk_trabalho_fase_manut9_alt_seq = "NULL"
    lFM9_fk_trabalho = "NULL"
    lFM9_fk_fase = "NULL"
    sFM9_R_F_fase = "NULL"
    lFM9_fk_operador = "NULL"
    sFM9_R_O_operador = "NULL"
    lFM9_fk_avaliacao_tipo = "NULL"
    sFM9_R_AT_avaliacao_tipo = "NULL"
    dFM9_dt_Avaliacao = "NULL"
    dFM9_dt_envio_cliente = "NULL"
    dFM9_dt_envio_padroes_CQ = "NULL"
    sFM9_observacoes = "NULL"
    dFM9_dt_inclusao = "NULL"
    dFM9_dt_alteracao = "NULL"
    lFM10_pk_trabalho_fase_manut10 = "NULL"
    lFM10_fk_trabalho_fase_manut10_alt_seq = "NULL"
    lFM10_fk_trabalho = "NULL"
    lFM10_fk_fase = "NULL"
    sFM10_R_F_fase = "NULL"
    lFM10_fk_operador = "NULL"
    sFM10_R_O_operador = "NULL"
    lFM10_total_cores = "NULL"
    sFM10_observacoes = "NULL"
    dFM10_dt_inclusao = "NULL"
    dFM10_dt_alteracao = "NULL"
    lFM11_pk_trabalho_fase_manut11 = "NULL"
    lFM11_fk_trabalho_fase_manut11_alt_seq = "NULL"
    lFM11_fk_trabalho = "NULL"
    lFM11_fk_fase = "NULL"
    sFM11_R_F_fase = "NULL"
    lFM11_fk_operador = "NULL"
    sFM11_R_O_operador = "NULL"
    sFM11_observacoes = "NULL"
    dFM11_dt_inclusao = "NULL"
    dFM11_dt_alteracao = "NULL"
    lFM12_pk_trabalho_fase_manut12 = "NULL"
    lFM12_fk_trabalho_fase_manut12_alt_seq = "NULL"
    lFM12_fk_trabalho = "NULL"
    lFM12_fk_fase = "NULL"
    sFM12_R_F_fase = "NULL"
    lFM12_fk_operador = "NULL"
    sFM12_R_O_operador = "NULL"
    lFM12_fk_avaliacao_tipo = "NULL"
    sFM12_R_AT_avaliacao_tipo = "NULL"
    dFM12_dt_Avaliacao = "NULL"
    dFM12_dt_receb_rolinho = "NULL"
    dFM12_dt_envio_laminacao = "NULL"
    dFM12_dt_receb_laminacao = "NULL"
    dFM12_dt_envio_padroes_cliente = "NULL"
    dFM12_dt_envio_padroes_CQ = "NULL"
    sFM12_observacoes = "NULL"
    dFM12_dt_inclusao = "NULL"
    dFM12_dt_alteracao = "NULL"
    lGEN_pk_trabalho_fase_manut_gen = "NULL"
    lGEN_fk_trabalho_fase_manut_gen_alt_seq = "NULL"
    lGEN_fk_trabalho = "NULL"
    lGEN_fk_fase = "NULL"
    sGEN_R_F_fase = "NULL"
    lGEN_fk_operador = "NULL"
    sGEN_R_O_operador = "NULL"
    dGEN_dt_liberacao_gravacao = "NULL"
    dGEN_dt_inclusao = "NULL"
    dGEN_dt_alteracao = "NULL"
    lTFSO_pk_trabalho_fase_status_operador = "NULL"
    lTFSO_fk_trabalho_fase_status_tipo_operador = "NULL"
    sTFSO_tipo_status_operador = "NULL"
    lTFSO_pk_fase_operador = "NULL"
    lTFSO_codigo_fase_operador = "NULL"
    sTFSO_nome_fase_operador = "NULL"
    lTFSO_fk_operador_status = "NULL"
    sTFSO_nome_operador_status = "NULL"
    sTFSO_observacao_status = "NULL"
    dTFSO_data_inclusao_status = "NULL"
    lTFSA_pk_trabalho_fase_status_atual = "NULL"
    lTFSA_fk_trabalho_fase_status_tipo_atual = "NULL"
    sTFSA_tipo_status_atual = "NULL"
    lTFSA_pk_fase_atual = "NULL"
    lTFSA_codigo_fase_atual = "NULL"
    sTFSA_nome_fase_atual = "NULL"
    
    LimpaStatus
End Sub

Private Sub Class_Terminate()
    Set oBasico = Nothing
    Set cConexao = Nothing
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
End Sub

Public Property Let TB_pk_trabalho(ByVal vNewValue As Variant)
    lTB_pk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_pk_trabalho() As Variant
    TB_pk_trabalho = lTB_pk_trabalho
End Property

Public Property Let TB_fk_trabalho_alt_seq(ByVal vNewValue As Variant)
    lTB_fk_trabalho_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_trabalho_alt_seq() As Variant
    TB_fk_trabalho_alt_seq = lTB_fk_trabalho_alt_seq
End Property

Public Property Let TB_fk_trabalho_tipo(ByVal vNewValue As Variant)
    lTB_fk_trabalho_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_trabalho_tipo() As Variant
    TB_fk_trabalho_tipo = lTB_fk_trabalho_tipo
End Property

Public Property Let TB_R_TT_tipo_trabalho(ByVal vNewValue As Variant)
    sTB_R_TT_tipo_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_R_TT_tipo_trabalho() As Variant
    TB_R_TT_tipo_trabalho = sTB_R_TT_tipo_trabalho
End Property

Public Property Let TB_fk_produto(ByVal vNewValue As Variant)
    lTB_fk_produto = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_produto() As Variant
    TB_fk_produto = lTB_fk_produto
End Property

Public Property Let TB_R_P_produto(ByVal vNewValue As Variant)
    sTB_R_P_produto = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_R_P_produto() As Variant
    TB_R_P_produto = sTB_R_P_produto
End Property

Public Property Let TB_fk_cliente(ByVal vNewValue As Variant)
    lTB_fk_cliente = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_cliente() As Variant
    TB_fk_cliente = lTB_fk_cliente
End Property

Public Property Let TB_R_C_cliente(ByVal vNewValue As Variant)
    sTB_R_C_cliente = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_R_C_cliente() As Variant
    TB_R_C_cliente = sTB_R_C_cliente
End Property

Public Property Let TB_fk_representante(ByVal vNewValue As Variant)
    lTB_fk_representante = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_representante() As Variant
    TB_fk_representante = lTB_fk_representante
End Property

Public Property Let TB_R_R_representante(ByVal vNewValue As Variant)
    sTB_R_R_representante = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_R_R_representante() As Variant
    TB_R_R_representante = sTB_R_R_representante
End Property

Public Property Let TB_fk_aprovacao_tipo(ByVal vNewValue As Variant)
    lTB_fk_aprovacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_aprovacao_tipo() As Variant
    TB_fk_aprovacao_tipo = lTB_fk_aprovacao_tipo
End Property

Public Property Let TB_R_AT_aprovacao_tipo(ByVal vNewValue As Variant)
    sTB_R_AT_aprovacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_R_AT_aprovacao_tipo() As Variant
    TB_R_AT_aprovacao_tipo = sTB_R_AT_aprovacao_tipo
End Property

Public Property Let TB_fk_filial(ByVal vNewValue As Variant)
    lTB_fk_filial = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_fk_filial() As Variant
    TB_fk_filial = lTB_fk_filial
End Property

Public Property Let TB_R_F_filial(ByVal vNewValue As Variant)
    sTB_R_F_filial = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_R_F_filial() As Variant
    TB_R_F_filial = sTB_R_F_filial
End Property

Public Property Let TB_num_pedido(ByVal vNewValue As Variant)
    sTB_num_pedido = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_num_pedido() As Variant
    TB_num_pedido = sTB_num_pedido
End Property

Public Property Let TB_num_pedido_antigo(ByVal vNewValue As Variant)
    sTB_num_pedido_antigo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_num_pedido_antigo() As Variant
    TB_num_pedido_antigo = sTB_num_pedido_antigo
End Property

Public Property Let TB_num_pedido_novo(ByVal vNewValue As Variant)
    sTB_num_pedido_novo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_num_pedido_novo() As Variant
    TB_num_pedido_novo = sTB_num_pedido_novo
End Property

Public Property Let TB_lead_time_programado(ByVal vNewValue As Variant)
    sTB_lead_time_programado = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_lead_time_programado() As Variant
    TB_lead_time_programado = sTB_lead_time_programado
End Property

Public Property Let TB_dt_inclusao(ByVal vNewValue As Variant)
    dTB_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_dt_inclusao() As Variant
    TB_dt_inclusao = dTB_dt_inclusao
End Property

Public Property Let TB_dt_alteracao(ByVal vNewValue As Variant)
    dTB_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_dt_alteracao() As Variant
    TB_dt_alteracao = dTB_dt_alteracao
End Property

Public Property Let TB_dt_exclusao(ByVal vNewValue As Variant)
    dTB_dt_exclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TB_dt_exclusao() As Variant
    TB_dt_exclusao = dTB_dt_exclusao
End Property

Public Property Let FM1_pk_trabalho_fase_manut1(ByVal vNewValue As Variant)
    lFM1_pk_trabalho_fase_manut1 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_pk_trabalho_fase_manut1() As Variant
    FM1_pk_trabalho_fase_manut1 = lFM1_pk_trabalho_fase_manut1
End Property

Public Property Let FM1_fk_trabalho_fase_manut1_alt_seq(ByVal vNewValue As Variant)
    lFM1_fk_trabalho_fase_manut1_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_fk_trabalho_fase_manut1_alt_seq() As Variant
    FM1_fk_trabalho_fase_manut1_alt_seq = lFM1_fk_trabalho_fase_manut1_alt_seq
End Property

Public Property Let FM1_fk_trabalho(ByVal vNewValue As Variant)
    lFM1_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_fk_trabalho() As Variant
    FM1_fk_trabalho = lFM1_fk_trabalho
End Property

Public Property Let FM1_fk_fase(ByVal vNewValue As Variant)
    lFM1_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_fk_fase() As Variant
    FM1_fk_fase = lFM1_fk_fase
End Property

Public Property Let FM1_R_F_fase(ByVal vNewValue As Variant)
    sFM1_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_R_F_fase() As Variant
    FM1_R_F_fase = sFM1_R_F_fase
End Property

Public Property Let FM1_fk_operador(ByVal vNewValue As Variant)
    lFM1_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_fk_operador() As Variant
    FM1_fk_operador = lFM1_fk_operador
End Property

Public Property Let FM1_R_O_operador(ByVal vNewValue As Variant)
    sFM1_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_R_O_operador() As Variant
    FM1_R_O_operador = sFM1_R_O_operador
End Property

Public Property Let FM1_cores_alteradas(ByVal vNewValue As Variant)
    lFM1_cores_alteradas = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_cores_alteradas() As Variant
    FM1_cores_alteradas = lFM1_cores_alteradas
End Property

Public Property Let FM1_circ_cilindro(ByVal vNewValue As Variant)
    lFM1_circ_cilindro = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_circ_cilindro() As Variant
    FM1_circ_cilindro = lFM1_circ_cilindro
End Property

Public Property Let FM1_observacoes(ByVal vNewValue As Variant)
    sFM1_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_observacoes() As Variant
    FM1_observacoes = sFM1_observacoes
End Property

Public Property Let FM1_dt_inclusao(ByVal vNewValue As Variant)
    dFM1_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_dt_inclusao() As Variant
    FM1_dt_inclusao = dFM1_dt_inclusao
End Property

Public Property Let FM1_dt_alteracao(ByVal vNewValue As Variant)
    dFM1_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM1_dt_alteracao() As Variant
    FM1_dt_alteracao = dFM1_dt_alteracao
End Property

Public Property Let FM2_pk_trabalho_fase_manut2(ByVal vNewValue As Variant)
    lFM2_pk_trabalho_fase_manut2 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_pk_trabalho_fase_manut2() As Variant
    FM2_pk_trabalho_fase_manut2 = lFM2_pk_trabalho_fase_manut2
End Property

Public Property Let FM2_fk_trabalho_fase_manut2_alt_seq(ByVal vNewValue As Variant)
    lFM2_fk_trabalho_fase_manut2_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_fk_trabalho_fase_manut2_alt_seq() As Variant
    FM2_fk_trabalho_fase_manut2_alt_seq = lFM2_fk_trabalho_fase_manut2_alt_seq
End Property

Public Property Let FM2_fk_trabalho(ByVal vNewValue As Variant)
    lFM2_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_fk_trabalho() As Variant
    FM2_fk_trabalho = lFM2_fk_trabalho
End Property

Public Property Let FM2_fk_fase(ByVal vNewValue As Variant)
    lFM2_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_fk_fase() As Variant
    FM2_fk_fase = lFM2_fk_fase
End Property

Public Property Let FM2_R_F_fase(ByVal vNewValue As Variant)
    sFM2_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_R_F_fase() As Variant
    FM2_R_F_fase = sFM2_R_F_fase
End Property

Public Property Let FM2_fk_operador(ByVal vNewValue As Variant)
    lFM2_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_fk_operador() As Variant
    FM2_fk_operador = lFM2_fk_operador
End Property

Public Property Let FM2_R_O_operador(ByVal vNewValue As Variant)
    sFM2_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_R_O_operador() As Variant
    FM2_R_O_operador = sFM2_R_O_operador
End Property

Public Property Let FM2_observacoes(ByVal vNewValue As Variant)
    sFM2_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_observacoes() As Variant
    FM2_observacoes = sFM2_observacoes
End Property

Public Property Let FM2_dt_inclusao(ByVal vNewValue As Variant)
    dFM2_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_dt_inclusao() As Variant
    FM2_dt_inclusao = dFM2_dt_inclusao
End Property

Public Property Let FM2_dt_alteracao(ByVal vNewValue As Variant)
    dFM2_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM2_dt_alteracao() As Variant
    FM2_dt_alteracao = dFM2_dt_alteracao
End Property

Public Property Let FM3_pk_trabalho_fase_manut3(ByVal vNewValue As Variant)
    lFM3_pk_trabalho_fase_manut3 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_pk_trabalho_fase_manut3() As Variant
    FM3_pk_trabalho_fase_manut3 = lFM3_pk_trabalho_fase_manut3
End Property

Public Property Let FM3_fk_trabalho_fase_manut3_alt_seq(ByVal vNewValue As Variant)
    lFM3_fk_trabalho_fase_manut3_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_fk_trabalho_fase_manut3_alt_seq() As Variant
    FM3_fk_trabalho_fase_manut3_alt_seq = lFM3_fk_trabalho_fase_manut3_alt_seq
End Property

Public Property Let FM3_fk_trabalho(ByVal vNewValue As Variant)
    lFM3_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_fk_trabalho() As Variant
    FM3_fk_trabalho = lFM3_fk_trabalho
End Property

Public Property Let FM3_fk_fase(ByVal vNewValue As Variant)
    lFM3_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_fk_fase() As Variant
    FM3_fk_fase = lFM3_fk_fase
End Property

Public Property Let FM3_R_F_fase(ByVal vNewValue As Variant)
    sFM3_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_R_F_fase() As Variant
    FM3_R_F_fase = sFM3_R_F_fase
End Property

Public Property Let FM3_fk_operador(ByVal vNewValue As Variant)
    lFM3_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_fk_operador() As Variant
    FM3_fk_operador = lFM3_fk_operador
End Property

Public Property Let FM3_R_O_operador(ByVal vNewValue As Variant)
    sFM3_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_R_O_operador() As Variant
    FM3_R_O_operador = sFM3_R_O_operador
End Property

Public Property Let FM3_observacoes(ByVal vNewValue As Variant)
    sFM3_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_observacoes() As Variant
    FM3_observacoes = sFM3_observacoes
End Property

Public Property Let FM3_dt_inclusao(ByVal vNewValue As Variant)
    dFM3_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_dt_inclusao() As Variant
    FM3_dt_inclusao = dFM3_dt_inclusao
End Property

Public Property Let FM3_dt_alteracao(ByVal vNewValue As Variant)
    dFM3_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM3_dt_alteracao() As Variant
    FM3_dt_alteracao = dFM3_dt_alteracao
End Property

Public Property Let FM4_pk_trabalho_fase_manut4(ByVal vNewValue As Variant)
    lFM4_pk_trabalho_fase_manut4 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_pk_trabalho_fase_manut4() As Variant
    FM4_pk_trabalho_fase_manut4 = lFM4_pk_trabalho_fase_manut4
End Property

Public Property Let FM4_fk_trabalho_fase_manut4_alt_seq(ByVal vNewValue As Variant)
    lFM4_fk_trabalho_fase_manut4_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_fk_trabalho_fase_manut4_alt_seq() As Variant
    FM4_fk_trabalho_fase_manut4_alt_seq = lFM4_fk_trabalho_fase_manut4_alt_seq
End Property

Public Property Let FM4_fk_trabalho(ByVal vNewValue As Variant)
    lFM4_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_fk_trabalho() As Variant
    FM4_fk_trabalho = lFM4_fk_trabalho
End Property

Public Property Let FM4_fk_fase(ByVal vNewValue As Variant)
    lFM4_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_fk_fase() As Variant
    FM4_fk_fase = lFM4_fk_fase
End Property

Public Property Let FM4_R_F_fase(ByVal vNewValue As Variant)
    sFM4_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_R_F_fase() As Variant
    FM4_R_F_fase = sFM4_R_F_fase
End Property

Public Property Let FM4_fk_operador(ByVal vNewValue As Variant)
    lFM4_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_fk_operador() As Variant
    FM4_fk_operador = lFM4_fk_operador
End Property

Public Property Let FM4_R_O_operador(ByVal vNewValue As Variant)
    sFM4_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_R_O_operador() As Variant
    FM4_R_O_operador = sFM4_R_O_operador
End Property

Public Property Let FM4_observacoes(ByVal vNewValue As Variant)
    sFM4_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_observacoes() As Variant
    FM4_observacoes = sFM4_observacoes
End Property

Public Property Let FM4_dt_inclusao(ByVal vNewValue As Variant)
    dFM4_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_dt_inclusao() As Variant
    FM4_dt_inclusao = dFM4_dt_inclusao
End Property

Public Property Let FM4_dt_alteracao(ByVal vNewValue As Variant)
    dFM4_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM4_dt_alteracao() As Variant
    FM4_dt_alteracao = dFM4_dt_alteracao
End Property

Public Property Let FM5_pk_trabalho_fase_manut5(ByVal vNewValue As Variant)
    lFM5_pk_trabalho_fase_manut5 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_pk_trabalho_fase_manut5() As Variant
    FM5_pk_trabalho_fase_manut5 = lFM5_pk_trabalho_fase_manut5
End Property

Public Property Let FM5_fk_trabalho_fase_manut5_alt_seq(ByVal vNewValue As Variant)
    lFM5_fk_trabalho_fase_manut5_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_fk_trabalho_fase_manut5_alt_seq() As Variant
    FM5_fk_trabalho_fase_manut5_alt_seq = lFM5_fk_trabalho_fase_manut5_alt_seq
End Property

Public Property Let FM5_fk_trabalho(ByVal vNewValue As Variant)
    lFM5_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_fk_trabalho() As Variant
    FM5_fk_trabalho = lFM5_fk_trabalho
End Property

Public Property Let FM5_fk_fase(ByVal vNewValue As Variant)
    lFM5_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_fk_fase() As Variant
    FM5_fk_fase = lFM5_fk_fase
End Property

Public Property Let FM5_R_F_fase(ByVal vNewValue As Variant)
    sFM5_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_R_F_fase() As Variant
    FM5_R_F_fase = sFM5_R_F_fase
End Property

Public Property Let FM5_fk_operador(ByVal vNewValue As Variant)
    lFM5_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_fk_operador() As Variant
    FM5_fk_operador = lFM5_fk_operador
End Property

Public Property Let FM5_R_O_operador(ByVal vNewValue As Variant)
    sFM5_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_R_O_operador() As Variant
    FM5_R_O_operador = sFM5_R_O_operador
End Property

Public Property Let FM5_observacoes(ByVal vNewValue As Variant)
    sFM5_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_observacoes() As Variant
    FM5_observacoes = sFM5_observacoes
End Property

Public Property Let FM5_dt_inclusao(ByVal vNewValue As Variant)
    dFM5_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_dt_inclusao() As Variant
    FM5_dt_inclusao = dFM5_dt_inclusao
End Property

Public Property Let FM5_dt_alteracao(ByVal vNewValue As Variant)
    dFM5_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM5_dt_alteracao() As Variant
    FM5_dt_alteracao = dFM5_dt_alteracao
End Property

Public Property Let FM6_pk_trabalho_fase_manut6(ByVal vNewValue As Variant)
    lFM6_pk_trabalho_fase_manut6 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_pk_trabalho_fase_manut6() As Variant
    FM6_pk_trabalho_fase_manut6 = lFM6_pk_trabalho_fase_manut6
End Property

Public Property Let FM6_fk_trabalho_fase_manut6_alt_seq(ByVal vNewValue As Variant)
    lFM6_fk_trabalho_fase_manut6_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_fk_trabalho_fase_manut6_alt_seq() As Variant
    FM6_fk_trabalho_fase_manut6_alt_seq = lFM6_fk_trabalho_fase_manut6_alt_seq
End Property

Public Property Let FM6_fk_trabalho(ByVal vNewValue As Variant)
    lFM6_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_fk_trabalho() As Variant
    FM6_fk_trabalho = lFM6_fk_trabalho
End Property

Public Property Let FM6_fk_fase(ByVal vNewValue As Variant)
    lFM6_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_fk_fase() As Variant
    FM6_fk_fase = lFM6_fk_fase
End Property

Public Property Let FM6_R_F_fase(ByVal vNewValue As Variant)
    sFM6_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_R_F_fase() As Variant
    FM6_R_F_fase = sFM6_R_F_fase
End Property

Public Property Let FM6_fk_operador(ByVal vNewValue As Variant)
    lFM6_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_fk_operador() As Variant
    FM6_fk_operador = lFM6_fk_operador
End Property

Public Property Let FM6_R_O_operador(ByVal vNewValue As Variant)
    sFM6_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_R_O_operador() As Variant
    FM6_R_O_operador = sFM6_R_O_operador
End Property

Public Property Let FM6_observacoes(ByVal vNewValue As Variant)
    sFM6_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_observacoes() As Variant
    FM6_observacoes = sFM6_observacoes
End Property

Public Property Let FM6_dt_inclusao(ByVal vNewValue As Variant)
    dFM6_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_dt_inclusao() As Variant
    FM6_dt_inclusao = dFM6_dt_inclusao
End Property

Public Property Let FM6_dt_alteracao(ByVal vNewValue As Variant)
    dFM6_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM6_dt_alteracao() As Variant
    FM6_dt_alteracao = dFM6_dt_alteracao
End Property

Public Property Let FM7_pk_trabalho_fase_manut7(ByVal vNewValue As Variant)
    lFM7_pk_trabalho_fase_manut7 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_pk_trabalho_fase_manut7() As Variant
    FM7_pk_trabalho_fase_manut7 = lFM7_pk_trabalho_fase_manut7
End Property

Public Property Let FM7_fk_trabalho_fase_manut7_alt_seq(ByVal vNewValue As Variant)
    lFM7_fk_trabalho_fase_manut7_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_fk_trabalho_fase_manut7_alt_seq() As Variant
    FM7_fk_trabalho_fase_manut7_alt_seq = lFM7_fk_trabalho_fase_manut7_alt_seq
End Property

Public Property Let FM7_fk_trabalho(ByVal vNewValue As Variant)
    lFM7_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_fk_trabalho() As Variant
    FM7_fk_trabalho = lFM7_fk_trabalho
End Property

Public Property Let FM7_fk_fase(ByVal vNewValue As Variant)
    lFM7_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_fk_fase() As Variant
    FM7_fk_fase = lFM7_fk_fase
End Property

Public Property Let FM7_R_F_fase(ByVal vNewValue As Variant)
    sFM7_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_R_F_fase() As Variant
    FM7_R_F_fase = sFM7_R_F_fase
End Property

Public Property Let FM7_fk_operador(ByVal vNewValue As Variant)
    lFM7_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_fk_operador() As Variant
    FM7_fk_operador = lFM7_fk_operador
End Property

Public Property Let FM7_R_O_operador(ByVal vNewValue As Variant)
    sFM7_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_R_O_operador() As Variant
    FM7_R_O_operador = sFM7_R_O_operador
End Property

Public Property Let FM7_observacoes(ByVal vNewValue As Variant)
    sFM7_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_observacoes() As Variant
    FM7_observacoes = sFM7_observacoes
End Property

Public Property Let FM7_dt_inclusao(ByVal vNewValue As Variant)
    dFM7_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_dt_inclusao() As Variant
    FM7_dt_inclusao = dFM7_dt_inclusao
End Property

Public Property Let FM7_dt_alteracao(ByVal vNewValue As Variant)
    dFM7_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM7_dt_alteracao() As Variant
    FM7_dt_alteracao = dFM7_dt_alteracao
End Property

Public Property Let FM8_pk_trabalho_fase_manut8(ByVal vNewValue As Variant)
    lFM8_pk_trabalho_fase_manut8 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_pk_trabalho_fase_manut8() As Variant
    FM8_pk_trabalho_fase_manut8 = lFM8_pk_trabalho_fase_manut8
End Property

Public Property Let FM8_fk_trabalho_fase_manut8_alt_seq(ByVal vNewValue As Variant)
    lFM8_fk_trabalho_fase_manut8_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_fk_trabalho_fase_manut8_alt_seq() As Variant
    FM8_fk_trabalho_fase_manut8_alt_seq = lFM8_fk_trabalho_fase_manut8_alt_seq
End Property

Public Property Let FM8_fk_trabalho(ByVal vNewValue As Variant)
    lFM8_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_fk_trabalho() As Variant
    FM8_fk_trabalho = lFM8_fk_trabalho
End Property

Public Property Let FM8_fk_fase(ByVal vNewValue As Variant)
    lFM8_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_fk_fase() As Variant
    FM8_fk_fase = lFM8_fk_fase
End Property

Public Property Let FM8_R_F_fase(ByVal vNewValue As Variant)
    sFM8_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_R_F_fase() As Variant
    FM8_R_F_fase = sFM8_R_F_fase
End Property

Public Property Let FM8_fk_operador(ByVal vNewValue As Variant)
    lFM8_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_fk_operador() As Variant
    FM8_fk_operador = lFM8_fk_operador
End Property

Public Property Let FM8_R_O_operador(ByVal vNewValue As Variant)
    sFM8_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_R_O_operador() As Variant
    FM8_R_O_operador = sFM8_R_O_operador
End Property

Public Property Let FM8_observacoes(ByVal vNewValue As Variant)
    sFM8_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_observacoes() As Variant
    FM8_observacoes = sFM8_observacoes
End Property

Public Property Let FM8_dt_inclusao(ByVal vNewValue As Variant)
    dFM8_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_dt_inclusao() As Variant
    FM8_dt_inclusao = dFM8_dt_inclusao
End Property

Public Property Let FM8_dt_alteracao(ByVal vNewValue As Variant)
    dFM8_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM8_dt_alteracao() As Variant
    FM8_dt_alteracao = dFM8_dt_alteracao
End Property

Public Property Let FM9_pk_trabalho_fase_manut9(ByVal vNewValue As Variant)
    lFM9_pk_trabalho_fase_manut9 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_pk_trabalho_fase_manut9() As Variant
    FM9_pk_trabalho_fase_manut9 = lFM9_pk_trabalho_fase_manut9
End Property

Public Property Let FM9_fk_trabalho_fase_manut9_alt_seq(ByVal vNewValue As Variant)
    lFM9_fk_trabalho_fase_manut9_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_fk_trabalho_fase_manut9_alt_seq() As Variant
    FM9_fk_trabalho_fase_manut9_alt_seq = lFM9_fk_trabalho_fase_manut9_alt_seq
End Property

Public Property Let FM9_fk_trabalho(ByVal vNewValue As Variant)
    lFM9_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_fk_trabalho() As Variant
    FM9_fk_trabalho = lFM9_fk_trabalho
End Property

Public Property Let FM9_fk_fase(ByVal vNewValue As Variant)
    lFM9_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_fk_fase() As Variant
    FM9_fk_fase = lFM9_fk_fase
End Property

Public Property Let FM9_R_F_fase(ByVal vNewValue As Variant)
    sFM9_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_R_F_fase() As Variant
    FM9_R_F_fase = sFM9_R_F_fase
End Property

Public Property Let FM9_fk_operador(ByVal vNewValue As Variant)
    lFM9_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_fk_operador() As Variant
    FM9_fk_operador = lFM9_fk_operador
End Property

Public Property Let FM9_R_O_operador(ByVal vNewValue As Variant)
    sFM9_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_R_O_operador() As Variant
    FM9_R_O_operador = sFM9_R_O_operador
End Property

Public Property Let FM9_fk_avaliacao_tipo(ByVal vNewValue As Variant)
    lFM9_fk_avaliacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_fk_avaliacao_tipo() As Variant
    FM9_fk_avaliacao_tipo = lFM9_fk_avaliacao_tipo
End Property

Public Property Let FM9_R_AT_avaliacao_tipo(ByVal vNewValue As Variant)
    sFM9_R_AT_avaliacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_R_AT_avaliacao_tipo() As Variant
    FM9_R_AT_avaliacao_tipo = sFM9_R_AT_avaliacao_tipo
End Property

Public Property Let FM9_dt_avaliacao(ByVal vNewValue As Variant)
    dFM9_dt_Avaliacao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_dt_avaliacao() As Variant
    FM9_dt_avaliacao = dFM9_dt_Avaliacao
End Property

Public Property Let FM9_dt_envio_cliente(ByVal vNewValue As Variant)
    dFM9_dt_envio_cliente = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_dt_envio_cliente() As Variant
    FM9_dt_envio_cliente = dFM9_dt_envio_cliente
End Property

Public Property Let FM9_dt_envio_padroes_CQ(ByVal vNewValue As Variant)
    dFM9_dt_envio_padroes_CQ = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_dt_envio_padroes_CQ() As Variant
    FM9_dt_envio_padroes_CQ = dFM9_dt_envio_padroes_CQ
End Property

Public Property Let FM9_observacoes(ByVal vNewValue As Variant)
    sFM9_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_observacoes() As Variant
    FM9_observacoes = sFM9_observacoes
End Property

Public Property Let FM9_dt_inclusao(ByVal vNewValue As Variant)
    dFM9_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_dt_inclusao() As Variant
    FM9_dt_inclusao = dFM9_dt_inclusao
End Property

Public Property Let FM9_dt_alteracao(ByVal vNewValue As Variant)
    dFM9_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM9_dt_alteracao() As Variant
    FM9_dt_alteracao = dFM9_dt_alteracao
End Property

Public Property Let FM10_pk_trabalho_fase_manut10(ByVal vNewValue As Variant)
    lFM10_pk_trabalho_fase_manut10 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_pk_trabalho_fase_manut10() As Variant
    FM10_pk_trabalho_fase_manut10 = lFM10_pk_trabalho_fase_manut10
End Property

Public Property Let FM10_fk_trabalho_fase_manut10_alt_seq(ByVal vNewValue As Variant)
    lFM10_fk_trabalho_fase_manut10_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_fk_trabalho_fase_manut10_alt_seq() As Variant
    FM10_fk_trabalho_fase_manut10_alt_seq = lFM10_fk_trabalho_fase_manut10_alt_seq
End Property

Public Property Let FM10_fk_trabalho(ByVal vNewValue As Variant)
    lFM10_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_fk_trabalho() As Variant
    FM10_fk_trabalho = lFM10_fk_trabalho
End Property

Public Property Let FM10_fk_fase(ByVal vNewValue As Variant)
    lFM10_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_fk_fase() As Variant
    FM10_fk_fase = lFM10_fk_fase
End Property

Public Property Let FM10_R_F_fase(ByVal vNewValue As Variant)
    sFM10_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_R_F_fase() As Variant
    FM10_R_F_fase = sFM10_R_F_fase
End Property

Public Property Let FM10_fk_operador(ByVal vNewValue As Variant)
    lFM10_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_fk_operador() As Variant
    FM10_fk_operador = lFM10_fk_operador
End Property

Public Property Let FM10_R_O_operador(ByVal vNewValue As Variant)
    sFM10_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_R_O_operador() As Variant
    FM10_R_O_operador = sFM10_R_O_operador
End Property

Public Property Let FM10_total_cores(ByVal vNewValue As Variant)
    lFM10_total_cores = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_total_cores() As Variant
    FM10_total_cores = lFM10_total_cores
End Property

Public Property Let FM10_observacoes(ByVal vNewValue As Variant)
    sFM10_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_observacoes() As Variant
    FM10_observacoes = sFM10_observacoes
End Property

Public Property Let FM10_dt_inclusao(ByVal vNewValue As Variant)
    dFM10_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_dt_inclusao() As Variant
    FM10_dt_inclusao = dFM10_dt_inclusao
End Property

Public Property Let FM10_dt_alteracao(ByVal vNewValue As Variant)
    dFM10_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM10_dt_alteracao() As Variant
    FM10_dt_alteracao = dFM10_dt_alteracao
End Property

Public Property Let FM11_pk_trabalho_fase_manut11(ByVal vNewValue As Variant)
    lFM11_pk_trabalho_fase_manut11 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_pk_trabalho_fase_manut11() As Variant
    FM11_pk_trabalho_fase_manut11 = lFM11_pk_trabalho_fase_manut11
End Property

Public Property Let FM11_fk_trabalho_fase_manut11_alt_seq(ByVal vNewValue As Variant)
    lFM11_fk_trabalho_fase_manut11_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_fk_trabalho_fase_manut11_alt_seq() As Variant
    FM11_fk_trabalho_fase_manut11_alt_seq = lFM11_fk_trabalho_fase_manut11_alt_seq
End Property

Public Property Let FM11_fk_trabalho(ByVal vNewValue As Variant)
    lFM11_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_fk_trabalho() As Variant
    FM11_fk_trabalho = lFM11_fk_trabalho
End Property

Public Property Let FM11_fk_fase(ByVal vNewValue As Variant)
    lFM11_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_fk_fase() As Variant
    FM11_fk_fase = lFM11_fk_fase
End Property

Public Property Let FM11_R_F_fase(ByVal vNewValue As Variant)
    sFM11_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_R_F_fase() As Variant
    FM11_R_F_fase = sFM11_R_F_fase
End Property

Public Property Let FM11_fk_operador(ByVal vNewValue As Variant)
    lFM11_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_fk_operador() As Variant
    FM11_fk_operador = lFM11_fk_operador
End Property

Public Property Let FM11_R_O_operador(ByVal vNewValue As Variant)
    sFM11_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_R_O_operador() As Variant
    FM11_R_O_operador = sFM11_R_O_operador
End Property

Public Property Let FM11_observacoes(ByVal vNewValue As Variant)
    sFM11_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_observacoes() As Variant
    FM11_observacoes = sFM11_observacoes
End Property

Public Property Let FM11_dt_inclusao(ByVal vNewValue As Variant)
    dFM11_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_dt_inclusao() As Variant
    FM11_dt_inclusao = dFM11_dt_inclusao
End Property

Public Property Let FM11_dt_alteracao(ByVal vNewValue As Variant)
    dFM11_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM11_dt_alteracao() As Variant
    FM11_dt_alteracao = dFM11_dt_alteracao
End Property

Public Property Let FM12_pk_trabalho_fase_manut12(ByVal vNewValue As Variant)
    lFM12_pk_trabalho_fase_manut12 = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_pk_trabalho_fase_manut12() As Variant
    FM12_pk_trabalho_fase_manut12 = lFM12_pk_trabalho_fase_manut12
End Property

Public Property Let FM12_fk_trabalho_fase_manut12_alt_seq(ByVal vNewValue As Variant)
    lFM12_fk_trabalho_fase_manut12_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_fk_trabalho_fase_manut12_alt_seq() As Variant
    FM12_fk_trabalho_fase_manut12_alt_seq = lFM12_fk_trabalho_fase_manut12_alt_seq
End Property

Public Property Let FM12_fk_trabalho(ByVal vNewValue As Variant)
    lFM12_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_fk_trabalho() As Variant
    FM12_fk_trabalho = lFM12_fk_trabalho
End Property

Public Property Let FM12_fk_fase(ByVal vNewValue As Variant)
    lFM12_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_fk_fase() As Variant
    FM12_fk_fase = lFM12_fk_fase
End Property

Public Property Let FM12_R_F_fase(ByVal vNewValue As Variant)
    sFM12_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_R_F_fase() As Variant
    FM12_R_F_fase = sFM12_R_F_fase
End Property

Public Property Let FM12_fk_operador(ByVal vNewValue As Variant)
    lFM12_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_fk_operador() As Variant
    FM12_fk_operador = lFM12_fk_operador
End Property

Public Property Let FM12_R_O_operador(ByVal vNewValue As Variant)
    sFM12_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_R_O_operador() As Variant
    FM12_R_O_operador = sFM12_R_O_operador
End Property

Public Property Let FM12_fk_avaliacao_tipo(ByVal vNewValue As Variant)
    lFM12_fk_avaliacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_fk_avaliacao_tipo() As Variant
    FM12_fk_avaliacao_tipo = lFM12_fk_avaliacao_tipo
End Property

Public Property Let FM12_R_AT_avaliacao_tipo(ByVal vNewValue As Variant)
    sFM12_R_AT_avaliacao_tipo = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_R_AT_avaliacao_tipo_R_F_fase() As Variant
    FM12_R_AT_avaliacao_tipo = sFM12_R_AT_avaliacao_tipo
End Property

Public Property Let FM12_dt_avaliacao(ByVal vNewValue As Variant)
    dFM12_dt_Avaliacao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_avaliacao() As Variant
    FM12_dt_avaliacao = dFM12_dt_Avaliacao
End Property

Public Property Let FM12_dt_receb_rolinho(ByVal vNewValue As Variant)
    dFM12_dt_receb_rolinho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_receb_rolinho() As Variant
    FM12_dt_receb_rolinho = dFM12_dt_receb_rolinho
End Property

Public Property Let FM12_dt_envio_laminacao(ByVal vNewValue As Variant)
    dFM12_dt_envio_laminacao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_envio_laminacao() As Variant
    FM12_dt_envio_laminacao = dFM12_dt_envio_laminacao
End Property

Public Property Let FM12_dt_receb_laminacao(ByVal vNewValue As Variant)
    dFM12_dt_receb_laminacao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_receb_laminacao() As Variant
    FM12_dt_receb_laminacao = dFM12_dt_receb_laminacao
End Property

Public Property Let FM12_dt_envio_padroes_cliente(ByVal vNewValue As Variant)
    dFM12_dt_envio_padroes_cliente = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_envio_padroes_cliente() As Variant
    FM12_dt_envio_padroes_cliente = dFM12_dt_envio_padroes_cliente
End Property

Public Property Let FM12_dt_envio_padroes_CQ(ByVal vNewValue As Variant)
    dFM12_dt_envio_padroes_CQ = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_envio_padroes_CQ() As Variant
    FM12_dt_envio_padroes_CQ = dFM12_dt_envio_padroes_CQ
End Property

Public Property Let FM12_observacoes(ByVal vNewValue As Variant)
    sFM12_observacoes = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_observacoes() As Variant
    FM12_observacoes = sFM12_observacoes
End Property

Public Property Let FM12_dt_inclusao(ByVal vNewValue As Variant)
    dFM12_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_inclusao() As Variant
    FM12_dt_inclusao = dFM12_dt_inclusao
End Property

Public Property Let FM12_dt_alteracao(ByVal vNewValue As Variant)
    dFM12_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get FM12_dt_alteracao() As Variant
    FM12_dt_alteracao = dFM12_dt_alteracao
End Property

Public Property Let GEN_pk_trabalho_fase_manut_gen(ByVal vNewValue As Variant)
    lGEN_pk_trabalho_fase_manut_gen = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_pk_trabalho_fase_manut_gen() As Variant
    GEN_pk_trabalho_fase_manut_gen = lGEN_pk_trabalho_fase_manut_gen
End Property

Public Property Let GEN_fk_trabalho_fase_manut_gen_alt_seq(ByVal vNewValue As Variant)
    lGEN_fk_trabalho_fase_manut_gen_alt_seq = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_fk_trabalho_fase_manut_gen_alt_seq() As Variant
    GEN_fk_trabalho_fase_manut_gen_alt_seq = lGEN_fk_trabalho_fase_manut_gen_alt_seq
End Property

Public Property Let GEN_fk_trabalho(ByVal vNewValue As Variant)
    lGEN_fk_trabalho = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_fk_trabalho() As Variant
    GEN_fk_trabalho = lGEN_fk_trabalho
End Property

Public Property Let GEN_fk_fase(ByVal vNewValue As Variant)
    lGEN_fk_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_fk_fase() As Variant
    GEN_fk_fase = lGEN_fk_fase
End Property

Public Property Let GEN_R_F_fase(ByVal vNewValue As Variant)
    sGEN_R_F_fase = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_R_F_fase() As Variant
    GEN_R_F_fase = sGEN_R_F_fase
End Property

Public Property Let GEN_fk_operador(ByVal vNewValue As Variant)
    lGEN_fk_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_fk_operador() As Variant
    GEN_fk_operador = lGEN_fk_operador
End Property

Public Property Let GEN_R_O_operador(ByVal vNewValue As Variant)
    sGEN_R_O_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_R_O_operador() As Variant
    GEN_R_O_operador = sGEN_R_O_operador
End Property

Public Property Let GEN_dt_liberacao_gravacao(ByVal vNewValue As Variant)
    dGEN_dt_liberacao_gravacao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_dt_liberacao_gravacao() As Variant
    GEN_dt_liberacao_gravacao = dGEN_dt_liberacao_gravacao
End Property

Public Property Let GEN_dt_inclusao(ByVal vNewValue As Variant)
    dGEN_dt_inclusao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_dt_inclusao() As Variant
    GEN_dt_inclusao = dGEN_dt_inclusao
End Property

Public Property Let GEN_dt_alteracao(ByVal vNewValue As Variant)
    dGEN_dt_alteracao = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get GEN_dt_alteracao() As Variant
    GEN_dt_alteracao = dGEN_dt_alteracao
End Property

Public Property Let TFSO_pk_trabalho_fase_status_operador(ByVal vNewValue As Variant)
    lTFSO_pk_trabalho_fase_status_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_pk_trabalho_fase_status_operador() As Variant
    TFSO_pk_trabalho_fase_status_operador = lTFSO_pk_trabalho_fase_status_operador
End Property

Public Property Let TFSO_fk_trabalho_fase_status_tipo_operador(ByVal vNewValue As Variant)
    lTFSO_fk_trabalho_fase_status_tipo_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_fk_trabalho_fase_status_tipo_operador() As Variant
    TFSO_fk_trabalho_fase_status_tipo_operador = lTFSO_fk_trabalho_fase_status_tipo_operador
End Property

Public Property Let TFSO_tipo_status_operador(ByVal vNewValue As Variant)
    sTFSO_tipo_status_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_tipo_status_operador() As Variant
    TFSO_tipo_status_operador = sTFSO_tipo_status_operador
End Property

Public Property Let TFSO_pk_fase_operador(ByVal vNewValue As Variant)
    lTFSO_pk_fase_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_pk_fase_operador() As Variant
    TFSO_pk_fase_operador = lTFSO_pk_fase_operador
End Property

Public Property Let TFSO_codigo_fase_operador(ByVal vNewValue As Variant)
    lTFSO_codigo_fase_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_codigo_fase_operador() As Variant
    TFSO_codigo_fase_operador = lTFSO_codigo_fase_operador
End Property

Public Property Let TFSO_nome_fase_operador(ByVal vNewValue As Variant)
    sTFSO_nome_fase_operador = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_nome_fase_operador() As Variant
    TFSO_nome_fase_operador = sTFSO_nome_fase_operador
End Property

Public Property Let TFSO_fk_operador_status(ByVal vNewValue As Variant)
    lTFSO_fk_operador_status = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_fk_operador_status() As Variant
    TFSO_fk_operador_status = lTFSO_fk_operador_status
End Property

Public Property Let TFSO_nome_operador_status(ByVal vNewValue As Variant)
    sTFSO_nome_operador_status = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_nome_operador_status() As Variant
    TFSO_nome_operador_status = sTFSO_nome_operador_status
End Property

Public Property Let TFSO_observacao_status(ByVal vNewValue As Variant)
    sTFSO_observacao_status = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_observacao_status() As Variant
    TFSO_observacao_status = sTFSO_observacao_status
End Property

Public Property Let TFSO_data_inclusao_status(ByVal vNewValue As Variant)
    dTFSO_data_inclusao_status = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSO_data_inclusao_status() As Variant
    TFSO_data_inclusao_status = dTFSO_data_inclusao_status
End Property

Public Property Let TFSA_pk_trabalho_fase_status_atual(ByVal vNewValue As Variant)
    lTFSA_pk_trabalho_fase_status_atual = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSA_pk_trabalho_fase_status_atual() As Variant
    TFSA_pk_trabalho_fase_status_atual = lTFSA_pk_trabalho_fase_status_atual
End Property

Public Property Let TFSA_fk_trabalho_fase_status_tipo_atual(ByVal vNewValue As Variant)
    lTFSA_fk_trabalho_fase_status_tipo_atual = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSA_fk_trabalho_fase_status_tipo_atual() As Variant
    TFSA_fk_trabalho_fase_status_tipo_atual = lTFSA_fk_trabalho_fase_status_tipo_atual
End Property

Public Property Let TFSA_tipo_status_atual(ByVal vNewValue As Variant)
    sTFSA_tipo_status_atual = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSA_tipo_status_atual() As Variant
    TFSA_tipo_status_atual = sTFSA_tipo_status_atual
End Property

Public Property Let TFSA_pk_fase_atual(ByVal vNewValue As Variant)
    lTFSA_pk_fase_atual = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSA_pk_fase_atual() As Variant
    TFSA_pk_fase_atual = lTFSA_pk_fase_atual
End Property

Public Property Let TFSA_codigo_fase_atual(ByVal vNewValue As Variant)
    lTFSA_codigo_fase_atual = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSA_codigo_fase_atual() As Variant
    TFSA_codigo_fase_atual = lTFSA_codigo_fase_atual
End Property

Public Property Let TFSA_nome_fase_atual(ByVal vNewValue As Variant)
    sTFSA_nome_fase_atual = oBasico.Geral.TrocaNuLL(vNewValue, "NULL")
End Property
Public Property Get TFSA_nome_fase_atual() As Variant
    TFSA_nome_fase_atual = sTFSA_nome_fase_atual
End Property

Public Property Let AcaoFasesManut(ByVal vNewValue As eAcaoFasesManut)
    iAcaoFasesManut = vNewValue
End Property
Public Property Get AcaoFasesManut() As eAcaoFasesManut
    AcaoFasesManut = iAcaoFasesManut
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

Public Sub Atualizar_BD()
    On Error GoTo ErroRotina

    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    If Not VerificaPropriedadesAtualizacao Then Exit Sub
    
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            iStatus = Atualiza_SQL_Server
    End Select
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sObjetoClasse & "." & "Atualizar_BD > "
End Sub

Public Sub Consultar_BD()
    On Error GoTo ErroRotina

    LimpaStatus
    
    If Not VerificaPropriedadesGerais Then Exit Sub
    If Not VerificaPropriedadesConsulta Then Exit Sub
    
    Select Case cConexao.TipoBancoDados
        Case SQL_Server
            iStatus = Consulta_SQL_Server
    End Select
    
    If iStatus = Encontrado Then
        Select Case iAcaoFasesManut
            Case CON_Resumo_Trab
                frmResumoTrab.TituloResumo = "Resumo do Trabalho"
                Set frmResumoTrab.RegistroResumo = rsRegistros
                frmResumoTrab.Show vbModal
                Exit Sub
                
            Case CON_Status_Fase_Hist
                frmHistorico.TituloHist = "Histórico de Status das Fases do Trabalho"
                Set frmHistorico.RegistrosHist = rsRegistros
                frmHistorico.Show vbModal
                Exit Sub
        End Select
    End If
    
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

Private Function VerificaPropriedadesAtualizacao() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesAtualizacao = True
    
    If iAcaoFasesManut <> INC_Trab_Manut_1 And _
       iAcaoFasesManut <> ALT_Trab_Manut_1 And _
       iAcaoFasesManut <> INC_Manut_2 And _
       iAcaoFasesManut <> ALT_Manut_2 And _
       iAcaoFasesManut <> INC_Manut_3 And _
       iAcaoFasesManut <> ALT_Manut_3 And _
       iAcaoFasesManut <> INC_Manut_4 And _
       iAcaoFasesManut <> ALT_Manut_4 And _
       iAcaoFasesManut <> INC_Manut_5 And _
       iAcaoFasesManut <> ALT_Manut_5 And _
       iAcaoFasesManut <> INC_Manut_6 And _
       iAcaoFasesManut <> ALT_Manut_6 And _
       iAcaoFasesManut <> INC_Manut_7 And _
       iAcaoFasesManut <> ALT_Manut_7 And _
       iAcaoFasesManut <> INC_Manut_8 And _
       iAcaoFasesManut <> ALT_Manut_8 And _
       iAcaoFasesManut <> INC_Manut_9 And _
       iAcaoFasesManut <> ALT_Manut_9 And _
       iAcaoFasesManut <> INC_Manut_10 And _
       iAcaoFasesManut <> ALT_Manut_10 And _
       iAcaoFasesManut <> INC_Manut_11 And _
       iAcaoFasesManut <> ALT_Manut_11 And _
       iAcaoFasesManut <> INC_Manut_12 And _
       iAcaoFasesManut <> ALT_Manut_12 Then
        
       If iAcaoFasesManut <> INC_Gener_DtLibGrav And _
          iAcaoFasesManut <> ALT_Gener_DtLibGrav And _
          iAcaoFasesManut <> ATU_Muda_Fase Then
           
           iStatus = Erro
           sMensagem = "Informe a AÇÃO para atualização dos dados!"
           VerificaPropriedadesAtualizacao = False
           Exit Function
        End If
    End If
    
    Exit Function
        
ErroRotina:
    iStatus = Erro
    VerificaPropriedadesAtualizacao = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sObjetoClasse & "." & "VerificaPropriedadesAtualizacao > "
End Function

Private Function VerificaPropriedadesConsulta() As Boolean
    On Error GoTo ErroRotina
    
    VerificaPropriedadesConsulta = True
    
    If iAcaoFasesManut <> CON_Trab_Fases_Manut And _
       iAcaoFasesManut <> CON_Fase_Atual And _
       iAcaoFasesManut <> CON_GER_Tipo_1 And _
       iAcaoFasesManut <> CON_GER_Tipo_2 And _
       iAcaoFasesManut <> CON_GER_Tipo_3 And _
       iAcaoFasesManut <> CON_GER_Tipo_4 And _
       iAcaoFasesManut <> CON_GER_Tipo_5 And _
       iAcaoFasesManut <> CON_GER_Tipo_6 And _
       iAcaoFasesManut <> CON_GER_Tipo_7 And _
       iAcaoFasesManut <> CON_GER_Tipo_8 And _
       iAcaoFasesManut <> CON_GER_Tipo_9 And _
       iAcaoFasesManut <> CON_GER_Tipo_10 And _
       iAcaoFasesManut <> CON_GER_Tipo_11 And _
       iAcaoFasesManut <> CON_GER_Tipo_12 And _
       iAcaoFasesManut <> CON_Resumo_Trab And _
       iAcaoFasesManut <> CON_Status_Fase_Hist Then
       
        iStatus = Erro
        sMensagem = "Informe a AÇÃO para consulta dos dados!"
        VerificaPropriedadesConsulta = False
        Exit Function
    End If
    
    If iAcaoFasesManut = CON_Fase_Atual Then
        If lTB_pk_trabalho = "NULL" Then
            iStatus = Erro
            sMensagem = "Para consultar a fase atual de um trabalho, informe obrigatoriamente a propriedade < TB_pk_trabalho > !"
            VerificaPropriedadesConsulta = False
            Exit Function
        End If
        If Trim(sColunasConsulta) <> Empty Then
            iStatus = Erro
            sMensagem = "Para consultar a fase atual de um trabalho, não informar a propriedade < ColunasConsulta > !"
            VerificaPropriedadesConsulta = False
            Exit Function
        End If
        If Trim(sOrdemConsulta) <> Empty Then
            iStatus = Erro
            sMensagem = "Para consultar a fase atual de um trabalho, não informar a propriedade < OrdemConsulta > !"
            VerificaPropriedadesConsulta = False
            Exit Function
        End If
        If Trim(sCondicao) <> Empty Then
            iStatus = Erro
            sMensagem = "Para consultar a fase atual de um trabalho, não informar a propriedade < Condicao > !"
            VerificaPropriedadesConsulta = False
            Exit Function
        End If
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

Private Function Atualiza_SQL_Server() As eStatus
    On Error GoTo ErroRotina
    
    Atualiza_SQL_Server = Erro
    
    Select Case iAcaoFasesManut
        Case INC_Trab_Manut_1
            'GRF_Trabalhos / GRF_Trabalhos_Fases_Manut_1 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_1"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho_tipo     = " & lTB_fk_trabalho_tipo & vbCrLf
            sSQL = sSQL & ",@fk_aprovacao_tipo    = " & lTB_fk_aprovacao_tipo & vbCrLf
            sSQL = sSQL & ",@fk_cliente           = " & lTB_fk_cliente & vbCrLf
            sSQL = sSQL & ",@fk_filial            = " & lTB_fk_filial & vbCrLf
            sSQL = sSQL & ",@num_pedido           = " & oBasico.Geral.Aspas(CStr(sTB_num_pedido), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@num_pedido_antigo    = " & oBasico.Geral.Aspas(CStr(sTB_num_pedido_antigo), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@num_pedido_novo      = " & oBasico.Geral.Aspas(CStr(sTB_num_pedido_novo), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@lead_time_programado = " & oBasico.Geral.Aspas(CStr(sTB_lead_time_programado), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@produto_nome         = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sTB_R_P_produto)), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@representante_nome   = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sTB_R_R_representante)), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@fk_fase              = " & lFM1_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador          = " & lFM1_fk_operador & vbCrLf
            sSQL = sSQL & ",@cores_alteradas      = " & lFM1_cores_alteradas & vbCrLf
            sSQL = sSQL & ",@circ_cilindro        = " & lFM1_circ_cilindro & vbCrLf
            sSQL = sSQL & ",@observacoes          = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM1_observacoes)), AspasSimples) & vbCrLf
            
        Case ALT_Trab_Manut_1
            'GRF_Trabalhos / GRF_Trabalhos_Fases_Manut_1 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_2"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho             = " & lTB_pk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_trabalho_tipo        = " & lTB_fk_trabalho_tipo & vbCrLf
            sSQL = sSQL & ",@fk_aprovacao_tipo       = " & lTB_fk_aprovacao_tipo & vbCrLf
            sSQL = sSQL & ",@fk_cliente              = " & lTB_fk_cliente & vbCrLf
            sSQL = sSQL & ",@fk_filial               = " & lTB_fk_filial & vbCrLf
            sSQL = sSQL & ",@num_pedido_antigo       = " & oBasico.Geral.Aspas(CStr(sTB_num_pedido_antigo), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@num_pedido_novo         = " & oBasico.Geral.Aspas(CStr(sTB_num_pedido_novo), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@lead_time_programado    = " & oBasico.Geral.Aspas(CStr(sTB_lead_time_programado), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@produto_nome            = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sTB_R_P_produto)), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@representante_nome      = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sTB_R_R_representante)), AspasSimples) & vbCrLf
            sSQL = sSQL & ",@pk_trabalho_fase_manut1 = " & lFM1_pk_trabalho_fase_manut1 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM1_fk_operador & vbCrLf
            sSQL = sSQL & ",@cores_alteradas         = " & lFM1_cores_alteradas & vbCrLf
            sSQL = sSQL & ",@circ_cilindro           = " & lFM1_circ_cilindro & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM1_observacoes)), AspasSimples) & vbCrLf

        Case INC_Manut_2
            'GRF_Trabalhos_Fases_Manut_2 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_3"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM2_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM2_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM2_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM2_observacoes)), AspasSimples) & vbCrLf

        Case ALT_Manut_2
            'GRF_Trabalhos_Fases_Manut_2 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_3"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut2 = " & lFM2_pk_trabalho_fase_manut2 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM2_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM2_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_3
            'GRF_Trabalhos_Fases_Manut_3 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_4"
        
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM3_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM3_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM3_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM3_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_3
            'GRF_Trabalhos_Fases_Manut_3 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_4"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut3 = " & lFM3_pk_trabalho_fase_manut3 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM3_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM3_observacoes)), AspasSimples) & vbCrLf
    
        Case INC_Manut_4
            'GRF_Trabalhos_Fases_Manut_4 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_5"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM4_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM4_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM4_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM4_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_4
            'GRF_Trabalhos_Fases_Manut_4 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_5"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut4 = " & lFM4_pk_trabalho_fase_manut4 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM4_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM4_observacoes)), AspasSimples) & vbCrLf
            
        Case INC_Manut_5
            'GRF_Trabalhos_Fases_Manut_5 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_6"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM5_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM5_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM5_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM5_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_5
            'GRF_Trabalhos_Fases_Manut_5 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_6"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut5 = " & lFM5_pk_trabalho_fase_manut5 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM5_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM5_observacoes)), AspasSimples) & vbCrLf
    
        Case INC_Manut_6
            'GRF_Trabalhos_Fases_Manut_6 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_7"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM6_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM6_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM6_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM6_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_6
            'GRF_Trabalhos_Fases_Manut_6 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_7"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut6 = " & lFM6_pk_trabalho_fase_manut6 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM6_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM6_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_7
            'GRF_Trabalhos_Fases_Manut_7 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_8"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM7_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM7_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM7_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM7_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_7
            'GRF_Trabalhos_Fases_Manut_7 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_8"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut7 = " & lFM7_pk_trabalho_fase_manut7 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM7_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM7_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_8
            'GRF_Trabalhos_Fases_Manut_8 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_9"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM8_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM8_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM8_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM8_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_8
            'GRF_Trabalhos_Fases_Manut_8 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_9"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut8 = " & lFM8_pk_trabalho_fase_manut8 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM8_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM8_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_9
            'GRF_Trabalhos_Fases_Manut_9 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_10"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM9_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM9_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM9_fk_operador & vbCrLf
            sSQL = sSQL & ",@fk_avaliacao_tipo       = " & lFM9_fk_avaliacao_tipo & vbCrLf
            sSQL = sSQL & ",@dt_avaliacao            = " & IIf(dFM9_dt_Avaliacao = "NULL", dFM9_dt_Avaliacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM9_dt_Avaliacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_cliente        = " & IIf(dFM9_dt_envio_cliente = "NULL", dFM9_dt_envio_cliente, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM9_dt_envio_cliente), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_padroes_CQ     = " & IIf(dFM9_dt_envio_padroes_CQ = "NULL", dFM9_dt_envio_padroes_CQ, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM9_dt_envio_padroes_CQ), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM9_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_9
            'GRF_Trabalhos_Fases_Manut_9 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_10"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut9 = " & lFM9_pk_trabalho_fase_manut9 & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM9_fk_operador & vbCrLf
            sSQL = sSQL & ",@fk_avaliacao_tipo       = " & lFM9_fk_avaliacao_tipo & vbCrLf
            sSQL = sSQL & ",@dt_avaliacao            = " & IIf(dFM9_dt_Avaliacao = "NULL", dFM9_dt_Avaliacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM9_dt_Avaliacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_cliente        = " & IIf(dFM9_dt_envio_cliente = "NULL", dFM9_dt_envio_cliente, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM9_dt_envio_cliente), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_padroes_CQ     = " & IIf(dFM9_dt_envio_padroes_CQ = "NULL", dFM9_dt_envio_padroes_CQ, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM9_dt_envio_padroes_CQ), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM9_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_10
            'GRF_Trabalhos_Fases_Manut_10 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_11"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM10_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM10_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM10_fk_operador & vbCrLf
            sSQL = sSQL & ",@total_cores             = " & lFM10_total_cores & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM10_observacoes)), AspasSimples) & vbCrLf
         
        Case ALT_Manut_10
            'GRF_Trabalhos_Fases_Manut_10 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_11"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut10 = " & lFM10_pk_trabalho_fase_manut10 & vbCrLf
            sSQL = sSQL & ",@fk_operador              = " & lFM10_fk_operador & vbCrLf
            sSQL = sSQL & ",@total_cores              = " & lFM10_total_cores & vbCrLf
            sSQL = sSQL & ",@observacoes              = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM10_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_11
            'GRF_Trabalhos_Fases_Manut_11 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_12"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho             = " & lFM11_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                 = " & lFM11_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador             = " & lFM11_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes             = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM11_observacoes)), AspasSimples) & vbCrLf
 
        Case ALT_Manut_11
            'GRF_Trabalhos_Fases_Manut_11 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_12"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut11 = " & lFM11_pk_trabalho_fase_manut11 & vbCrLf
            sSQL = sSQL & ",@fk_operador              = " & lFM11_fk_operador & vbCrLf
            sSQL = sSQL & ",@observacoes              = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM11_observacoes)), AspasSimples) & vbCrLf
        
        Case INC_Manut_12
            'GRF_Trabalhos_Fases_Manut_12 (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_13"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho              = " & lFM12_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                  = " & lFM12_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador              = " & lFM12_fk_operador & vbCrLf
            sSQL = sSQL & ",@fk_avaliacao_tipo        = " & lFM12_fk_avaliacao_tipo & vbCrLf
            sSQL = sSQL & ",@dt_avaliacao             = " & IIf(dFM12_dt_Avaliacao = "NULL", dFM12_dt_Avaliacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_Avaliacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_receb_rolinho         = " & IIf(dFM12_dt_receb_rolinho = "NULL", dFM12_dt_receb_rolinho, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_receb_rolinho), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_laminacao       = " & IIf(dFM12_dt_envio_laminacao = "NULL", dFM12_dt_envio_laminacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_envio_laminacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_receb_laminacao       = " & IIf(dFM12_dt_receb_laminacao = "NULL", dFM12_dt_receb_laminacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_receb_laminacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_padroes_cliente = " & IIf(dFM12_dt_envio_padroes_cliente = "NULL", dFM12_dt_envio_padroes_cliente, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_envio_padroes_cliente), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_padroes_CQ      = " & IIf(dFM12_dt_envio_padroes_CQ = "NULL", dFM12_dt_envio_padroes_CQ, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_envio_padroes_CQ), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@observacoes              = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM12_observacoes)), AspasSimples) & vbCrLf
        
        Case ALT_Manut_12
            'GRF_Trabalhos_Fases_Manut_12 (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_13"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut12 = " & lFM12_pk_trabalho_fase_manut12 & vbCrLf
            sSQL = sSQL & ",@fk_operador              = " & lFM12_fk_operador & vbCrLf
            sSQL = sSQL & ",@fk_avaliacao_tipo        = " & lFM12_fk_avaliacao_tipo & vbCrLf
            sSQL = sSQL & ",@dt_avaliacao             = " & IIf(dFM12_dt_Avaliacao = "NULL", dFM12_dt_Avaliacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_Avaliacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_receb_rolinho         = " & IIf(dFM12_dt_receb_rolinho = "NULL", dFM12_dt_receb_rolinho, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_receb_rolinho), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_laminacao       = " & IIf(dFM12_dt_envio_laminacao = "NULL", dFM12_dt_envio_laminacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_envio_laminacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_receb_laminacao       = " & IIf(dFM12_dt_receb_laminacao = "NULL", dFM12_dt_receb_laminacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_receb_laminacao), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_padroes_cliente = " & IIf(dFM12_dt_envio_padroes_cliente = "NULL", dFM12_dt_envio_padroes_cliente, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_envio_padroes_cliente), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@dt_envio_padroes_CQ      = " & IIf(dFM12_dt_envio_padroes_CQ = "NULL", dFM12_dt_envio_padroes_CQ, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dFM12_dt_envio_padroes_CQ), AspasSimples)) & vbCrLf
            sSQL = sSQL & ",@observacoes              = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sFM12_observacoes)), AspasSimples) & vbCrLf

        Case INC_Gener_DtLibGrav
            'GRF_Trabalhos_Fases_Manut_Generico (Inclusão)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_98"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho              = " & lGEN_fk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                  = " & lGEN_fk_fase & vbCrLf
            sSQL = sSQL & ",@fk_operador              = " & lGEN_fk_operador & vbCrLf
            sSQL = sSQL & ",@dt_liberacao_gravacao    = " & IIf(dGEN_dt_liberacao_gravacao = "NULL", dGEN_dt_liberacao_gravacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dGEN_dt_liberacao_gravacao), AspasSimples)) & vbCrLf
        
        Case ALT_Gener_DtLibGrav
            'GRF_Trabalhos_Fases_Manut_Generico (Alteração)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_98"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho_fase_manut_gen = " & lGEN_pk_trabalho_fase_manut_gen & vbCrLf
            sSQL = sSQL & ",@fk_operador                = " & lGEN_fk_operador & vbCrLf
            sSQL = sSQL & ",@dt_liberacao_gravacao      = " & IIf(dGEN_dt_liberacao_gravacao = "NULL", dGEN_dt_liberacao_gravacao, oBasico.Geral.Aspas(oBasico.DataHora.FormataDataGravar(dGEN_dt_liberacao_gravacao), AspasSimples)) & vbCrLf

        Case ATU_Muda_Fase
            'PRC_GRF_Atualiza_Trabalhos_99 (Mudança de Fase na Manutenção do Trabalho)
            sProcedure = "PRC_GRF_Atualiza_Trabalhos_99"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @fk_trabalho                  = " & lTB_pk_trabalho & vbCrLf
            sSQL = sSQL & ",@fk_fase                      = " & lTFSO_pk_fase_operador & vbCrLf
            sSQL = sSQL & ",@fk_trabalho_fase_status_tipo = " & lTFSO_fk_trabalho_fase_status_tipo_operador & vbCrLf
            sSQL = sSQL & ",@fk_operador                  = " & lTFSO_fk_operador_status & vbCrLf
            sSQL = sSQL & ",@obs_status                   = " & oBasico.Geral.Aspas(SubstituiCaracteresEspeciaisGravacao(CStr(sTFSO_observacao_status)), AspasSimples) & vbCrLf
    End Select
    
    Set rsRegistros = Nothing
    Set rsRegistrosTemp = Nothing
    Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
    If rsRegistrosTemp.EOF Then
        sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA PROCEDURE dbo." & sProcedure & " !"
    Else
        sMensagem = rsRegistrosTemp!Mensagem
        If Val(rsRegistrosTemp!Status) = Sucesso Then Atualiza_SQL_Server = Sucesso
        Set rsRegistros = rsRegistrosTemp
    End If
    
    Exit Function
    
ErroRotina:
    If Err.Number = 424 Then
        sMensagem = "Não foi possível atualizar o(s) registro(s) da(s) tabela(s) através da procedure < " & sProcedure & " > ." & vbLf
        sMensagem = sMensagem & "Verificar violação de indices exclusivos, chaves estrangeiras inválidas ou qualquer parâmetro que possa ter invalidado a execução do procedimento!"
    Else
        sMensagem = "Ocorreu o erro: " & vbLf
        sMensagem = sMensagem & Err.Number & vbLf
        sMensagem = sMensagem & Err.Description & vbLf & vbLf
        sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sObjetoClasse & "." & "Atualiza_SQL_Server > "
    End If
End Function

Private Function Consulta_SQL_Server() As eStatus
    On Error GoTo ErroRotina
    
    Consulta_SQL_Server = Erro
    
    Select Case iAcaoFasesManut
        Case CON_Trab_Fases_Manut
            'Retorna Todos os Trabalhos Não Excluídos com suas Fases de Manutenção, Exceto os Históricos
            sConsulta = "VIW_GRF_Trabalhos_Fases_Manut"
        
            sSQL = Empty
            sSQL = sSQL & "dbo.PRC_SQL_PESQUISA_REGISTROS "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sConsulta, AspasSimples) & ", "
            sSQL = sSQL & "@Colunas=" & oBasico.Geral.Aspas(sColunasConsulta, AspasSimples) & ", "
            sSQL = sSQL & "@ColunasOrdem=" & oBasico.Geral.Aspas(sOrdemConsulta, AspasSimples) & ", "
            sSQL = sSQL & "@Condicoes=" & oBasico.Geral.Aspas(sCondicao, AspasSimples) & " "
        
        Case CON_Fase_Atual
            sConsulta = "GRF_Fases"
            
            'Retorna a Chave Primária do Status em que se Encontra Atualmente a Fase Atual do Trabalho
            sSQL = Empty
            sSQL = sSQL & "SELECT dbo.FNC_GRF_Trabalho_Fase_Atual(" & lTB_pk_trabalho & ")"
            
            Set rsRegistrosTemp = Nothing
            Set rsRegistrosTemp = cConexao.AbreRS(sSQL, True)
            If rsRegistrosTemp.EOF Then
                sMensagem = "OCORREU UM ERRO NÃO IDENTIFICADO NA EXECUÇÃO DA FUNÇÃO dbo.FNC_GRF_Trabalho_Fase_Atual !"
                Exit Function
            Else
                If Trim(UCase(rsRegistrosTemp.Fields(0).Name)) = "MENSAGEM" Then
                    sMensagem = rsRegistrosTemp!Mensagem
                    Exit Function
                Else
                    sCondicao = "pk_fase = " & rsRegistrosTemp.Fields(0).Value
                End If
            End If
    
            sSQL = Empty
            sSQL = sSQL & "dbo.PRC_SQL_PESQUISA_REGISTROS "
            sSQL = sSQL & "@Tabela=" & oBasico.Geral.Aspas(sConsulta, AspasSimples) & ", "
            sSQL = sSQL & "@Colunas=" & oBasico.Geral.Aspas(sColunasConsulta, AspasSimples) & ", "
            sSQL = sSQL & "@ColunasOrdem=" & oBasico.Geral.Aspas(sOrdemConsulta, AspasSimples) & ", "
            sSQL = sSQL & "@Condicoes=" & oBasico.Geral.Aspas(sCondicao, AspasSimples) & " "
            
        Case CON_GER_Tipo_1
            'Tipo de Consulta: Trabalho Fases Manutenção - Consulta 1 (TRAB.EM DETERMINADA FASE)
            sProcedure = "PRC_GRF_Consulta_Trabalhos_1"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_fase = " & lTFSA_pk_fase_atual & vbCrLf
        
        Case CON_GER_Tipo_2
            'Tipo de Consulta: Trabalho Fases Manutenção - Consulta 2 (TRAB.EM DETERMINADA FASE E STATUS)
            sProcedure = "PRC_GRF_Consulta_Trabalhos_2"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_fase                      = " & lTFSA_pk_fase_atual & ", " & vbCrLf
            sSQL = sSQL & " @fk_trabalho_fase_status_tipo = " & lTFSA_fk_trabalho_fase_status_tipo_atual & " " & vbCrLf

        Case CON_GER_Tipo_3
        Case CON_GER_Tipo_4
        Case CON_GER_Tipo_5
        Case CON_GER_Tipo_6
        Case CON_GER_Tipo_7
        Case CON_GER_Tipo_8
        Case CON_GER_Tipo_9
        Case CON_GER_Tipo_10
        Case CON_GER_Tipo_11
        Case CON_GER_Tipo_12
        
        Case CON_Resumo_Trab
            'Tipo de Consulta: Trabalho Fases Manutenção - Resumo de Lead Time
            sProcedure = "PRC_GRF_Consulta_Trabalhos_LeadTime_Resumo"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho = " & lTB_pk_trabalho & " " & vbCrLf
            
        Case CON_Status_Fase_Hist
            'Tipo de Consulta: Histórico de Status das Fases do Trabalho
            sProcedure = "PRC_GRF_Consulta_Status_Fases_Trabalhos_Hist"
            
            sSQL = Empty
            sSQL = sSQL & " dbo." & sProcedure & vbCrLf
            sSQL = sSQL & " @pk_trabalho = " & lTB_pk_trabalho & " " & vbCrLf
            sSQL = sSQL & ",@pk_fase     = " & lTFSO_pk_fase_operador & " " & vbCrLf
    End Select
    
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
