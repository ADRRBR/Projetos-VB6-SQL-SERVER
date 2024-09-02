Attribute VB_Name = "mdlGRF_Trab_Fases"
Option Explicit

Public oBasico  As New ADRRBR_SIS_Basico.clsSIS_Basico
   
Public Function SubstituiCaracteresEspeciaisGravacao(pCampoTexto As String) As String
    SubstituiCaracteresEspeciaisGravacao = pCampoTexto
    SubstituiCaracteresEspeciaisGravacao = Replace(SubstituiCaracteresEspeciaisGravacao, "'", " ")
End Function
