
Private Sub btCalcular_Click()
On Error Resume Next
'Linguagem dos cálculos

'Cálculo do Takt e tempo de ciclo

'Dimensionando variáveis

'Preparação
Dim TaktPrepA
Dim TaktPrepB
Dim TaktPrepC
Dim TaktPrepD
Dim TaktPrepE
'------------------------------
'Montagem
Dim TaktMontA
Dim TaktMontB
Dim TaktMontC
'------------------------------
'Funilaria
Dim TaktFunA
Dim TaktFunB
Dim TaktFunC
Dim TempoDisp
'------------------------------
'Takt geral montagem
Dim TaktGeralMontagem

'Dimensionado ciclos de cada posição

'Preparação
Dim CicloPrepA
Dim CicloPrepB
Dim CicloPrepC
Dim CicloPrepD
Dim CicloPrepE
Dim CicloPrepTotal
'-------------
'Montagem
Dim CicloMontA
Dim CicloMontB
Dim CicloMontC
Dim CicloMontTotal
Dim CicloMontTotal2
Dim CicloMontTotal10
Dim CicloMontTotal20
Dim CicloMontTotal30

'-------------
'Funilaria
Dim CicloFuniA
Dim CicloFuniB
Dim CicloFuniC
Dim CicloFuniTotal
'-------------

'Condições
TaktFunA = TaktMontA
TaktFunB = TaktMontB
TaktFunC = TaktMontC
TempoDisp = 518

'Definição dos ciclos

'Preparação
CicloPrepA = 2072
CicloPrepB = 2072
CicloPrepC = 2072
CicloPrepD = 518
CicloPrepE = 518
'----------------
'Montagem
CicloMontA = 518
CicloMontB = 518
CicloMontC = 518
'----------------
'Funilaria
CicloFuniA = 518
CicloFuniB = 518
CicloFuniC = 518
'----------------


'Primeiro Cálculo (Takt time)

'Preparação de Chassis
TaktPrepA = TempoDisp / txtPrepA.Value
TaktPrepB = TempoDisp / txtPrepB.Value
TaktPrepC = TempoDisp / txtPrepC.Value
TaktPrepD = TempoDisp / txtPrepD.Value
TaktPrepE = TempoDisp / txtPrepE.Value
'-------------------------------------
'Montagem
TaktMontA = TempoDisp / txtMontA.Value
TaktMontB = TempoDisp / txtMontB.Value
TaktMontC = TempoDisp / txtMontC.Value
'-------------------------------------
'Funilaria
TaktFunA = TempoDisp / txtFuniA.Value
TaktFunB = TempoDisp / txtFuniB.Value
TaktFunC = TempoDisp / txtFuniC.Value
'-------------------------------------
'Takt geral montagem
TaktGeralMontagem = 518 / txtMontTotal.Value


'Cálculo dos ciclos
'Ciclo da Preparação
If cbPrepLinha.Value = "A" Then

        CicloPrepTotal = (((CicloPrepA / 8) * (8 - cbPrepPos.Value) / 60) / 8.8)
    
        ElseIf cbPrepLinha.Value = "B" Then
        CicloPrepTotal = (((CicloPrepB / 8) * (8 - cbPrepPos.Value) / 60) / 8.8)
    
        ElseIf cbPrepLinha.Value = "C" Then
        CicloPrepTotal = (((CicloPrepC / 8) * (8 - cbPrepPos.Value) / 60) / 8.8)
             
        ElseIf cbPrepLinha.Value = "D" Then
        CicloPrepTotal = (((TaktPrepD) * (8 - cbPrepPos.Value) / 60) / 8.8)
    
        ElseIf cbPrepLinha.Value = "E" Then
        CicloPrepTotal = (((TaktPrepE) * (8 - cbPrepPos.Value) / 60) / 8.8)
    
    
        End If
    
        txtPrepCiclo.Value = FormatNumber(CicloPrepTotal, 2)
'----------------------------------------------------------------
'----------------------------------------------------------------
'Ciclo da Entrada da Montagem

If cbMontLinha.Value = "A" Then

        CicloMontTotal = (((TaktGeralMontagem * (4 - cbMontPosi.Value)) + TaktGeralMontagem) / 60) / 8.8
        
        ElseIf cbMontLinha.Value = "B" Then
        
        CicloMontTotal10 = TaktGeralMontagem * (4 - cbMontPosi.Value)
        CicloMontTotal20 = TaktGeralMontagem
        CicloMontTotal30 = 518 / (CDbl(txtMontC.Value) + CDbl(txtMontB.Value))
        
        CicloMontTotal = ((CicloMontTotal10 + CicloMontTotal20 + CicloMontTotal30) / 60) / 8.8
        
    
        
        ElseIf cbMontLinha.Value = "C" Then
        
        CicloMontTotal10 = TaktGeralMontagem * (4 - cbMontPosi.Value)
        CicloMontTotal20 = TaktGeralMontagem
        CicloMontTotal30 = 518 / (CDbl(txtMontC.Value) + CDbl(txtMontB.Value)) + (518 / txtMontC.Value)
        
        
        CicloMontTotal = ((CicloMontTotal10 + CicloMontTotal20 + CicloMontTotal30) / 60) / 8.8
   
        End If
    
         txtMontCicloEntrada.Value = FormatNumber(CicloMontTotal, 2)
         
         
'------------------------------------------------------------------
'------------------------------------------------------------------

'Ciclo da Linha de Montagem
If txtMontLinha.Value = "A" Then

        CicloMontTotal2 = ((TaktMontA * (3 - cbMontaPosi)) / 60) / 8.8
    
    
        ElseIf txtMontLinha.Value = "B" Then
        
        CicloMontTotal2 = ((TaktMontB * (3 - cbMontaPosi)) / 60) / 8.8
    
     
        ElseIf txtMontLinha.Value = "C" Then
             
        CicloMontTotal2 = ((TaktMontB * (3 - cbMontaPosi)) / 60) / 8.8
   
     
        End If
    
        txtMontCicloLinha.Value = FormatNumber(CicloMontTotal2, 2)
'---------------------------------------------------------------------

'Ciclo da Funilaria
If txtFuniLinha.Value = "A" Then

    CicloFuniTotal = (((TaktFunA) * (5 - cbFuniPosi.Value) / 60) / 8.8)
    
        ElseIf txtFuniLinha.Value = "B" Then
        CicloFuniTotal = (((TaktFunB) * (5 - cbFuniPosi.Value) / 60) / 8.8)
    
        ElseIf txtFuniLinha.Value = "C" Then
        CicloFuniTotal = (((TaktFunC) * (5 - cbFuniPosi.Value) / 60) / 8.8)
             
    
        End If
    
        txtFUniCiclo.Value = FormatNumber(CicloFuniTotal, 2)
        
        'soma dos ciclos
        txtCicloTotal.Value = FormatNumber(CicloMontTotal + CicloMontTotal2 + CicloPrepTotal + CicloFuniTotal, 2) & " DIA(S)"
        
        'MsgBox "O ciclo para cada área é de: " & vbCrLf & vbCrLf & "- PREPARAÇÃO DE CHASSIS: " & CicloPrepTotal & " dia(s)" & vbCrLf & "Entrada Montagem: " & CicloMontTotal & " dias(s)"
        
        MsgBox "Cálculo efetuado com sucesso!", vbOKOnly, "CÁLCULO DO CICLO"
        
        
End Sub


Private Sub cbMontLinha_Change()
txtMontLinha.Value = cbMontLinha.Value
txtFuniLinha.Value = cbMontLinha.Value
End Sub

Private Sub CommandButton1_Click()

End Sub



Private Sub txtFuniA_Change()
Dim soma_Carros As Integer
Dim FuniA As Integer
Dim FuniB As Integer
Dim FuniC As Integer
Dim FuniD As Integer
Dim FuniE As Integer


On Error Resume Next

FuniA = txtFuniA.Value
FuniB = txtFuniB.Value
FuniC = txtFuniC.Value
FuniD = txtFuniD.Value
FuniE = txtFuniE.Value

soma_Carros = FuniA + FuniB + FuniC + FuniD + FuniE

txtFuniTotal.Value = soma_Carros
End Sub

Private Sub txtFuniB_Change()
Dim soma_Carros As Integer
Dim FuniA As Integer
Dim FuniB As Integer
Dim FuniC As Integer
Dim FuniD As Integer
Dim FuniE As Integer


On Error Resume Next

FuniA = txtFuniA.Value
FuniB = txtFuniB.Value
FuniC = txtFuniC.Value
FuniD = txtFuniD.Value
FuniE = txtFuniE.Value

soma_Carros = FuniA + FuniB + FuniC + FuniD + FuniE

txtFuniTotal.Value = soma_Carros
End Sub

Private Sub txtFuniC_Change()
Dim soma_Carros As Integer
Dim FuniA As Integer
Dim FuniB As Integer
Dim FuniC As Integer



On Error Resume Next

FuniA = txtFuniA.Value
FuniB = txtFuniB.Value
FuniC = txtFuniC.Value


soma_Carros = FuniA + FuniB + FuniC + FuniD + FuniE

txtFuniTotal.Value = soma_Carros
End Sub


Private Sub txtMontA_Change()
Dim soma_Carros As Integer
Dim MontA As Integer
Dim MontB As Integer
Dim MontC As Integer



On Error Resume Next

MontA = txtMontA.Value
MontB = txtMontB.Value
MontC = txtMontC.Value


txtFuniA = txtMontA.Value



soma_Carros = MontA + MontB + MontC + MontD + MontE

txtMontTotal.Value = soma_Carros
End Sub

Private Sub txtMontB_Change()
Dim soma_Carros As Integer
Dim MontA As Integer
Dim MontB As Integer
Dim MontC As Integer



On Error Resume Next

MontA = txtMontA.Value
MontB = txtMontB.Value
MontC = txtMontC.Value

txtFuniB = txtMontB.Value


soma_Carros = MontA + MontB + MontC + MontD + MontE

txtMontTotal.Value = soma_Carros
End Sub

Private Sub txtMontC_Change()
Dim soma_Carros As Integer
Dim MontA As Integer
Dim MontB As Integer
Dim MontC As Integer



On Error Resume Next

MontA = txtMontA.Value
MontB = txtMontB.Value
MontC = txtMontC.Value

txtFuniC = txtMontC.Value

soma_Carros = MontA + MontB + MontC + MontD + MontE

txtMontTotal.Value = soma_Carros
End Sub

Private Sub txtPrepA_Change()

Dim soma_Carros As Integer
Dim PrepA As Integer
Dim PrepB As Integer
Dim PrepC As Integer
Dim PrepD As Integer
Dim PrepE As Integer


On Error Resume Next

PrepA = txtPrepA.Value
PrepB = txtPrepB.Value
PrepC = txtPrepC.Value
PrepD = txtPrepD.Value
PrepE = txtPrepE.Value

soma_Carros = PrepA + PrepB + PrepC + PrepD + PrepE

txtPrepTotal.Value = soma_Carros

End Sub



Private Sub txtPrepB_Change()
Dim soma_Carros As Integer
Dim PrepA As Integer
Dim PrepB As Integer
Dim PrepC As Integer
Dim PrepD As Integer
Dim PrepE As Integer


On Error Resume Next

PrepA = txtPrepA.Value
PrepB = txtPrepB.Value
PrepC = txtPrepC.Value
PrepD = txtPrepD.Value
PrepE = txtPrepE.Value

soma_Carros = PrepA + PrepB + PrepC + PrepD + PrepE

txtPrepTotal.Value = soma_Carros
End Sub

Private Sub txtPrepC_Change()
Dim soma_Carros As Integer
Dim PrepA As Integer
Dim PrepB As Integer
Dim PrepC As Integer
Dim PrepD As Integer
Dim PrepE As Integer


On Error Resume Next

PrepA = txtPrepA.Value
PrepB = txtPrepB.Value
PrepC = txtPrepC.Value
PrepD = txtPrepD.Value
PrepE = txtPrepE.Value

soma_Carros = PrepA + PrepB + PrepC + PrepD + PrepE

txtPrepTotal.Value = soma_Carros
End Sub

Private Sub txtPrepD_Change()
Dim soma_Carros As Integer
Dim PrepA As Integer
Dim PrepB As Integer
Dim PrepC As Integer
Dim PrepD As Integer
Dim PrepE As Integer


On Error Resume Next

PrepA = txtPrepA.Value
PrepB = txtPrepB.Value
PrepC = txtPrepC.Value
PrepD = txtPrepD.Value
PrepE = txtPrepE.Value

soma_Carros = PrepA + PrepB + PrepC + PrepD + PrepE

txtPrepTotal.Value = soma_Carros
End Sub

Private Sub txtPrepE_Change()
Dim soma_Carros As Integer
Dim PrepA As Integer
Dim PrepB As Integer
Dim PrepC As Integer
Dim PrepD As Integer
Dim PrepE As Integer


On Error Resume Next

PrepA = txtPrepA.Value
PrepB = txtPrepB.Value
PrepC = txtPrepC.Value
PrepD = txtPrepD.Value
PrepE = txtPrepE.Value

soma_Carros = PrepA + PrepB + PrepC + PrepD + PrepE

txtPrepTotal.Value = soma_Carros
End Sub


Private Sub UserForm_Initialize()

'configurando os combo box

cbPrepLinha.AddItem "A"
cbPrepLinha.AddItem "B"
cbPrepLinha.AddItem "C"
cbPrepLinha.AddItem "D"
cbPrepLinha.AddItem "E"

cbPrepPos.AddItem "0"
cbPrepPos.AddItem "1"
cbPrepPos.AddItem "2"
cbPrepPos.AddItem "3"
cbPrepPos.AddItem "4"
cbPrepPos.AddItem "5"
cbPrepPos.AddItem "6"
cbPrepPos.AddItem "7"
cbPrepPos.AddItem "8"

cbMontLinha.AddItem "A"
cbMontLinha.AddItem "B"
cbMontLinha.AddItem "C"

cbMontPosi.AddItem "0"
cbMontPosi.AddItem "1"
cbMontPosi.AddItem "2"
cbMontPosi.AddItem "3"
cbMontPosi.AddItem "4"
cbMontPosi.AddItem "5"
cbMontPosi.AddItem "6"

cbMontaPosi.AddItem "0"
cbMontaPosi.AddItem "1"
cbMontaPosi.AddItem "2"
cbMontaPosi.AddItem "3"
cbMontaPosi.AddItem "4"
cbMontaPosi.AddItem "5"
cbMontaPosi.AddItem "6"

cbFuniPosi.AddItem "0"
cbFuniPosi.AddItem "1"
cbFuniPosi.AddItem "2"
cbFuniPosi.AddItem "3"
cbFuniPosi.AddItem "4"
cbFuniPosi.AddItem "5"
cbFuniPosi.AddItem "6"

End Sub