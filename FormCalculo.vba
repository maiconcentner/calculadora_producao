Private Sub CommandButton1_Click()
On Error Resume Next
'calculo do takt time

Dim resultadotakttime

resultadotakttime = FormatNumber((tDisponivel.Value) / (dDiaria / nPosicoes), 2)

resultadoTakt.Value = resultadotakttime

MsgBox "O TAKT TIME PARA ESTE CASO É DE:  " & resultadotakttime & " MINUTOS/HORAS POR CARRO", vbOKOnly, "RESULTADO TAKT"

End Sub

Private Sub CommandButton10_Click()

Unload Me
FormFadiga.Show

End Sub

Private Sub CommandButton3_Click()
On Error Resume Next
'calculo da MOD

Dim calculomod
Dim calculomod2
Dim calculomod3

calculomod = (dMod * (tempoUnidadeMOD + ((toleranciaMOD.Value / 100) * tempoUnidadeMOD))) / (tDisponivelMOD)
calculomod2 = calculomod * (absenteismo / 100)
calculomod3 = calculomod + calculomod2


modNecessaria.Value = FormatNumber(calculomod3, 2)

capacidadeProdutiva.Value = FormatNumber((tDisponivelMOD / ((tempoUnidadeMOD + (toleranciaMOD.Value / 100) * tempoUnidadeMOD))), 2)


MsgBox "A MOD NECESSÁRIA É DE: " & Chr(13) & FormatNumber(calculomod3, 2) & " COLABORADORES", vbOKOnly, "RESULTADO MOD"


End Sub



Private Sub CommandButton2_Click()
On Error Resume Next
'calculo de ocupação

Dim resultadocapacidadeprodutiva
Dim ocioso
Dim toleranciapadrao
Dim minutohomem

minutohomem = tempoUnidade.Value * dCapacidade.Value

toleranciapadrao = tolerencia.Value / 100

'resultadocapacidadeprodutiva = FormatNumber(((dCapacidade * (tempoUnidade * (1 + toleranciapadrao))) / (tDisponivelCapacidade * qModCapacidade)) * 100, 2) & "%"
'ocioso = FormatNumber((100 - ((dCapacidade * (tempoUnidade * (1 + toleranciapadrao))) / (tDisponivelCapacidade * qModCapacidade)) * 100), 2)


resultadocapacidadeprodutiva = FormatNumber(((minutohomem * (1 + toleranciapadrao)) / (tDisponivelCapacidade * qModCapacidade)) * 100, 2) & "%"
ocioso = FormatNumber((100 - ((minutohomem * (1 + toleranciapadrao)) / (tDisponivelCapacidade * qModCapacidade)) * 100), 2)




ociosidade.Value = ocioso & "%"
resultadoCapacidade.Value = resultadocapacidadeprodutiva

MsgBox "A OCUPAÇÃO É DE:  " & resultadocapacidadeprodutiva & Chr(13) & Chr(13) & "A OCIOSIDADE É DE:  " & ocioso & "%", vbOKOnly, "RESULTADO OCUPAÇÃO"

End Sub




Private Sub CommandButton4_Click()


'calculo do tempo de ciclo


Dim tciclo

tciclo = FormatNumber(tProducao / qntProduzidaTC, 2)

'calculo mod

'modTC = FormatNumber(tDisponivelTC / tciclo, 2)

tempoCiclo.Value = tciclo
'qntMOD.Value = modTC

MsgBox "O TEMPO DE CICLO É DE:  " & tciclo & " MINUTOS POR CARRO", vbOKOnly, "RESULTADO TEMPO DE CICLO"


End Sub

Private Sub CommandButton6_Click()

'informativo takt time

MsgBox "Takt Time é o ritmo no qual você precisa completar um produto para suprir a demanda do consumidor. Por exemplo, uma fábrica trabalha 8 horas/dia (480 minutos) e a demanda do mercado é de 120 unidades/dia. Desta forma, o takt time é de 4 minutos (480/120).", vbInformation, "TAKT TIME"

End Sub

Private Sub CommandButton7_Click()
'informativo ocupação/ociosidade

MsgBox "Se considerarmos que um dia de trabalho possuí 518 minutos e, que neste dia, um colaborador trabalhou por 418 minutos, obtremos sua ocupação realizando a divisão do tempo trabalhado pelo tempo disponível (418 / 518). A ociosidade é a diferença entre o tempo disponível e o tempo trabalhado divido pelo tempo disponível ([518-418] / 518]). Se a ociosidade for positiva, isso indica que o colaborador está com tempo de 'sobra', se for negativa o colaborador está com 'falta' de tempo.", vbInformation, "Cálculo da Ocupação e Ociosidade"


End Sub


Private Sub CommandButton5_Click()
'limpar campos
On Error Resume Next
tDisponivel.Value = ""
nPosicoes.Value = ""
tDisponivelCapacidade.Value = ""
qModCapacidade.Value = ""
tolerencia.Value = ""
absenteismo.Value = ""
tDisponivelMOD = ""
dDiaria.Value = ""
resultadoTakt.Value = ""
dCapacidade.Value = ""
tempoUnidade.Value = ""
resultadoCapacidade.Value = ""
ociosidade.Value = ""
dMod.Value = ""
tempoUnidadeMOD.Value = ""
capacidadeProdutiva.Value = ""
modNecessaria.Value = ""
tProducao.Value = ""
qntProduzidaTC.Value = ""
tempoCiclo.Value = ""

Call UserForm_Initialize


End Sub


Private Sub CommandButton8_Click()
'informativo calculo mao de obra direta

    MsgBox "Para cálculo da MOD (Mão de Obra Direta), é utilizado alguns parâmetros, como: tempo disponível de trabalho (518 minutos), índice de absenteísmo, fator de tolerância, demanda e tempo cronometrado para produzir cada unidade.", vbInformation, "Cálculo da MOD"


End Sub

Private Sub CommandButton9_Click()
'informativo tempo de ciclo

    MsgBox "Por exemplo, uma linha produziu 800 peças em 6 horas de trabalho. Neste caso, o tempo de ciclo será de 6 horas x 60 minutos / 800 = 0,45 minuto/peça, ou seja, a cada 0,45 minuto a linha deverá produzir uma peça para que se alcance a produção de 800 peças no final do dia.", vbInformation, "Cálculo  do Tempo de Ciclo"

End Sub


Private Sub UserForm_Initialize()

    
    On Error Resume Next
    tDisponivel.Value = 518
    nPosicoes.Value = 1
    tDisponivelCapacidade.Value = 518
    qModCapacidade.Value = 1
    tolerencia.Value = 0
    absenteismo.Value = 0
    tDisponivelMOD = 518
    toleranciaMOD.Value = 0

End Sub