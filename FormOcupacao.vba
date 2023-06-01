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

Private Sub CommandButton3_Click()

Unload Me
formMenu.Show
End Sub

Private Sub CommandButton7_Click()
'informativo ocupação/ociosidade

MsgBox "Se considerarmos que um dia de trabalho possuí 518 minutos e, que neste dia, um colaborador trabalhou por 418 minutos, obtremos sua ocupação realizando a divisão do tempo trabalhado pelo tempo disponível (418 / 518). A ociosidade é a diferença entre o tempo disponível e o tempo trabalhado divido pelo tempo disponível ([518-418] / 518]). Se a ociosidade for positiva, isso indica que o colaborador está com tempo de 'sobra', se for negativa o colaborador está com 'falta' de tempo.", vbInformation, "Cálculo da Ocupação e Ociosidade"


End Sub