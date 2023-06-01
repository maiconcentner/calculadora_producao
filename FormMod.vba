Private Sub CommandButton2_Click()
Unload Me
formMenu.Show
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

Private Sub CommandButton8_Click()
'informativo calculo mao de obra direta

    MsgBox "Para cálculo da MOD (Mão de Obra Direta), é utilizado alguns parâmetros, como: tempo disponível de trabalho (518 minutos), índice de absenteísmo, fator de tolerância, demanda e tempo cronometrado para produzir cada unidade.", vbInformation, "Cálculo da MOD"


End Sub