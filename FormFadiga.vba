Sub salvar_pdf()

'setando as variáveis
Dim Monotonia1
Dim Ambientais1
Dim Fisico1
Dim Mental1
Dim Recuperacao1
Dim Termicas1
Dim Atmosfericas1
Dim Ruido1
Dim Umidade1
Dim Vibracao1

'Esforço Físico

If OptionButton1.Value = True Then
    Fisico1 = "Muito Leve"
        
        ElseIf OptionButton2.Value = True Then
        Fisico1 = "Leve"
        
            ElseIf OptionButton3.Value = True Then
            Fisico1 = "Médio"
        
                ElseIf OptionButton4.Value = True Then
                Fisico1 = "Pesado"
                    Else
                    Fisico1 = "Muito pesado"
            End If


'Esforço Mental

 If OptionButton7.Value = True Then
    Mental1 = "Leve"
    
        ElseIf OptionButton8.Value = True Then
        Mental1 = "Médio"
        
            Else
            Mental1 = "Pesado"
                   
    End If
    
    
    'tempo de recuperação
    
    If cbRecuperacao.Value = "00 a 05" Then
        Recuperacao1 = "00 a 05"
         
            ElseIf cbRecuperacao.Value = "06 a 10" Then
            Recuperacao1 = "06 a 10"
            
                ElseIf cbRecuperacao.Value = "11 a 15" Then
                Recuperacao1 = "11 a 15"
                    
                    ElseIf cbRecuperacao.Value = "16 a 20" Then
                    Recuperacao1 = "16 a 20"
                        
                        ElseIf cbRecuperacao.Value = "21 a 25" Then
                        Recuperacao1 = "21 a 25"
                        
                            ElseIf cbRecuperacao.Value = "26 a 30" Then
                            Recuperacao1 = "26 a 30"
                            
                                ElseIf cbRecuperacao.Value = "31 a 35" Then
                                Recuperacao1 = "31 a 35"
                                
                                    ElseIf cbRecuperacao.Value = "36 a 40" Then
                                    Recuperacao1 = "36 a 40"
                                        
                                        ElseIf cbRecuperacao.Value = "41 a 45" Then
                                        Recuperacao1 = "41 a 45"
                                        
                                            ElseIf cbRecuperacao.Value = "46 a 50" Then
                                            Recuperacao1 = "46 a 50"
                                                
                                                ElseIf cbRecuperacao.Value = "51 a 55" Then
                                                Recuperacao1 = "51 a 55"
                                                    
                                                    ElseIf cbRecuperacao.Value = "56 a 60" Then
                                                    Recuperacao1 = "56 a 60"
                                        
                                            End If
            
    
'Monotonia

      If cbMonotonia.Value = "0,00 a 0,05" Then
        Monotonia1 = "0,00 a 0,05"
         
            ElseIf cbMonotonia.Value = "0,06 a 0,25" Then
            Monotonia1 = "0,06 a 0,25"
            
                ElseIf cbMonotonia.Value = "0,26 a 0,50" Then
                Monotonia1 = "0,26 a 0,50"
                    
                    ElseIf cbMonotonia.Value = "0,51 a 1,00" Then
                    Monotonia1 = "0,51 a 1,00"
                        
                        ElseIf cbMonotonia.Value = "1,01 a 2,00" Then
                        Monotonia1 = "1,01 a 2,00"
                        
                            ElseIf cbMonotonia.Value = "2,01 a 3,00" Then
                            Monotonia1 = "2,01 a 3,00"
                            
                                ElseIf cbMonotonia.Value = "3,01 a 4,00" Then
                                Monotonia1 = "3,01 a 4,00"
                                
                                    ElseIf cbMonotonia.Value = "Acima de 4,00" Then
                                    Monotonia1 = "Acima de 4,00"
                                    
                            End If

'Condições térmicas
    
    If OptionButton10.Value = True Then
        Termicas = 3.6 / 100
        Termicas1 = "Gelado"
    
          ElseIf OptionButton11.Value = True Then
            Termicas = 1.8 / 100
            Termicas1 = "Baixa"
        
              ElseIf OptionButton12.Value = True Then
                Termicas = 0 / 100
                Termicas1 = "Normal"
        
                 ElseIf OptionButton13.Value = True Then
                 Termicas = 1.8 / 100
                 Termicas1 = "Alta"
        
                    Else
                      Termicas = 3.6 / 100
                      Termicas1 = "Excessiva"
                    
            End If


'Condições atmosféricas

    If OptionButton15.Value = True Then
        Atmosfericas = 0 / 100
        Atmosfericas1 = "Boas"
    
        ElseIf OptionButton16.Value = True Then
          Atmosfericas = 2.4 / 100
          Atmosfericas1 = "Razoáveis"
            Else
              Atmosfericas = 5.6 / 100
              Atmosfericas1 = "Más"
                   
    End If
    
'Ruido

    If OptionButton18.Value = True Then
        Ruido = 0 / 100
        Ruido1 = "Baixo nível"
    
        Else
          Ruido = 1.8 / 100
          Ruido1 = "Excessivo"
                                  
    End If

'Umidade
 
    If OptionButton20.Value = True Then
        Umidade = 0 / 100
        Umidade1 = "Ambiente agradável e seco"
    
        ElseIf OptionButton21.Value = True Then
          Umidade = 1.8 / 100
          Umidade1 = "Umidade excessiva até 26°C"
        
            Else
              Umidade = 3.6 / 100
              Umidade1 = "Umidade excessiva até 40°C"
                   
    End If
       
   
'Vibração
   
    If OptionButton23.Value = True Then
        Vibracao = 0 / 100
        Vibracao1 = "Não há vibração"
    
        Else
          Vibracao = 1.8 / 100
          Vibracao1 = "Vibração do solo ou máquina"
                                  
    End If



'Cadastro de dados na aba PDF

'Informações gerais
Sheets("PDF").Range("G3").Value = txtSetor.Value
Sheets("PDF").Range("G5").Value = txtAtividade.Value
Sheets("PDF").Range("G7").Value = txtNome.Value
Sheets("PDF").Range("L5").Value = txtPosto.Value
Sheets("PDF").Range("P3").Value = txtFuncao.Value
Sheets("PDF").Range("P5").Value = txtCC.Value
Sheets("PDF").Range("P7").Value = Date
Sheets("PDF").Range("L27").Value = txtFatorA.Value

'Cálculos
Sheets("PDF").Range("F16").Value = Fisico1
Sheets("PDF").Range("F19").Value = Mental1
Sheets("PDF").Range("F22").Value = Recuperacao1
Sheets("PDF").Range("J16").Value = Monotonia1
Sheets("PDF").Range("J19").Value = Termicas1
Sheets("PDF").Range("J22").Value = Atmosfericas1
Sheets("PDF").Range("O16").Value = Ruido1
Sheets("PDF").Range("O19").Value = Umidade1
Sheets("PDF").Range("O22").Value = Vibracao1
Sheets("PDF").Range("Q27").Value = txtFadiga.Value

Application.ScreenUpdating = False

    Sheets("PDF").Select
        ChDir "G:\15- Equipe - C&C e Balanceamento\18 - Calculadora de Produção\Calculos_tolerancia"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "G:\15- Equipe - C&C e Balanceamento\18 - Calculadora de Produção\Calculos_tolerancia\F_" & txtPosto.Value & "_" & txtSetor & "_" & txtFuncao & "_" & txtAtividade & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
        Sheets("Painel").Select
        
        
          
Application.ScreenUpdating = True

End Sub

Sub CadastrarFator()

Dim GM As String

'CÓDIGO DE CADASTRO NA BASE DE DADOS

'PROCV NO GM

'On Error Resume Next
'GM = Sheets("BaseFator").Range("A:A").Find(txtPosto.Value).Row

'If GM <> "" Then

    'Linha = Sheets("BaseFator").Range("A:A").Find(txtPosto.Value).Row
 
    'Sheets("BaseFator").Cells(Linha, 1).Value = txtPosto.Value
    'Sheets("BaseFator").Cells(Linha, 2).Value = txtCC.Value
    'Sheets("BaseFator").Cells(Linha, 3).Value = txtSetor.Value
    'Sheets("BaseFator").Cells(Linha, 4).Value = txtAtividade.Value
    'Sheets("BaseFator").Cells(Linha, 5).Value = txtFuncao.Value
    'Sheets("BaseFator").Cells(Linha, 6).Value = Date
    'Sheets("BaseFator").Cells(Linha, 7).Value = txtNome.Value
    'Sheets("BaseFator").Cells(Linha, 8).Value = txtFadiga.Value

    'MsgBox "Informações cadastradas com sucesso!", vbOKOnly, "Cadastro de Tolerância"
    
    'Else
    
    Linha = 1 + Sheets("BaseFator").Range("A1000000").End(xlUp).Row
    Sheets("BaseFator").Cells(Linha, 1).Value = txtPosto.Value
    Sheets("BaseFator").Cells(Linha, 2).Value = txtCC.Value
    Sheets("BaseFator").Cells(Linha, 3).Value = txtSetor.Value
    Sheets("BaseFator").Cells(Linha, 4).Value = txtAtividade.Value
    Sheets("BaseFator").Cells(Linha, 5).Value = txtFuncao.Value
    Sheets("BaseFator").Cells(Linha, 6).Value = Date
    Sheets("BaseFator").Cells(Linha, 7).Value = txtNome.Value
    Sheets("BaseFator").Cells(Linha, 8).Value = txtFadiga.Value

    MsgBox "Informações cadastradas com sucesso!", vbOKOnly, "Cadastro de Tolerância"


'End If

End Sub



Private Sub CommandButton1_Click()
'calculo da fadiga


'declaração de variáveis
Dim FatorA
Dim FatorB
Dim Monotonia
Dim Ambientais
Dim Pessoais
Dim Fadiga
Dim Fisico
Dim Mental
Dim Recuperacao
Dim Termicas
Dim Atmosfericas
Dim Ruido
Dim Umidade
Dim Vibracao
Dim TotalAmbiental
Dim Resultado As VbMsgBoxResult

'definindo intervalos

'Esforço Físico

If OptionButton1.Value = True Then
    Fisico = 1.8 / 100
    
        ElseIf OptionButton2.Value = True Then
        Fisico = 3.6 / 100
        
            ElseIf OptionButton3.Value = True Then
            Fisico = 5.4 / 100
        
                ElseIf OptionButton4.Value = True Then
                Fisico = 7.2 / 100
        
                    Else
                    Fisico = 9 / 100
            End If


'Esforço Mental

 If OptionButton7.Value = True Then
    Mental = 0.6 / 100
    
        ElseIf OptionButton8.Value = True Then
        Mental = 1.8 / 100
        
            Else
            Mental = 3 / 100
                   
    End If
    
    
    'tempo de recuperação
    
    If cbRecuperacao.Value = "00 a 05" Then
        Recuperacao = 1
         
            ElseIf cbRecuperacao.Value = "06 a 10" Then
            Recuperacao = 0.9
            
                ElseIf cbRecuperacao.Value = "11 a 15" Then
                Recuperacao = 0.8
                    
                    ElseIf cbRecuperacao.Value = "16 a 20" Then
                    Recuperacao = 0.71
                        
                        ElseIf cbRecuperacao.Value = "21 a 25" Then
                        Recuperacao = 0.62
                        
                            ElseIf cbRecuperacao.Value = "26 a 30" Then
                            Recuperacao = 0.54
                            
                                ElseIf cbRecuperacao.Value = "31 a 35" Then
                                Recuperacao = 0.46
                                
                                    ElseIf cbRecuperacao.Value = "36 a 40" Then
                                    Recuperacao = 0.39
                                        
                                        ElseIf cbRecuperacao.Value = "41 a 45" Then
                                        Recuperacao = 0.32
                                        
                                            ElseIf cbRecuperacao.Value = "46 a 50" Then
                                            Recuperacao = 0.26
                                                
                                                ElseIf cbRecuperacao.Value = "51 a 55" Then
                                                Recuperacao = 0.2
                                                    
                                                    ElseIf cbRecuperacao.Value = "56 a 60" Then
                                                    Recuperacao = 0.15
                                        
                                            End If
            
    
'Monotonia

      If cbMonotonia.Value = "0,00 a 0,05" Then
        Monotonia = 7.8 / 100
         
            ElseIf cbMonotonia.Value = "0,06 a 0,25" Then
            Monotonia = 5.4 / 100
            
                ElseIf cbMonotonia.Value = "0,26 a 0,50" Then
                Monotonia = 3.6 / 100
                    
                    ElseIf cbMonotonia.Value = "0,51 a 1,00" Then
                    Monotonia = 2.1 / 100
                        
                        ElseIf cbMonotonia.Value = "1,01 a 2,00" Then
                        Monotonia = 1 / 100
                        
                            ElseIf cbMonotonia.Value = "2,01 a 3,00" Then
                            Monotonia = 0.5 / 100
                            
                                ElseIf cbMonotonia.Value = "3,01 a 4,00" Then
                                Monotonia = 0.2 / 100
                                
                                    ElseIf cbMonotonia.Value = "Acima de 4,00" Then
                                    Monotonia = 0 / 100
                                    
                            End If

'Condições térmicas
    
    If OptionButton10.Value = True Then
        Termicas = 3.6 / 100
    
          ElseIf OptionButton11.Value = True Then
            Termicas = 1.8 / 100
        
              ElseIf OptionButton12.Value = True Then
                Termicas = 0 / 100
        
                 ElseIf OptionButton13.Value = True Then
                 Termicas = 1.8 / 100
        
                    Else
                      Termicas = 3.6 / 100
                    
            End If


'Condições atmosféricas

    If OptionButton15.Value = True Then
        Atmosfericas = 0 / 100
    
        ElseIf OptionButton16.Value = True Then
          Atmosfericas = 2.4 / 100
        
            Else
              Atmosfericas = 5.6 / 100
                   
    End If
    
'Ruido

    If OptionButton18.Value = True Then
        Ruido = 0 / 100
    
        Else
          Ruido = 1.8 / 100
                                  
    End If

'Umidade
 
    If OptionButton20.Value = True Then
        Umidade = 0 / 100
    
        ElseIf OptionButton21.Value = True Then
          Umidade = 1.8 / 100
        
            Else
              Umidade = 3.6 / 100
                   
    End If
       
   
'Vibração
   
    If OptionButton23.Value = True Then
        Vibracao = 0 / 100
    
        Else
          Vibracao = 1.8 / 100
                                  
    End If


'Calculo do Fator A

FatorA = Fisico + Mental
FatorB = Recuperacao
TotalAmbiental = Termicas + Atmosfericas + Ruido + Umidade + Vibracao

'calculando  a fadiga

Fadiga = ((FatorA * FatorB) + Monotonia + TotalAmbiental + (5 / 100)) * 100
txtFatorA.Value = FormatNumber(((FatorA * FatorB) + Monotonia + TotalAmbiental) * 100, 2) & "%"
txtFadiga.Value = FormatNumber(Fadiga, 2) & "%"

MsgBox "NECESSIDADES PESSOAIS:  5,00%" & Chr(13) & "FATOR A:  " & FormatNumber(Fadiga - 5, 2) & "%" & Chr(13) & Chr(13) & "FADIGA TOTAL:  " & FormatNumber(Fadiga, 2) & "%", vbOKOnly, "FATOR DE FADIGA"
Resultado = MsgBox("Deseja salvar informações na Base de Dados?", vbYesNo, "Cadastro de Informações!")

If Resultado = vbYes Then
Call CadastrarFator

End If


End Sub

Private Sub CommandButton2_Click()

Unload Me
formMenu.Show

End Sub


Private Sub CommandButton3_Click()

'limoar formulario
txtPosto.Value = ""
txtSetor.Value = ""
txtCC.Value = ""
txtFuncao.Value = ""
txtAtividade.Value = ""
txtFatorA.Value = ""
txtFadiga.Value = ""

End Sub

Private Sub CommandButton4_Click()

Call salvar_pdf

End Sub


Private Sub txtAtividade_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCC_Change()

End Sub

Private Sub txtFuncao_Change()

End Sub

Private Sub txtFuncao_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtPosto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSetor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub UserForm_Initialize()

'Tempo de recuperação
cbRecuperacao.AddItem "00 a 05"
cbRecuperacao.AddItem "06 a 10"
cbRecuperacao.AddItem "11 a 15"
cbRecuperacao.AddItem "16 a 20"
cbRecuperacao.AddItem "21 a 25"
cbRecuperacao.AddItem "26 a 30"
cbRecuperacao.AddItem "31 a 35"
cbRecuperacao.AddItem "36 a 40"
cbRecuperacao.AddItem "41 a 45"
cbRecuperacao.AddItem "46 a 50"
cbRecuperacao.AddItem "51 a 55"
cbRecuperacao.AddItem "56 a 60"

'Monotonia

cbMonotonia.AddItem "0,00 a 0,05"
cbMonotonia.AddItem "0,06 a 0,25"
cbMonotonia.AddItem "0,26 a 0,50"
cbMonotonia.AddItem "0,51 a 1,00"
cbMonotonia.AddItem "1,01 a 2,00"
cbMonotonia.AddItem "2,01 a 3,00"
cbMonotonia.AddItem "3,01 a 4,00"
cbMonotonia.AddItem "Acima de 4,00"

txtNec.Value = "5%"
txtData.Value = Date
txtNome.Value = Application.UserName


End Sub