Public Sub Planilha_OP_ReaisMil()

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1(6)
Dim colBalanco As Collection
Dim periodos As String
Dim count As Integer
Dim count2 As Integer
'ALTERAÇÃO FEITA 30/05: novas variáveis
Dim countLI11 As Integer
Dim countLI10 As Integer


Dim iCtr As Long

For iCtr = 0 To Front.ListBox1.ListCount - 1
   If Front.ListBox1.Selected(iCtr) = True Then
      arr1(count) = Front.ListBox1.list(iCtr)
      count = count + 1
   End If
Next iCtr

If count > 4 Then
MsgBox ("Limite de seleção de periodos ultrapassou")
Exit Sub
End If
       
 For iCtr = 0 To count - 1
    
   ax = Split(arr1(iCtr))
   cd_cli = ax(2)
   
   If iCtr = 0 Then
           
      periodos = "'" & ax(0) & "'"
           
    Else
       
      periodos = periodos & ", '" & ax(0) & "'"
           
    End If
       
  Next iCtr

  Set conn = getConnection()
  Set rs = New ADODB.Recordset
  
conn.Open

qry = "select * from LB_PLANI.FATO_balanco where dt_exerc in (" & periodos & ") and cd_cli = " & cd_cli

rs.Open qry, conn, adOpenStatic

Set colBalanco = New Collection

If Not IsNull(rs) Or rs.RecordCount > 0 Then
                       
   Do
          
     Dim blc As Balanco
     Set blc = New Balanco
     
    'ALTERAÇÃO FEITA 28/05: +NOVA PARAMETRIZAÇÃO

     With blc
        .DT_EXERC = rs![DT_EXERC]
      
      '------------------------ATIVO--------------------------
        
    
        .OP_DISPNVL = rs![OP_DISPNVL]
        .OP_CRED_A_CP = rs![OP_CRED_A_CP]
        .OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP = rs![OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP]
        .OP_ATIVO_INVESTIMENTOS = rs![OP_ATIVO_INVESTIMENTOS]
        .OP_ATIVO_CIRC_ESTOQ = rs![OP_ATIVO_CIRC_ESTOQ]
        .OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP = rs![OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP]
        .OP_CRED_A_LP = rs![OP_CRED_A_LP]
        .OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP = rs![OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP]
        .OP_INVESTIMENTOS = rs![OP_INVESTIMENTOS]
        .OP_ATV_RLZ_ESTOQ = rs![OP_ATV_RLZ_ESTOQ]
        .OP_ATV_RLZ_VPD_PAGAS_ANTECIP = rs![OP_ATV_RLZ_VPD_PAGAS_ANTECIP]
        .OP_INVESTIMENTOS = rs![OP_INVESTIMENTOS]
        .OP_IMOBILIZADO = rs![OP_IMOBILIZADO]
        .OP_INTANGIVEL = rs![OP_INTANGIVEL]

      '------------------------PASSIVO--------------------------
        
        .OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP = rs![OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP]
        .OP_EMPREST_FINAN_CP = rs![OP_EMPREST_FINAN_CP]
        .OP_FORN_CTAS_PG_CP = rs![OP_FORN_CTAS_PG_CP]
        .OP_OBRIG_FISCAIS_CP = rs![OP_OBRIG_FISCAIS_CP]
        .OP_OBRIG_REPART = rs![OP_OBRIG_REPART]
        .OP_PROV_CP = rs![OP_PROV_CP]
        .OP_DEMAIS_OBRIG_CP = rs![OP_DEMAIS_OBRIG_CP]
        '.OP_PASS_N_CIRC_OB_TRB_PREV_ASS_CP = rs![OP_PASS_N_CIRC_OB_TRB_PREV_ASS_CP]
        .OP_EMPREST_FINANC_LP = rs![OP_EMPREST_FINANC_LP]
        .OP_FORNECEDORES_LP = rs![OP_FORNECEDORES_LP]
        .OP_PREVISOES_LP = rs![OP_PREVISOES_LP]
        .OP_DEMAIS_OBRIG_LP = rs![OP_DEMAIS_OBRIG_LP]
        .OP_PATRIMONIO_LP = rs![OP_PATRIMONIO_LP]

        

      '--------------DEMONSTRATIVO DE RESULTADO----------------

        .OP_TRIBUTARIAS = rs![OP_TRIBUTARIAS]
        .OP_CONTRIBUICOES = rs![OP_CONTRIBUICOES]
        .OP_TRANSF_CORRENTES = rs![OP_TRANSF_CORRENTES]
        .OP_PATRIMONIAIS = rs![OP_PATRIMONIAIS]
        .OP_OUTRAS_RECEITAS_CORRENTES = rs![OP_OUTRAS_RECEITAS_CORRENTES]
        .OP_DEDUCOES = rs![OP_DEDUCOES]

        .OP_PESSOAL_ENCARGOS_SOCIAIS = rs![OP_PESSOAL_ENCARGOS_SOCIAIS]
        .OP_JUROS_ENCARGOS_DIVIDAS = rs![OP_JUROS_ENCARGOS_DIVIDAS]
        .OP_TRANSFERENCIAS_CORRENTES = rs![OP_TRANSFERENCIAS_CORRENTES]
        .OP_OUTRAS_DESPESAS_CORRENTES = rs![OP_OUTRAS_DESPESAS_CORRENTES]

        .OP_OPERACOES_CREDITO = rs![OP_OPERACOES_CREDITO]
        .OP_ALIENACAO_BENS = rs![OP_ALIENACAO_BENS]
        .OP_TRANSFERENCIA_CAPITAL = rs![OP_TRANSFERENCIA_CAPITAL]
        .OP_RECEITA_CAPITAL_OUTRAS = rs![OP_RECEITA_CAPITAL_OUTRAS]

        .OP_INVESTIMENTOS = rs![OP_INVESTIMENTOS]
        .OP_INVERSOES_FINANCEIRAS = rs![OP_INVERSOES_FINANCEIRAS]
        .OP_AMORTIZACAO_DIVIDA = rs![OP_AMORTIZACAO_DIVIDA]
        .OP_OUTRAS_DESPESAS_CAPITAL = rs![OP_OUTRAS_DESPESAS_CAPITAL]

        .OP_OUTRAS_RECEITAS_DESPESAS = rs![OP_OUTRAS_RECEITAS_DESPESAS]

        .OP_RESERVAS_CONTINGENCIAS = rs![OP_RESERVAS_CONTINGENCIAS]

      '--------------ORÇADO----------------

      .OP_ORCADO_TRIBUTARIAS = rs![OP_ORCADO_TRIBUTARIAS]
      .OP_ORCADO_CONTRIBUICOES = rs![OP_ORCADO_CONTRIBUICOES]
      .OP_ORCADO_TRANSF_CORRENTES = rs![OP_ORCADO_TRANSF_CORRENTES]
      .OP_ORCADO_PATRIMONIAIS = rs![OP_ORCADO_PATRIMONIAIS]
      .OP_ORCADO_OUTRAS_RECT_CORREN = rs![OP_ORCADO_OUTRAS_RECT_CORREN]
      .OP_ORCADO_DEDUCOES = rs![OP_ORCADO_DEDUCOES]

      .OP_ORCADO_PESS_ENCG_DIV = rs![OP_ORCADO_PESS_ENCG_DIV]
      .OP_ORCADO_JUROS_ENCARG_DIV = rs![OP_ORCADO_JUROS_ENCARG_DIV]
      .OP_ORCADO_TRANSF_CORR = rs![OP_ORCADO_TRANSF_CORR]
      .OP_ORCADO_OUTRAS_DESP_CORR = rs![OP_ORCADO_OUTRAS_DESP_CORR]

      .OP_ORCADO_OPER_CREDITO = rs![OP_ORCADO_OPER_CREDITO]
      .OP_ORCADO_ALIENACAO_BENS = rs![OP_ORCADO_ALIENACAO_BENS]
      .OP_ORCADO_TRANSF_CAPITAL = rs![OP_ORCADO_TRANSF_CAPITAL]
      .OP_ORCADO_RECT_CAPITAL_OUTRAS = rs![OP_ORCADO_RECT_CAPITAL_OUTRAS]

      .OP_ORCADO_INVESTIMENTOS = rs![OP_ORCADO_INVESTIMENTOS]
      .OP_ORCADO_INVERSOES_FIN = rs![OP_ORCADO_INVERSOES_FIN]
      .OP_ORCADO_AMORT_DIVIDA = rs![OP_ORCADO_AMORT_DIVIDA]
      .OP_ORCADO_OUTRAS_DESP_CAPITAL = rs![OP_ORCADO_OUTRAS_DESP_CAPITAL]

      .OP_ORCADO_OUTRAS_RECT_DESPESAS = rs![OP_ORCADO_OUTRAS_RECT_DESPESAS]

      .OP_ORCADO_RESERVAS_CONTI = rs![OP_ORCADO_RESERVAS_CONTI]

     End With
     
     colBalanco.Add Item:=blc
    
     rs.MoveNext
     Loop Until rs.EOF
     
End If
rs.Close
conn.Close

'ALTERAÇÃO FEITA 30/05: preenchendo células, +alteração no count, +alteração na Sheet

countLI1 = 2
countLI10 = 10
countLI11 = 11


For Each blc In colBalanco

    '------------------------ATIVO--------------------------

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(7, countLI1).Value = blc.OP_DISPNVL
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(8, countLI1).Value = blc.OP_CRED_A_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(9, countLI1).Value = blc.OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(10, countLI1).Value = blc.OP_ATIVO_INVESTIMENTOS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(11, countLI1).Value = blc.OP_ATIVO_CIRC_ESTOQ
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(12, countLI1).Value = blc.OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(16, countLI1).Value = blc.OP_CRED_A_LP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(17, countLI1).Value = blc.OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(18, countLI1).Value = blc.OP_INVESTIMENTOS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(19, countLI1).Value = blc.OP_ATV_RLZ_ESTOQ
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(20, countLI1).Value = blc.OP_ATV_RLZ_VPD_PAGAS_ANTECIP

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(21, countLI1).Value = blc.OP_INVESTIMENTOS

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(22, countLI1).Value = blc.OP_IMOBILIZADO

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(23, countLI1).Value = blc.OP_INTANGIVEL
         
    '------------------------PASSIVO------------------------

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(8, countLI11).Value = blc.OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(9, countLI11).Value = blc.OP_EMPREST_FINAN_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(10, countLI11).Value = blc.OP_FORN_CTAS_PG_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(11, countLI11).Value = blc.OP_OBRIG_FISCAIS_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(12, countLI11).Value = blc.OP_OBRIG_REPART
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(13, countLI11).Value = blc.OP_PROV_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(14, countLI11).Value = blc.OP_DEMAIS_OBRIG_CP

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(16, countLI11).Value = blc.OP_PASS_N_CIRC_OB_TRB_PREV_ASS_CP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(17, countLI11).Value = blc.OP_EMPREST_FINANC_LP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(18, countLI11).Value = blc.OP_FORNECEDORES_LP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(19, countLI11).Value = blc.OP_PREVISOES_LP
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(20, countLI11).Value = blc.OP_DEMAIS_OBRIG_LP

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(22, countLI11).Value = blc.OP_PATRIMONIO_LP

    '--------------DEMONSTRATIVO DE RESULTADO----------------

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(30, countLI1).Value = blc.OP_TRIBUTARIAS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(31, countLI1).Value = blc.OP_CONTRIBUICOES
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(32, countLI1).Value = blc.OP_TRANSF_CORRENTES
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(33, countLI1).Value = blc.OP_PATRIMONIAIS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(34, countLI1).Value = blc.OP_OUTRAS_RECEITAS_CORRENTES
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(35, countLI1).Value = blc.OP_DEDUCOES

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(37, countLI1).Value = blc.OP_PESSOAL_ENCARGOS_SOCIAIS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(38, countLI1).Value = blc.OP_JUROS_ENCARGOS_DIVIDAS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(39, countLI1).Value = blc.OP_TRANSFERENCIAS_CORRENTES
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(40, countLI1).Value = blc.OP_OUTRAS_DESPESAS_CORRENTES

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(43, countLI1).Value = blc.OP_OPERACOES_CREDITO
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(44, countLI1).Value = blc.OP_ALIENACAO_BENS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(45, countLI1).Value = blc.OP_TRANSFERENCIA_CAPITAL
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(46, countLI1).Value = blc.OP_RECEITA_CAPITAL_OUTRAS

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(48, countLI1).Value = blc.OP_INVESTIMENTOS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(49, countLI1).Value = blc.OP_INVERSOES_FINANCEIRAS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(50, countLI1).Value = blc.OP_AMORTIZACAO_DIVIDA
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(51, countLI1).Value = blc.OP_OUTRAS_DESPESAS_CAPITAL

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(52, countLI1).Value = blc.OP_OUTRAS_RECEITAS_DESPESAS

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(53, countLI1).Value = blc.OP_RESERVAS_CONTINGENCIAS

    '--------------ORÇADO----------------

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(30, countLI10).Value = blc.OP_ORCADO_TRIBUTARIAS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(31, countLI10).Value = blc.OP_ORCADO_CONTRIBUICOES
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(32, countLI10).Value = blc.OP_ORCADO_TRANSF_CORRENTES
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(33, countLI10).Value = blc.OP_ORCADO_PATRIMONIAIS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(34, countLI10).Value = blc.OP_ORCADO_OUTRAS_RECT_CORREN
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(35, countLI10).Value = blc.OP_ORCADO_DEDUCOES

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(37, countLI10).Value = blc.OP_ORCADO_PESS_ENCG_DIV
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(38, countLI10).Value = blc.OP_ORCADO_JUROS_ENCARG_DIV
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(39, countLI10).Value = blc.OP_ORCADO_TRANSF_CORR
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(40, countLI10).Value = blc.OP_ORCADO_OUTRAS_DESP_CORR

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(43, countLI10).Value = blc.OP_ORCADO_OPER_CREDITO
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(44, countLI10).Value = blc.OP_ORCADO_ALIENACAO_BENS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(45, countLI10).Value = blc.OP_ORCADO_TRANSF_CAPITAL
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(46, countLI10).Value = blc.OP_ORCADO_RECT_CAPITAL_OUTRAS

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(48, countLI10).Value = blc.OP_ORCADO_INVESTIMENTOS
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(49, countLI10).Value = blc.OP_ORCADO_INVERSOES_FIN
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(50, countLI10).Value = blc.OP_ORCADO_AMORT_DIVIDA
    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(51, countLI10).Value = blc.OP_ORCADO_OUTRAS_DESP_CAPITAL

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(52, countLI10).Value = blc.OP_ORCADO_OUTRAS_RECT_DESPESAS

    ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(53, countLI10).Value = blc.OP_ORCADO_RESERVAS_CONTI

    countLI1 = countLI1 + 1 'B
    countLI11 = countLI11 + 1 'K
    
    
    
     ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(6, countLI1).Value = blc.DT_EXERC
     ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(6, countLI11).Value = blc.DT_EXERC
    
     ActiveWorkbook.Worksheets("OP_ReaisMil").Cells(2, 17).Value = blc.CD_GRP
    
    ' Inserir os dados da planilha auxiliar

     'ActiveWorkbook.Worksheets("Aux").Cells(count5, 4).Value = blc.CD_GRP
     'ActiveWorkbook.Worksheets("Aux").Cells(count5, 5).Value = blc.cd_cli
     'ActiveWorkbook.Worksheets("Aux").Cells(count5, 7).Value = blc.FLG_GRP
     'ActiveWorkbook.Worksheets("Aux").Cells(count5, 13).Value = blc.CNPJ
     
     
'cd_grupo = blc.CD_GRP
'cd_cli = blc.cd_cli
'CNPJ = blc.CNPJ
'Layout = Front.ComboBox1.Text
     
     
          
    'count2 = count2 + 2 'C
    'count3 = count3 + 2 'S
    'count4 = count4 + 2 'AJ
    'count5 = count5 + 1
    
Next

colBalanco.count

'ALTERAÇÃO FEITA 28/05: +alteração na sheet, +alteração nas colunas.

  If colBalanco.count = 1 Then

    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("E:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("U:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("AL:AQ").EntireColumn.Hidden = True

   ElseIf colBalanco.count = 2 Then

    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("G:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("W:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("AN:AQ").EntireColumn.Hidden = True
    
   ElseIf colBalanco.count = 3 Then

    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("I:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("Y:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("OP_ReaisMil").Columns("AP:AQ").EntireColumn.Hidden = True
    
   End If
     

End Sub
