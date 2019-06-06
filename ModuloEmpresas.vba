Public Sub Planilha_PJ_ReaisMil()


Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1(6)
Dim colBalanco As Collection
Dim periodos As String
Dim count As Integer
Dim count2 As Integer
Dim count13 As Integer
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

        .EMPRS_DISPS = rs![EMPRS_DISPS]
        .EMPRS_CLI = rs![EMPRS_CLI]
        .EMPRS_PROV_DEV_DUVIDS = rs![EMPRS_PROV_DEV_DUVIDS]
        .EMPRS_ESTOQS = rs![EMPRS_ESTOQS]
        .EMPRS_ADTO_FORN = rs![EMPRS_ADTO_FORN]
        .EMPRS_TIT_VAL_MOBIL = rs![EMPRS_TIT_VAL_MOBIL]
        .EMPRS_DESP_PG_ANTEC = rs![EMPRS_DESP_PG_ANTEC]
        .EMPRS_ATIVO_OUTRAS_CTAS_OPERAC = rs![EMPRS_ATIVO_OUTRAS_CTAS_OPERAC]
        .EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER = rs![EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER]
        .EMPRS_ATIVO_CC_CONTROL_COLIG = rs![EMPRS_ATIVO_CC_CONTROL_COLIG]
        .EMPRS_ATIVO_OUTRAS_CONTAS = rs![EMPRS_ATIVO_OUTRAS_CONTAS]
        .EMPRS_PART_CONTROL_COLIGS = rs![EMPRS_PART_CONTROL_COLIGS]
        .EMPRS_OUTROS_INVEST = rs![EMPRS_OUTROS_INVEST]
        .EMPRS_IMOB_TECN_LIQ = rs![EMPRS_IMOB_TECN_LIQ]
        'O NOME DO CAMPO ABAIXO NÃO BATE COM O NOME DA CÉLULA, MAS FOI CONFERIDO NO DE-PARA
        .EMPRS_ATIVO_INTANG = rs![EMPRS_ATIVO_INTANG]
        .DT_EXERC = rs![DT_EXERC]
        .EMPRS_RECT_BRUTA = rs![EMPRS_RECT_BRUTA]
        .EMPRS_DEVOL_ABATIM = rs![EMPRS_DEVOL_ABATIM]
        .EMPRS_IMPOS_FATRDS = rs![EMPRS_IMPOS_FATRDS]
        .EMPRS_CUSTO_PROD_VENDS = rs![EMPRS_CUSTO_PROD_VENDS]
        .EMPRS_DEPREC = rs![EMPRS_DEPREC]
        .EMPRS_DESP_ADMINS = rs![EMPRS_DESP_ADMINS]
        .EMPRS_DESP_VNDAS = rs![EMPRS_DESP_VNDAS]
        .EMPRS_OUTRAS_DESP_REC_OPERAC = rs![EMPRS_OUTRAS_DESP_REC_OPERAC]
        .EMPRS_SALDO_COR_MONET = rs![EMPRS_SALDO_COR_MONET]
        .EMPRS_RECT_FINANC = rs![EMPRS_RECT_FINANC]
        .EMPRS_DESP_FINANC = rs![EMPRS_DESP_FINANC]
        .EMPRS_VAR_CAMBL_LIQ = rs![EMPRS_VAR_CAMBL_LIQ]
        .EMPRS_RECT_DESP_NAO_OPERAC = rs![EMPRS_RECT_DESP_NAO_OPERAC]
        .EMPRS_EQUIV_PATRIOM = rs![EMPRS_EQUIV_PATRIOM]
        .EMPRS_IMP_RNDA_CONTRIB_SOC = rs![EMPRS_IMP_RNDA_CONTRIB_SOC]
        .EMPRS_PARTIC = rs![EMPRS_PARTIC]

        '-----------------------PASSIVO-------------------------

        .EMPRS_FORNS = rs![EMPRS_FORNS]
        .EMPRS_OBRIG_SOC_TRIBUT = rs![EMPRS_OBRIG_SOC_TRIBUT]
        .EMPRS_ADTO_CLI = rs![EMPRS_ADTO_CLI]
        .EMPRS_EMPREST_FINANCS = rs![EMPRS_EMPREST_FINANCS]
        .EMPRS_DUPLIC_DESCTS = rs![EMPRS_DUPLIC_DESCTS]
        .EMPRS_CAMB = rs![EMPRS_CAMB]
        .EMPRS_PASS_CC_CONTROL_COLIG = rs![EMPRS_PASS_CC_CONTROL_COLIG]
        .EMPRS_PASS_OUTRAS_CTAS_OPERAC = rs![EMPRS_PASS_OUTRAS_CTAS_OPERAC]
        .EMPRS_PASS_OUTRAS_CTAS_NAO_OPER = rs![EMPRS_PASS_OUTRAS_CTAS_NAO_OPER]
        .EMPRS_EMPREST_FINANCS = rs![EMPRS_EMPREST_FINANCS]
        .EMPRS_PASS_OUTRAS_CONTAS = rs![EMPRS_PASS_OUTRAS_CONTAS]
        .EMPRS_RES_EXERC_FUT = rs![EMPRS_RES_EXERC_FUT]
        .EMPRS_CAPIT_SOC = rs![EMPRS_CAPIT_SOC]
        .EMPRS_RESERV_CAPIT_LUCRO = rs![EMPRS_RESERV_CAPIT_LUCRO]
        .EMPRS_RESERV_REAVAL = rs![EMPRS_RESERV_REAVAL]
        .EMPRS_PARTIC_MINOR = rs![EMPRS_PARTIC_MINOR]
        .EMPRS_LUCRO_PREJ_ACML = rs![EMPRS_LUCRO_PREJ_ACML]
 
     End With
     
     colBalanco.Add Item:=blc
    
     rs.MoveNext
     Loop Until rs.EOF
     
End If
rs.Close
conn.Close

'ALTERAÇÃO FEITA 28/05: preenchendo células, +alteração no count, +alteração na Sheet



cd_grupo = blc.CD_GRP
cd_cli = blc.cd_cli
CNPJ = blc.CNPJ
Layout = Front.ComboBox1.Text
     

ModuloBanco.trata_zeros
ModuloBanco.alimentacombobox

count2 = 3
count3 = 19
count4 = 36
count5 = 2

For Each blc In colBalanco

    '------------------------ATIVO--------------------------
     
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(7, count2).Value = blc.EMPRS_DISPS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(8, count2).Value = blc.EMPRS_CLI
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(9, count2).Value = blc.EMPRS_PROV_DEV_DUVIDS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(10, count2).Value = blc.EMPRS_ESTOQS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(11, count2).Value = blc.EMPRS_ADTO_FORN
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(12, count2).Value = blc.EMPRS_TIT_VAL_MOBIL
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(13, count2).Value = blc.EMPRS_DESP_PG_ANTEC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(14, count2).Value = blc.EMPRS_ATIVO_OUTRAS_CTAS_OPERAC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(15, count2).Value = blc.EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(17, count2).Value = blc.EMPRS_ATIVO_CC_CONTROL_COLIG
      'ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(18, count2).Value = blc.EMPRS_ATIVO_OUTRAS_CONTAS
      
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(20, count2).Value = blc.EMPRS_PART_CONTROL_COLIGS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(21, count2).Value = blc.EMPRS_OUTROS_INVEST
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(23, count2).Value = blc.EMPRS_IMOB_TECN_LIQ

      'O NOME DO CAMPO NÃO BATE COM O NOME DA CÉLULA, MAS FOI CONFERIDO NO DE-PARA
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(24, count2).Value = blc.EMPRS_ATIVO_INTANG
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(6, count2).Value = blc.DT_EXERC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(32, count2).Value = blc.EMPRS_RECT_BRUTA
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(33, count2).Value = blc.EMPRS_DEVOL_ABATIM
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(34, count2).Value = blc.EMPRS_IMPOS_FATRDS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(36, count2).Value = blc.EMPRS_CUSTO_PROD_VENDS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(37, count2).Value = blc.EMPRS_DEPREC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(39, count2).Value = blc.EMPRS_DESP_ADMINS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(40, count2).Value = blc.EMPRS_DESP_VNDAS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(41, count2).Value = blc.EMPRS_OUTRAS_DESP_REC_OPERAC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(42, count2).Value = blc.EMPRS_SALDO_COR_MONET
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(44, count2).Value = blc.EMPRS_RECT_FINANC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(45, count2).Value = blc.EMPRS_DESP_FINANC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(46, count2).Value = blc.EMPRS_VAR_CAMBL_LIQ
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(47, count2).Value = blc.EMPRS_RECT_DESP_NAO_OPERAC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(49, count2).Value = blc.EMPRS_EQUIV_PATRIOM
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(51, count2).Value = blc.EMPRS_IMP_RNDA_CONTRIB_SOC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(52, count2).Value = blc.EMPRS_PARTIC

    '------------------------PASSIVO------------------------

      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(7, count13).Value = blc.EMPRS_FORNS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(8, count13).Value = blc.EMPRS_OBRIG_SOC_TRIBUT
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(9, count13).Value = blc.EMPRS_ADTO_CLI
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(10, count13).Value = blc.EMPRS_EMPREST_FINANCS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(11, count13).Value = blc.EMPRS_DUPLIC_DESCTS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(12, count13).Value = blc.EMPRS_CAMB
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(13, count13).Value = blc.EMPRS_PASS_CC_CONTROL_COLIG
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(14, count13).Value = blc.EMPRS_PASS_OUTRAS_CTAS_OPERAC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(15, count13).Value = blc.EMPRS_PASS_OUTRAS_CTAS_NAO_OPER
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(17, count13).Value = blc.EMPRS_EMPREST_FINANCS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(18, count13).Value = blc.EMPRS_PASS_OUTRAS_CONTAS
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(20, count13).Value = blc.EMPRS_RES_EXERC_FUT
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(21, count13).Value = blc.EMPRS_CAPIT_SOC
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(22, count13).Value = blc.EMPRS_RESERV_CAPIT_LUCRO
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(23, count13).Value = blc.EMPRS_RESERV_REAVAL
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(24, count13).Value = blc.EMPRS_PARTIC_MINOR
      ActiveWorkbook.Worksheets("PJ_ReaisMil").Cells(25, count13).Value = blc.EMPRS_LUCRO_PREJ_ACML

    count2 = count2 + 2 'C
    count3 = count3 + 2 'S
    count4 = count4 + 2 'AJ
    count5 = count5 + 1
    
Next

colBalanco.count

'ALTERAÇÃO FEITA 28/05: +alteração na sheet, +alteração nas colunas.


  If colBalanco.count = 1 Then

    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("E:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("U:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("AL:AQ").EntireColumn.Hidden = True

   ElseIf colBalanco.count = 2 Then

    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("G:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("W:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("AN:AQ").EntireColumn.Hidden = True
    
   ElseIf colBalanco.count = 3 Then

    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("I:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("Y:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("PJ_ReaisMil").Columns("AP:AQ").EntireColumn.Hidden = True
    
   End If

End Sub


Public Sub Planilha_PJ_Fluxo()


Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1(6)
Dim colBalanco As Collection
Dim periodos As String
Dim count As Integer
Dim countR2 As Integer


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
     
     'ALTERAÇÃO FEITA 27/05: +NOVA PARAMETRIZAÇÃO
     With blc
        .DT_EXERC = rs![DT_EXERC]
        
        .BCO_EMPRS_APLIC_FINANC_LP = rs![BCO_EMPRS_APLIC_FINANC_LP]
        .BCO_EMPRS_CP_DERIVAT = rs![BCO_EMPRS_CP_DERIVAT]
        .BCO_EMPRS_LP_DERIVAT = rs![BCO_EMPRS_LP_DERIVAT]
        .BCO_EMPRS_DIVID_PAGOS = rs![BCO_EMPRS_DIVID_PAGOS]
        .BCO_EMPRS_DIVID_RECEB = rs![BCO_EMPRS_DIVID_RECEB]
     End With
     
     colBalanco.Add Item:=blc
    
     rs.MoveNext
     Loop Until rs.EOF
     
End If
rs.Close
conn.Close

countR2 = 4

'ALTERAÇÃO FEITA 27/05: +alteração na sheet, +alteração nas colunas.
For Each blc In colBalanco
     
    ActiveWorkbook.Worksheets("PJ_FLUXO").Cells(45, count2).Value = blc.BCO_EMPRS_APLIC_FINANC_LP
    ActiveWorkbook.Worksheets("PJ_FLUXO").Cells(42, count2).Value = blc.BCO_EMPRS_CP_DERIVAT
    ActiveWorkbook.Worksheets("PJ_FLUXO").Cells(46, count2).Value = blc.BCO_EMPRS_LP_DERIVAT
    ActiveWorkbook.Worksheets("PJ_FLUXO").Cells(65, count2).Value = blc.BCO_EMPRS_DIVID_PAGOS
    ActiveWorkbook.Worksheets("PJ_FLUXO").Cells(66, count2).Value = blc.BCO_EMPRS_DIVID_RECEB
   
    countR2 = countR2 + 1 'D
    
Next

colBalanco.count


 If colBalanco.count = 1 Then

    ActiveWorkbook.Worksheets("PJ_FLUXO").Columns("E:G").EntireColumn.Hidden = True

   ElseIf colBalanco.count = 2 Then

    ActiveWorkbook.Worksheets("PJ_FLUXO").Columns("F:G").EntireColumn.Hidden = True
    
   ElseIf colBalanco.count = 3 Then

    ActiveWorkbook.Worksheets("PJ_FLUXO").Columns("G").EntireColumn.Hidden = True
    
   End If



  
End Sub

Public Sub Planilha_PJ_EBITDA_AJUSTADO()


Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1(6)
Dim colBalanco As Collection
Dim periodos As String
'INICIO DAS ALTERAÇÕES 27/05
'ALTERAÇÃO 27/05 - DECLARAÇÃO DE VARIAVEL
Dim counta1 As Integer

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
     
     With blc
     
'ALTERAÇÃO FEITA 27/05 - INCREMENTO DE CÓDIGO DE ACORDO COM A TABELA DE MODELAGEM
        .BCO_EMPRS_AJUST1 = rs![BCO_EMPRS_AJUST1]
        .BCO_EMPRS_AJUST2 = rs![BCO_EMPRS_AJUST2]
        .BCO_EMPRS_AJUST3 = rs![BCO_EMPRS_AJUST3]
 
     End With
     
     colBalanco.Add Item:=blc
    
     rs.MoveNext
     Loop Until rs.EOF
     
End If
rs.Close
conn.Close

'ALTERAÇÃO FEITA 27/05 - MODIFICAÇÃO DE VARIAVEL PARA O CÓDIGO DO EBITDA
counta1 = 4

For Each blc In colBalanco
     
'ALTERAÇÃO FEITA 27/05 - INCREMENTO DE CÓDIGO DE ACORDO COM A TABELA DE MODELAGEM
    ActiveWorkbook.Worksheets("PJ_EBITDA AJUSTADO").Cells(12, counta1).Value = blc.BCO_EMPRS_AJUST1
    ActiveWorkbook.Worksheets("PJ_EBITDA AJUSTADO").Cells(13, counta1).Value = blc.BCO_EMPRS_AJUST2
    ActiveWorkbook.Worksheets("PJ_EBITDA AJUSTADO").Cells(14, counta1).Value = blc.BCO_EMPRS_AJUST3

End Sub
