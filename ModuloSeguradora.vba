Public Sub Planilha_SEGURADORA_ReaisMil()

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1(6)
Dim colBalanco As Collection
Dim periodos As String
Dim count As Integer
Dim count2 As Integer



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

If Not IsNull(rs) And rs.RecordCount > 0 Then
                       
   Do
          
     Dim blc As Balanco
     Set blc = New Balanco
     
    'ALTERAÇÃO FEITA 28/05: +NOVA PARAMETRIZAÇÃO

     With blc
        .DT_EXERC = rs![DT_EXERC]
      
      '------------------------ATIVO--------------------------
        
        .SEGUR_DISP = rs![SEGUR_DISP]
        .SEGUR_CRED_OPER_PREVID_COMPL = rs![SEGUR_CRED_OPER_PREVID_COMPL]
        .SEGUR_SEGURADORAS = rs![SEGUR_SEGURADORAS]
        .SEGUR_IRB = rs![SEGUR_IRB]
        .SEGUR_DESP_COMERC_DIFERD = rs![SEGUR_DESP_COMERC_DIFERD]
        .SEGUR_TITULO_VL_MBLRO = rs![SEGUR_TITULO_VL_MBLRO]
        .SEGUR_DESP_PAGTO_ANTCPO = rs![SEGUR_DESP_PAGTO_ANTCPO]
        .SEGUR_OUTRA_CONTA_OPER = rs![SEGUR_OUTRA_CONTA_OPER]
        .SEGUR_OUTRA_CONTA_NAO_OPER = rs![SEGUR_OUTRA_CONTA_NAO_OPER]

        .SEGUR_ATIV_CIRC = rs![SEGUR_ATIV_CIRC]

        .SEGUR_APLIC = rs![SEGUR_APLIC]
        .SEGUR_TITULO_CRED_RECEB = rs![SEGUR_TITULO_CRED_RECEB]

        .SEGUR_REALZV_LP = rs![SEGUR_REALZV_LP]

        .SEGUR_PART_CTRL_COLGD = rs![SEGUR_PART_CTRL_COLGD]
        .SEGUR_OUTRO_INVTMO = rs![SEGUR_OUTRO_INVTMO]

        .SEGUR_INVTMO = rs![SEGUR_INVTMO]

        .SEGUR_IMBRO_TECN_LIQ = rs![SEGUR_IMBRO_TECN_LIQ]

        .SEGUR_ATIV_DFRD = rs![SEGUR_ATIV_DFRD]

        .SEGUR_ATIV_PERMAN = rs![SEGUR_ATIV_PERMAN]

        .SEGUR_ATIV_TOTAL = rs![SEGUR_ATIV_TOTAL]

      '--------------DEMONSTRATIVO DE RESULTADO----------------

      .SEGUR_RENDA_CONTRIB = rs![SEGUR_RENDA_CONTRIB]
      .SEGUR_CONTRIB_RPS = rs![SEGUR_CONTRIB_RPS]
      .SEGUR_VAR_PROV_PREMIOS = rs![SEGUR_VAR_PROV_PREMIOS]

      .SEGUR_REC_OPER_LIQ = rs![SEGUR_REC_OPER_LIQ]

      .SEGUR_DESP_BENEF_RESGT = rs![SEGUR_DESP_BENEF_RESGT]
      .SEGUR_VAR_PROV_EVENTO_NAO_AVIS = rs![SEGUR_VAR_PROV_EVENTO_NAO_AVIS]

      .SEGUR_LCR_BRUTO = rs![SEGUR_LCR_BRUTO]

      .SEGUR_DESP_ADM = rs![SEGUR_DESP_ADM]
      .SEGUR_DESP_VDA = rs![SEGUR_DESP_VDA]
      .SEGUR_OUTRO_DESP_REC_OPER = rs![SEGUR_OUTRO_DESP_REC_OPER]
      .SEGUR_SALDO_CORREC_MONET = rs![SEGUR_SALDO_CORREC_MONET]

      .SEGUR_LCR_ANTES_RES_FIN = rs![SEGUR_LCR_ANTES_RES_FIN]

      .SEGUR_RECT_FIN = rs![SEGUR_RECT_FIN]
      .SEGUR_DESP_FIN = rs![SEGUR_DESP_FIN]
      .SEGUR_REC_DESP_NAO_OPER = rs![SEGUR_REC_DESP_NAO_OPER]

      .SEGUR_LCR_ANTES_EQUIV_PATRIM = rs![SEGUR_LCR_ANTES_EQUIV_PATRIM]

      .SEGUR_EQUIV_PATRIM = rs![SEGUR_EQUIV_PATRIM]

      .SEGUR_LCR_ANTES_IR = rs![SEGUR_LCR_ANTES_IR]

      .SEGUR_IR_RENDA_CONTRIB_SOC = rs![SEGUR_IR_RENDA_CONTRIB_SOC]
      .SEGUR_PARTICIP = rs![SEGUR_PARTICIP]

      .SEGUR_LCR_LIQ = rs![SEGUR_LCR_LIQ]

      '------------------------PASSIVO--------------------------

        .SEGUR_DEB_OPER_PREVID = rs![SEGUR_DEB_OPER_PREVID]
        .SEGUR_OBRIG_SOC_TRIB = rs![SEGUR_OBRIG_SOC_TRIB]
        .SEGUR_SINIS_LIQ = rs![SEGUR_SINIS_LIQ]
        .SEGUR_EMPREST_FIN = rs![SEGUR_EMPREST_FIN]
        .SEGUR_PROV_TECN = rs![SEGUR_PROV_TECN]
        .SEGUR_DEPOS_TERC = rs![SEGUR_DEPOS_TERC]
        .SEGUR_CTRL_COLGD = rs![SEGUR_CTRL_COLGD]
        .SEGUR_OUTRA_CONTA_OPER = rs![SEGUR_OUTRA_CONTA_OPER]
        .SEGUR_OUTRA_CONTA_NAO_OPER = rs![SEGUR_OUTRA_CONTA_NAO_OPER]

        .SEGUR_PASV_CIRC = rs![SEGUR_PASV_CIRC]

        .SEGUR_PROV_TECN = rs![SEGUR_PROV_TECN]
        .SEGUR_OUTRA_CONTA = rs![SEGUR_OUTRA_CONTA]

        .SEGUR_EXIG_LP = rs![SEGUR_EXIG_LP]

        .SEGUR_RES_EXERC_FUT = rs![SEGUR_RES_EXERC_FUT]

        .SEGUR_CAPITAL_SOC = rs![SEGUR_CAPITAL_SOC]
        .SEGUR_RES_CAPITAL_LCR = rs![SEGUR_RES_CAPITAL_LCR]
        .SEGUR_RES_REAVAL = rs![SEGUR_RES_REAVAL]
        .SEGUR_PARTICIP_MNTRO = rs![SEGUR_PARTICIP_MNTRO]
        .SEGUR_LCR_PREJ_ACUM = rs![SEGUR_LCR_PREJ_ACUM]

        .SEGUR_PATR_LIQ = rs![SEGUR_PATR_LIQ]

        .SEGUR_PASV_TOTAL = rs![SEGUR_PASV_TOTAL]
         

     End With
     
     colBalanco.Add Item:=blc
    
     rs.MoveNext
     Loop Until rs.EOF
     
End If
rs.Close
conn.Close

'ALTERAÇÃO FEITA 28/05: preenchendo células, +alteração no count, +alteração na Sheet

ModuloBanco.trata_zeros
ModuloBanco.alimentacombobox

count2 = 3
count3 = 19
count4 = 36
count5 = 2


For Each blc In colBalanco

    '------------------------ATIVO--------------------------

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(7, count2).Value = blc.SEGUR_DISP
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(8, count2).Value = blc.SEGUR_CRED_OPER_PREVID_COMPL
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(9, count2).Value = blc.SEGUR_SEGURADORAS
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(10, count2).Value = blc.SEGUR_IRB
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(11, count2).Value = blc.SEGUR_DESP_COMERC_DIFERD
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(12, count2).Value = blc.SEGUR_TITULO_VL_MBLRO
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(13, count2).Value = blc.SEGUR_DESP_PAGTO_ANTCPO
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(14, count2).Value = blc.SEGUR_OUTRA_CONTA_OPER
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(15, count2).Value = blc.SEGUR_OUTRA_CONTA_NAO_OPER

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(16, count2).Value = blc.SEGUR_ATIV_CIRC

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(17, count2).Value = blc.SEGUR_APLIC
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(18, count2).Value = blc.SEGUR_TITULO_CRED_RECEB

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(19, count2).Value = blc.SEGUR_REALZV_LP

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(20, count2).Value = blc.SEGUR_PART_CTRL_COLGD
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(21, count2).Value = blc.SEGUR_OUTRO_INVTMO

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(22, count2).Value = blc.SEGUR_INVTMO

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(23, count2).Value = blc.SEGUR_IMBRO_TECN_LIQ

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(24, count2).Value = blc.SEGUR_ATIV_DFRD

      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(25, count2).Value = blc.SEGUR_ATIV_PERMAN

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(26, count2).Value = blc.SEGUR_ATIV_TOTAL

    '--------------DEMONSTRATIVO DE RESULTADO----------------

      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(32, count2).Value = blc.SEGUR_RENDA_CONTRIB
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(33, count2).Value = blc.SEGUR_CONTRIB_RPS
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(34, count2).Value = blc.SEGUR_VAR_PROV_PREMIOS

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(35, count2).Value = blc.SEGUR_REC_OPER_LIQ

      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(36, count2).Value = blc.SEGUR_DESP_BENEF_RESGT
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(37, count2).Value = blc.SEGUR_VAR_PROV_EVENTO_NAO_AVIS

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(38, count2).Value = blc.SEGUR_LCR_BRUTO

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(39, count2).Value = blc.SEGUR_DESP_ADM
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(40, count2).Value = blc.SEGUR_DESP_VDA
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(41, count2).Value = blc.SEGUR_OUTRO_DESP_REC_OPER
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(42, count2).Value = blc.SEGUR_SALDO_CORREC_MONET

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(43, count2).Value = blc.SEGUR_LCR_ANTES_RES_FIN

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(44, count2).Value = blc.SEGUR_RECT_FIN
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(45, count2).Value = blc.SEGUR_DESP_FIN
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(46, count2).Value = blc.SEGUR_REC_DESP_NAO_OPER

      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(47, count2).Value = blc.SEGUR_LCR_ANTES_EQUIV_PATRIM

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(48, count2).Value = blc.SEGUR_EQUIV_PATRIM

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(49, count2).Value = blc.SEGUR_LCR_ANTES_IR

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(50, count2).Value = blc.SEGUR_IR_RENDA_CONTRIB_SOC
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(51, count2).Value = blc.SEGUR_PARTICIP

      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(52, count2).Value = blc.SEGUR_LCR_LIQ

    '------------------------PASSIVO------------------------

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(7, count3).Value = blc.SEGUR_DEB_OPER_PREVID
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(8, count3).Value = blc.SEGUR_OBRIG_SOC_TRIB
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(9, count3).Value = blc.SEGUR_SINIS_LIQ
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(10, count3).Value = blc.SEGUR_EMPREST_FIN
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(11, count3).Value = blc.SEGUR_PROV_TECN
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(12, count3).Value = blc.SEGUR_DEPOS_TERC
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(13, count3).Value = blc.SEGUR_CTRL_COLGD
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(14, count3).Value = blc.SEGUR_OUTRA_CONTA_OPER
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(15, count3).Value = blc.SEGUR_OUTRA_CONTA_NAO_OPER

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(16, count3).Value = blc.SEGUR_PASV_CIRC

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(17, count3).Value = blc.SEGUR_PROV_TECN
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(18, count3).Value = blc.SEGUR_OUTRA_CONTA

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(19, count3).Value = blc.SEGUR_EXIG_LP

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(20, count3).Value = blc.SEGUR_RES_EXERC_FUT

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(21, count3).Value = blc.SEGUR_CAPITAL_SOC
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(22, count3).Value = blc.SEGUR_RES_CAPITAL_LCR
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(23, count3).Value = blc.SEGUR_RES_REAVAL
      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(24, count3).Value = blc.SEGUR_PARTICIP_MNTRO
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(25, count3).Value = blc.SEGUR_LCR_PREJ_ACUM

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(26, count3).Value = blc.SEGUR_PATR_LIQ

      ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(27, count3).Value = blc.SEGUR_PASV_TOTAL
      
      'ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(47, count3).Value = blc.Bco_Cdi_liqdz_Dia


    

     ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(6, count2).Value = blc.DT_EXERC
     ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(6, count3).Value = blc.DT_EXERC
    
     ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Cells(2, 17).Value = blc.CD_GRP
    
    ' Inserir os dados da planilha auxiliar

     ActiveWorkbook.Worksheets("Aux").Cells(count5, 4).Value = blc.CD_GRP
     ActiveWorkbook.Worksheets("Aux").Cells(count5, 5).Value = blc.cd_cli
     ActiveWorkbook.Worksheets("Aux").Cells(count5, 7).Value = blc.FLG_GRP
     ActiveWorkbook.Worksheets("Aux").Cells(count5, 13).Value = blc.CNPJ
     
     
cd_grupo = blc.CD_GRP
cd_cli = blc.cd_cli
CNPJ = blc.CNPJ
Layout = Front.ComboBox1.Text
     
     
          
    count2 = count2 + 2 'C
    count3 = count3 + 2 'S
    count4 = count4 + 2 'AJ
    count5 = count5 + 1
    
Next

colBalanco.count


  If colBalanco.count = 1 Then

    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("E:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("U:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("AL:AQ").EntireColumn.Hidden = True

   ElseIf colBalanco.count = 2 Then

    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("G:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("W:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("AN:AQ").EntireColumn.Hidden = True
    
   ElseIf colBalanco.count = 3 Then

    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("I:J").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("Y:Z").EntireColumn.Hidden = True
    ActiveWorkbook.Worksheets("SEGURADORA_ReaisMil").Columns("AP:AQ").EntireColumn.Hidden = True
    
   End If


End Sub
