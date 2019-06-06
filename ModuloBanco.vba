
Public cd_grupo As String
Public cd_cli As String
Public CNPJ As String
Public FLG_GRP As String
Public Layout As String
Public NM_EMP As String
Public periodos As String

Dim colBalanco As Collection
Dim blc As Balanco

Public Sub alimenta_combobox()

    'Montando Lista de Auditor

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim j As Integer
   
    Set conn = getConnection()
    Set rs = New ADODB.Recordset
             
    conn.Open
   
    qry = "select distinct DS_auditor from LB_PLANI.DIM_AUDITOR"
       
    rs.Open qry, conn, adOpenStatic
               
        If rs.RecordCount > 0 Then
        
            With Sheet6.ComboBox1
                .Clear
                Do
                    .ColumnCount = 1
                    .ColumnWidths = "60"
                    .AddItem
                    '.List(j, 0) = rs![DT_CRG]
                    .list(j, 0) = rs![DS_AUDITOR]
                   
                    j = j + 1
                   
                    rs.MoveNext
                   
                Loop Until rs.EOF
            End With
           
        Else
            Sheet6.ComboBox1.Clear
            MsgBox " Informe o Auditor"
        End If


End Sub

Public Sub Planilha_Bancos()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim arr1(6)
    Set colBalanco = New Collection

    Set conn = getConnection()
    Set rs = New ADODB.Recordset
  
    If periodos <> "" Then
  
    conn.Open

    'qry = "select * from LB_PLANI.FATO_balanco where dt_exerc in (" & periodos & ") and cd_cli = " & cd_cli
    qry = "select * from LB_PLANI.FATO_balanco where dt_exerc in (" & periodos & ") and cd_cli = " & cd_cli
    qry = qry & " And cat(DT_CRG, dt_exerc) in (Select cat(Max(DT_CRG), dt_exerc) From LB_PLANI.FATO_balanco "
    qry = qry & " Where dt_exerc in (" & periodos & ") and cd_cli = " & cd_cli & " Group By dt_exerc)"

    rs.Open qry, conn, adOpenStatic

    If Not IsNull(rs) And rs.RecordCount > 0 Then
        Do
            'Dim blc As Balanco
            Set blc = New Balanco
    
            With blc
                ' BANCOS MIL
                .cd_cli = rs![cd_cli]
                .CD_GRP = rs![CD_GRP]
                .FLG_GRP = rs![FLG_GRP]
                .CNPJ = rs![CNPJ]
                .DT_EXERC = rs![DT_EXERC]
                .MES_DE_FECHAMENTO = rs![MES_DE_FECHAMENTO]
                .Bco_Ativo_Disp = rs![Bco_Ativo_Disp]
                .Bco_Ativo_Cdi = rs![Bco_Ativo_Cdi]
                .Bco_Ativo_Titulo_Merc_Abert = rs![Bco_Ativo_Titulo_Merc_Abert]
                .Bco_Ativo_Tvm = rs![Bco_Ativo_Tvm]
                .Bco_Ativo_Operac_Cred = rs![Bco_Ativo_Operac_Cred]
                .Bco_Ativo_Pdd = rs![Bco_Ativo_Pdd]
                .Bco_Ativo_Op_Arrend_Mercatl = rs![Bco_Ativo_Op_Arrend_Mercatl]
                .Bco_Ativo_Pdd_Arrend_Mercatl = rs![Bco_Ativo_Pdd_Arrend_Mercatl]
                .Bco_Ativo_Desp_Antec = rs![Bco_Ativo_Desp_Antec]
                .Bco_Ativo_Cart_Camb = rs![Bco_Ativo_Cart_Camb]
                .Bco_Ativo_Outros_Creds = rs![Bco_Ativo_Outros_Creds]
                .Bco_Ativo_Circ_Tvm = rs![Bco_Ativo_Circ_Tvm]
                .Bco_Ativo_Circ_Operac_Cred = rs![Bco_Ativo_Circ_Operac_Cred]
                .Bco_Ativo_Circ_Pdd_op_Cred = rs![Bco_Ativo_Circ_Pdd_op_Cred]
                .Bco_Ativo_Circ_Op_arrend_merc = rs![Bco_Ativo_Circ_Op_arrend_merc]
                .Bco_Ativo_Circ_Pdd_op_Arr_merc = rs![Bco_Ativo_Circ_Pdd_op_Arr_merc]
                .Bco_Ativo_Circ_Outros_Cred = rs![Bco_Ativo_Circ_Outros_Cred]
                .Bco_Atv_N_Circ_Part_Ctrl_Colig = rs![Bco_Atv_N_Circ_Part_Ctrl_Colig]
                .Bco_Atv_N_Circ_Outros_Invest = rs![Bco_Atv_N_Circ_Outros_Invest]
                .Bco_Atv_N_Circ_Invest = rs![Bco_Atv_N_Circ_Invest]
                .Bco_Atv_N_Circ_Imob_Tec_Liq = rs![Bco_Atv_N_Circ_Imob_Tec_Liq]
                .Bco_Atv_N_Circ_Atv_Intang = rs![Bco_Atv_N_Circ_Atv_Intang]
                .Bco_Dre_Operac_Cred = rs![Bco_Dre_Operac_Cred]
                .Bco_Dre_Tvm = rs![Bco_Dre_Tvm]
                .Bco_Dre_Outras_Rec_Interm = rs![Bco_Dre_Outras_Rec_Interm]
                .Bco_Dre_Capt_Merc = rs![Bco_Dre_Capt_Merc]
                .Bco_Dre_Empr_Cess_Repass = rs![Bco_Dre_Empr_Cess_Repass]
                .Bco_Dre_Outras_Desp_Interm = rs![Bco_Dre_Outras_Desp_Interm]
                .Bco_Dre_Const_Pdd = rs![Bco_Dre_Const_Pdd]
                .Bco_Dre_Rect_Prest_Serv = rs![Bco_Dre_Rect_Prest_Serv]
                .Bco_Dre_Custo_Operac = rs![Bco_Dre_Custo_Operac]
                .Bco_Dre_Desp_Tribut = rs![Bco_Dre_Desp_Tribut]
                .Bco_Dre_Outras_Rect_Desp_Operac = rs![Bco_Dre_Outras_Rect_Desp_Operac]
                .Bco_Dre_Equiv_Patrim = rs![Bco_Dre_Equiv_Patrim]
                .Bco_Dre_Rect_Desp_N_Operac = rs![Bco_Dre_Rect_Desp_N_Operac]
                .Bco_Dre_Impst_Renda_Ctrl_Soc = rs![Bco_Dre_Impst_Renda_Ctrl_Soc]
                .Bco_Dre_Part = rs![Bco_Dre_Part]
                
                .Bco_Pass_Depos_Avista = rs![Bco_Pass_Depos_Avista]
                .Bco_Pass_Poupanca = rs![Bco_Pass_Poupanca]
                .Bco_Pass_Depos_Interfinan = rs![Bco_Pass_Depos_Interfinan]
                .Bco_Pass_Depos_Aprazo = rs![Bco_Pass_Depos_Aprazo]
                .Bco_Pass_Capt_Merc_Abert = rs![Bco_Pass_Capt_Merc_Abert]
                .Bco_Pass_Emprest_Pais = rs![Bco_Pass_Emprest_Pais]
                .Bco_Pass_Repass_Pais = rs![Bco_Pass_Repass_Pais]
                .Bco_Pass_Emprest_Exterior = rs![Bco_Pass_Emprest_Exterior]
                .Bco_Pass_Repass_Exterior = rs![Bco_Pass_Repass_Exterior]
                .Bco_Pass_Cart_Camb = rs![Bco_Pass_Cart_Camb]
                .Bco_Pass_Outras_Contas = rs![Bco_Pass_Outras_Contas]
                .Bco_Pass_Depos = rs![Bco_Pass_Depos]
                .Bco_Pass_Circ_Emprest_Pais = rs![Bco_Pass_Circ_Emprest_Pais]
                .Bco_Pass_Circ_Repass_Pais = rs![Bco_Pass_Circ_Repass_Pais]
                .Bco_Pass_Circ_Emprest_Exterior = rs![Bco_Pass_Circ_Emprest_Exterior]
                .Bco_Pass_Circ_Repass_Exterior = rs![Bco_Pass_Circ_Repass_Exterior]
                .Bco_Pass_Circ_Outras_Contas = rs![Bco_Pass_Circ_Outras_Contas]
                .Bco_Pass_N_Circ_Capit_Soc = rs![Bco_Pass_N_Circ_Capit_Soc]
                .Bco_Pass_N_Circ_Reserv_Capt = rs![Bco_Pass_N_Circ_Reserv_Capt]
                .Bco_Pass_N_Circ_Part_Minor = rs![Bco_Pass_N_Circ_Part_Minor]
                .Bco_Pass_N_Circ_Ajust_Vlr_Merc = rs![Bco_Pass_N_Circ_Ajust_Vlr_Merc]
                .Bco_Pass_N_Circ_Lcr_Prej_Acml = rs![Bco_Pass_N_Circ_Lcr_Prej_Acml]
                .Bco_Bxdos_Sectzdos = rs![Bco_Bxdos_Sectzdos]
                .Bco_Ind_Basileia_Br = rs![Bco_Ind_Basileia_Br]
                .Bco_Basileia_Tier_I = rs![Bco_Basileia_Tier_I]
                .Bco_Dpge_I = rs![Bco_Dpge_I]
                .Bco_Dpge_II = rs![Bco_Dpge_II]
                .Bco_Avais_Fiancas_Prestdos = rs![Bco_Avais_Fiancas_Prestdos]
                .Bco_Ag = rs![Bco_Ag]
                .Bco_Func = rs![Bco_Func]
                .Bco_Fnds_Admn = rs![Bco_Fnds_Admn]
                .Bco_Cred_Trib = rs![Bco_Cred_Trib]
                .Bco_Cdi_liqdz_Dia = rs![Bco_Cdi_liqdz_Dia]
                .Bco_Capt_Merc_Aber = rs![Bco_Capt_Merc_Aber]
                .Bco_Div_Subord = rs![Bco_Div_Subord]
                .Bco_Instrm_Fin_Deriv = rs![Bco_Instrm_Fin_Deriv]
                .Bco_Depos_Aprazo = rs![Bco_Depos_Aprazo]
                .Bco_Depos_Interfin = rs![Bco_Depos_Interfin]
                 
                ' BANCOS TVM
                .Bco_P_negoc_Valor_Custo = rs![Bco_P_negoc_Valor_Custo]
                .Bco_P_negoc_Valor_Contab = rs![Bco_P_negoc_Valor_Contab]
                .Bco_P_negoc_Mtm = rs![Bco_P_negoc_Mtm]
                .Bco_Disp_Venda_Valor_Custo = rs![Bco_Disp_Venda_Valor_Custo]
                .Bco_Disp_Venda_Valor_Contab = rs![Bco_Disp_Venda_Valor_Contab]
                .Bco_Disp_Venda_Mtm = rs![Bco_Disp_Venda_Mtm]
                .Bco_Mtdos_Vcto_Valor_Custo = rs![Bco_Mtdos_Vcto_Valor_Custo]
                .Bco_Mtdos_Vcto_Valor_Contab = rs![Bco_Mtdos_Vcto_Valor_Contab]
                .Bco_Mtdos_Vcto_Mtm = rs![Bco_Mtdos_Vcto_Mtm]
                .Bco_Instr_Financ_Deriv_Vlr_custo = rs![Bco_Instr_Financ_Deriv_Vlr_custo]
                .Bco_Instr_Fin_Deriv_Vlr_Contab = rs![Bco_Instr_Fin_Deriv_Vlr_Contab]
                .Bco_Instr_Financ_Deriv_Mtm = rs![Bco_Instr_Financ_Deriv_Mtm]
                .Bco_Disp_Venda_Prov_P_Desv = rs![Bco_Disp_Venda_Prov_P_Desv]
                .Bco_Instr_Fin_Deriv_Prov_P_Desv = rs![Bco_Instr_Fin_Deriv_Prov_P_Desv]
                .Bco_Mtdos_Vcto_Prov_P_Desv = rs![Bco_Mtdos_Vcto_Prov_P_Desv]
                .Bco_P_Negoc_Prov_P_Desv = rs![Bco_P_Negoc_Prov_P_Desv]
                
                'BANCOS CONTIGENCIA
                .Bco_Civeis_Conting_Provs = rs![Bco_Civeis_Conting_Provs]
                .Bco_Trablstas_Conting_Provs = rs![Bco_Trablstas_Conting_Provs]
                .Bco_Fiscais_Conting_Provs = rs![Bco_Fiscais_Conting_Provs]
                .Bco_Total_Conting_Provs = rs![Bco_Total_Conting_Provs]
                .Bco_Civeis_Depos_Judc = rs![Bco_Civeis_Depos_Judc]
                .Bco_Trablstas_Depos_Judc = rs![Bco_Trablstas_Depos_Judc]
                .Bco_Fiscais_Depos_Judc = rs![Bco_Fiscais_Depos_Judc]
                .Bco_Total_Depos_Judc = rs![Bco_Total_Depos_Judc]
                .Bco_Civeis_Conting_Nao_Provs = rs![Bco_Civeis_Conting_Nao_Provs]
                .Bco_Trablstas_Conting_Nao_Provs = rs![Bco_Trablstas_Conting_Nao_Provs]
                .Bco_Fiscais_Conting_Nao_Provs = rs![Bco_Fiscais_Conting_Nao_Provs]
                .Bco_Total_Conting_Nao_Provs = rs![Bco_Total_Conting_Nao_Provs]
                
                'BANCOS CARTEIRA
                .BCO_AA = rs![BCO_AA]
                .BCO_A = rs![BCO_A]
                .BCO_B = rs![BCO_B]
                .BCO_C = rs![BCO_C]
                .BCO_D = rs![BCO_D]
                .BCO_E = rs![BCO_E]
                .BCO_F = rs![BCO_F]
                .BCO_G = rs![BCO_G]
                .BCO_H = rs![BCO_H]
                .BCO_TOTAL_CART = rs![BCO_TOTAL_CART]
                .BCO_PDD_CONST = rs![BCO_PDD_CONST]
                .BCO_VENCD = rs![BCO_VENCD]
                .BCO_VENCD_90D = rs![BCO_VENCD_90D]
                .BCO_PDD_CARACT_CRED = rs![BCO_PDD_CARACT_CRED]
                .BCO_PDD_AVAIS_FIANCAS = rs![BCO_PDD_AVAIS_FIANCAS]
                .BCO_PDD_CART_EXPAND = rs![BCO_PDD_CART_EXPAND]
                
                'BANCOS PDD
                .BCO_SALDO_INICIAL = rs![BCO_SALDO_INICIAL]
                .BCO_CONST = rs![BCO_CONST]
                .BCO_REVERSAO = rs![BCO_REVERSAO]
                .BCO_BAIXAS = rs![BCO_BAIXAS]
                .BCO_RENEG_FLUXO = rs![BCO_RENEG_FLUXO]
                .BCO_RENEG_ESTOQ = rs![BCO_RENEG_ESTOQ]
                .BCO_RECUP = rs![BCO_RECUP]
                
                'Bancos Funding Liquidez Ativos
                .BCO_LCA = rs![BCO_LCA]
                .BCO_LCI = rs![BCO_LCI]
                .BCO_LF = rs![BCO_LF]
                .BCO_LETRA_CMBIO = rs![BCO_LETRA_CMBIO]
                .BCO_LFSN = rs![BCO_LFSN]
                .BCO_CREDORES_CRED_C_OBRIG = rs![BCO_CREDORES_CRED_C_OBRIG]
                .BCO_OUTROS = rs![BCO_OUTROS]
                .BCO_PARTS_RELAC = rs![BCO_PARTS_RELAC]
                .BCO_TVM_VINC_PREST_GAR_NEG = rs![BCO_TVM_VINC_PREST_GAR_NEG]
                .BCO_TVM_BAIXA_LIQDZ = rs![BCO_TVM_BAIXA_LIQDZ]
                .Bco_Tvm_Caract_Cred_Neg = rs![Bco_Tvm_Caract_Cred_Neg]
                .BCO_DEPOS_JUDC = rs![BCO_DEPOS_JUDC]
                .BCO_BNDU = rs![BCO_BNDU]
                
                'Bancos Rentabilidade
                .BCO_NIM_AJUST_CLI = rs![BCO_NIM_AJUST_CLI]
                .BCO_EFICIENCY_RATIO_AJUS_CLI = rs![BCO_EFICIENCY_RATIO_AJUS_CLI]
                
            End With
         
            colBalanco.Add Item:=blc
            
            rs.MoveNext
        Loop Until rs.EOF
        
        cd_grupo = blc.CD_GRP
        cd_cli = blc.cd_cli
        CNPJ = blc.CNPJ
        FLG_GRP = blc.FLG_GRP
    End If
    rs.Close
    conn.Close
    End If

    Planilha_Bancos_Mil
    Planilha_Bancos_PDD
    Planilha_Bancos_Funding
    Planilha_Bancos_Carteira
    Planilha_Bancos_Rentabilidada
    Planilha_Bancos_Contingencias
    Planilha_Bancos_TVM
    Planilha_Bancos_CGP_Alavancagem

End Sub

Public Sub Planilha_Bancos_CGP_Alavancagem()

    ActiveWorkbook.Worksheets("BANCOS_CGP e Alavancagem").Columns("C:F").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_CGP e Alavancagem").Columns("C:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_CGP e Alavancagem").Columns("D:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_CGP e Alavancagem").Columns("E:E").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_CGP e Alavancagem").Columns("F").EntireColumn.Hidden = True
    End If

End Sub

Public Sub Limpa_Planilha_Bancos_Mil()

    'Auditor
    Sheet6.ComboBox1.Text = ""
    'Planilhador
    Sheet6.Range("S3").Value = ""
    
    count2 = 3
    count3 = 19
    count4 = 36
    count5 = 2

    For i = 1 To 7
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(5, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(7, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(8, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(9, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(11, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(12, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(13, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(14, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(15, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(16, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(17, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(19, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(20, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(21, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(22, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(23, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(24, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(26, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(27, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(28, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(29, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(30, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(38, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(39, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(40, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(42, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(43, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(44, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(46, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(48, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(49, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(50, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(51, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(53, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(55, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(57, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(58, count2).Value = ""
    
        '------Passivo-------
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(7, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(8, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(9, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(11, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(12, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(13, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(14, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(15, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(16, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(17, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(19, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(20, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(21, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(22, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(23, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(24, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(26, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(27, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(28, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(29, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(30, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(37, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(38, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(39, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(40, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(41, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(42, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(43, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(44, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(45, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(46, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(47, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(7, count4).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(8, count4).Value = ""
         ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(9, count4).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count4).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(6, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(6, count3).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(2, 17).Value = ""

        ' Inserir os dados da planilha auxiliar
'        ActiveWorkbook.Worksheets("Aux").Cells(count5, 4).Value = blc.CD_GRP
'        ActiveWorkbook.Worksheets("Aux").Cells(count5, 5).Value = blc.cd_cli
'        ActiveWorkbook.Worksheets("Aux").Cells(count5, 7).Value = blc.FLG_GRP
'        ActiveWorkbook.Worksheets("Aux").Cells(count5, 13).Value = blc.CNPJ
         
        count2 = count2 + 2 'C
        count3 = count3 + 2 'S
        count4 = count4 + 2 'AJ
        count5 = count5 + 1
    Next

    cd_grupo = ""
    cd_cli = ""
    CNPJ = ""
    Layout = ""

End Sub

Public Sub Planilha_Bancos_Mil()

    'Auditor
    alimenta_combobox
    'Planilhador
    Sheet6.Range("S3").Value = Application.UserName
    
    count2 = 3
    count3 = 19
    count4 = 36
    count5 = 2

    For Each blc In colBalanco
    
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(5, count2).Value = blc.MES_DE_FECHAMENTO
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(7, count2).Value = blc.Bco_Ativo_Disp
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(8, count2).Value = blc.Bco_Ativo_Cdi
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(9, count2).Value = blc.Bco_Ativo_Titulo_Merc_Abert
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count2).Value = blc.Bco_Ativo_Tvm
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(11, count2).Value = blc.Bco_Ativo_Operac_Cred
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(12, count2).Value = blc.Bco_Ativo_Pdd
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(13, count2).Value = blc.Bco_Ativo_Op_Arrend_Mercatl
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(14, count2).Value = blc.Bco_Ativo_Pdd_Arrend_Mercatl
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(15, count2).Value = blc.Bco_Ativo_Desp_Antec
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(16, count2).Value = blc.Bco_Ativo_Cart_Camb
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(17, count2).Value = blc.Bco_Ativo_Outros_Creds
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(19, count2).Value = blc.Bco_Ativo_Circ_Tvm
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(20, count2).Value = blc.Bco_Ativo_Circ_Operac_Cred
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(21, count2).Value = blc.Bco_Ativo_Circ_Pdd_op_Cred
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(22, count2).Value = blc.Bco_Ativo_Circ_Op_arrend_merc
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(23, count2).Value = blc.Bco_Ativo_Circ_Pdd_op_Arr_merc
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(24, count2).Value = blc.Bco_Ativo_Circ_Outros_Cred
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(26, count2).Value = blc.Bco_Atv_N_Circ_Part_Ctrl_Colig
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(27, count2).Value = blc.Bco_Atv_N_Circ_Outros_Invest
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(28, count2).Value = blc.Bco_Atv_N_Circ_Invest
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(29, count2).Value = blc.Bco_Atv_N_Circ_Imob_Tec_Liq
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(30, count2).Value = blc.Bco_Atv_N_Circ_Atv_Intang
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(38, count2).Value = blc.Bco_Dre_Operac_Cred
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(39, count2).Value = blc.Bco_Dre_Tvm
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(40, count2).Value = blc.Bco_Dre_Outras_Rec_Interm
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(42, count2).Value = blc.Bco_Dre_Capt_Merc
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(43, count2).Value = blc.Bco_Dre_Empr_Cess_Repass
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(44, count2).Value = blc.Bco_Dre_Outras_Desp_Interm
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(46, count2).Value = blc.Bco_Dre_Const_Pdd
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(48, count2).Value = blc.Bco_Dre_Rect_Prest_Serv
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(49, count2).Value = blc.Bco_Dre_Custo_Operac
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(50, count2).Value = blc.Bco_Dre_Desp_Tribut
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(51, count2).Value = blc.Bco_Dre_Outras_Rect_Desp_Operac
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(53, count2).Value = blc.Bco_Dre_Equiv_Patrim
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(55, count2).Value = blc.Bco_Dre_Rect_Desp_N_Operac
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(57, count2).Value = blc.Bco_Dre_Impst_Renda_Ctrl_Soc
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(58, count2).Value = blc.Bco_Dre_Part
    
        '------Passivo-------
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(7, count3).Value = blc.Bco_Pass_Depos_Avista
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(8, count3).Value = blc.Bco_Pass_Poupanca
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(9, count3).Value = blc.Bco_Pass_Depos_Interfinan
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count3).Value = blc.Bco_Pass_Depos_Aprazo
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(11, count3).Value = blc.Bco_Pass_Capt_Merc_Abert
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(12, count3).Value = blc.Bco_Pass_Emprest_Pais
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(13, count3).Value = blc.Bco_Pass_Repass_Pais
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(14, count3).Value = blc.Bco_Pass_Emprest_Exterior
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(15, count3).Value = blc.Bco_Pass_Repass_Exterior
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(16, count3).Value = blc.Bco_Pass_Cart_Camb
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(17, count3).Value = blc.Bco_Pass_Outras_Contas
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(19, count3).Value = blc.Bco_Pass_Depos
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(20, count3).Value = blc.Bco_Pass_Circ_Emprest_Pais
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(21, count3).Value = blc.Bco_Pass_Circ_Repass_Pais
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(22, count3).Value = blc.Bco_Pass_Circ_Emprest_Exterior
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(23, count3).Value = blc.Bco_Pass_Circ_Repass_Exterior
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(24, count3).Value = blc.Bco_Pass_Circ_Outras_Contas
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(26, count3).Value = blc.Bco_Pass_N_Circ_Capit_Soc
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(27, count3).Value = blc.Bco_Pass_N_Circ_Reserv_Capt
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(28, count3).Value = blc.Bco_Pass_N_Circ_Part_Minor
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(29, count3).Value = blc.Bco_Pass_N_Circ_Ajust_Vlr_Merc
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(30, count3).Value = blc.Bco_Pass_N_Circ_Lcr_Prej_Acml
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(37, count3).Value = blc.Bco_Bxdos_Sectzdos
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(38, count3).Value = blc.Bco_Ind_Basileia_Br
        'ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(38, count3).Value = blc.Bco_Ind_Basileia_Br
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(39, count3).Value = blc.Bco_Basileia_Tier_I
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(40, count3).Value = blc.Bco_Dpge_I
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(41, count3).Value = blc.Bco_Dpge_II
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(42, count3).Value = blc.Bco_Avais_Fiancas_Prestdos
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(43, count3).Value = blc.Bco_Ag
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(44, count3).Value = blc.Bco_Func
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(45, count3).Value = blc.Bco_Fnds_Admn
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(46, count3).Value = blc.Bco_Cred_Trib
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(47, count3).Value = blc.Bco_Cdi_liqdz_Dia
        'ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(50, count4).Value = blc.Bco_Capt_Merc_Aber
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(7, count4).Value = blc.Bco_Capt_Merc_Aber
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(8, count4).Value = blc.Bco_Div_Subord
        'ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(54, count3).Value = blc.Bco_Instrm_Fin_Deriv
         ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(9, count4).Value = blc.Bco_Instrm_Fin_Deriv
        'ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(57, count3).Value = blc.Bco_Depos_Aprazo
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count4).Value = blc.Bco_Depos_Aprazo
        'ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(58, count3).Value = blc.Bco_Depos_Interfin
        'ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(10, count3).Value = blc.Bco_Depos_Interfin
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(6, count2).Value = blc.DT_EXERC
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(6, count3).Value = blc.DT_EXERC
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Cells(2, 17).Value = blc.CD_GRP

        ' Inserir os dados da planilha auxiliar
        ActiveWorkbook.Worksheets("Aux").Cells(count5, 4).Value = blc.CD_GRP
        ActiveWorkbook.Worksheets("Aux").Cells(count5, 5).Value = blc.cd_cli
        ActiveWorkbook.Worksheets("Aux").Cells(count5, 7).Value = blc.FLG_GRP
        ActiveWorkbook.Worksheets("Aux").Cells(count5, 13).Value = blc.CNPJ
         
'        cd_grupo = blc.CD_GRP
'        cd_cli = blc.cd_cli
'        CNPJ = blc.CNPJ
'        Layout = Front.ComboBox1.Text
     
        count2 = count2 + 2 'C
        count3 = count3 + 2 'S
        count4 = count4 + 2 'AJ
        count5 = count5 + 1
    Next
    
    ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("C:J").EntireColumn.Hidden = False
    ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("S:Z").EntireColumn.Hidden = False
    ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("AJ:AQ").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("C:J").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("S:Z").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("AJ:AQ").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("E:J").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("U:Z").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("AL:AQ").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("G:J").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("W:Z").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("AN:AQ").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("I:J").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("Y:Z").EntireColumn.Hidden = True
        ActiveWorkbook.Worksheets("BANCOS_Mil_v5").Columns("AP:AQ").EntireColumn.Hidden = True
    End If

End Sub

Public Sub Limpa_Planilha_Contingencias()

    count2 = 3
    
    For i = 1 To 7
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(6, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(7, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(8, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(5, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(10, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(11, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(12, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(9, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(14, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(15, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(16, count2).Value = ""
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(13, count2).Value = ""
            
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(14, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(15, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(16, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(13, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(10, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(11, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(12, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(9, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(6, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(7, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(8, count2).Interior.Color = vbWhite
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(5, count2).Interior.Color = vbWhite
    
        count2 = count2 + 2
        count3 = count1 + 1
    
    Next i

End Sub

Public Sub Planilha_Bancos_Contingencias()

    count2 = 3
    
    For Each blc In colBalanco
       
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(6, count2).Value = blc.Bco_Civeis_Conting_Provs
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(7, count2).Value = blc.Bco_Trablstas_Conting_Provs
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(8, count2).Value = blc.Bco_Fiscais_Conting_Provs
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(5, count2).Value = blc.Bco_Total_Conting_Provs
    
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(10, count2).Value = blc.Bco_Civeis_Depos_Judc
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(11, count2).Value = blc.Bco_Trablstas_Depos_Judc
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(12, count2).Value = blc.Bco_Fiscais_Depos_Judc
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(9, count2).Value = blc.Bco_Total_Depos_Judc
    
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(14, count2).Value = blc.Bco_Civeis_Conting_Nao_Provs
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(15, count2).Value = blc.Bco_Trablstas_Conting_Nao_Provs
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(16, count2).Value = blc.Bco_Fiscais_Conting_Nao_Provs
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(13, count2).Value = blc.Bco_Total_Conting_Nao_Provs
            
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(14, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(15, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(16, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(13, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(10, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(11, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(12, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(9, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(6, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(7, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(8, count2).Interior.Color = vbYellow
            ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(5, count2).Interior.Color = vbYellow
            
            'ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(4, count2).Value = blc.DT_EXERC
           ' ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Cells(6, count3).Value = blc.DT_EXERC
    
        count2 = count2 + 2
        count3 = count1 + 1
    
    Next
    
    colBalanco.count

    ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Columns("C:J").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Columns("C:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Columns("E:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Columns("G:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_CONTINGENCIAS").Columns("I:J").EntireColumn.Hidden = True
    End If

End Sub

Public Sub Limpa_Planilha_Funding()

    count2 = 3

    For i = 1 To 7
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(11, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(12, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(14, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(13, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(15, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(20, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(21, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(25, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(35, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(36, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(37, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(57, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(58, count2).Value = ""
        
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(11, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(12, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(14, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(13, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(15, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(20, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(21, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(25, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(35, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(36, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(37, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(57, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(58, count2).Interior.Color = vbWhite
       
        count2 = count2 + 1 'C
    Next i
    
End Sub

Public Sub Planilha_Bancos_Funding()

    count2 = 3

    For Each blc In colBalanco
        'If countFor = 1 Then
        '    ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(1, 3).Value = blc.CD_GRP & " / " & blc.cd_cli
        'End If
        
        'ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(5, count2).Value = blc.DT_EXERC
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(11, count2).Value = blc.BCO_LCA
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(12, count2).Value = blc.BCO_LCI
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(14, count2).Value = blc.BCO_LETRA_CMBIO
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(13, count2).Value = blc.BCO_LF
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(15, count2).Value = blc.BCO_LFSN
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(20, count2).Value = blc.BCO_CREDORES_CRED_C_OBRIG
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(21, count2).Value = blc.BCO_OUTROS
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(25, count2).Value = blc.BCO_PARTS_RELAC
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(35, count2).Value = blc.BCO_TVM_VINC_PREST_GAR_NEG
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(36, count2).Value = blc.BCO_TVM_BAIXA_LIQDZ
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(37, count2).Value = blc.Bco_Tvm_Caract_Cred_Neg
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(57, count2).Value = blc.BCO_DEPOS_JUDC
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(58, count2).Value = blc.BCO_BNDU
        
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(11, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(12, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(14, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(13, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(15, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(20, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(21, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(25, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(35, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(36, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(37, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(57, count2).Interior.Color = vbMagenta
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Cells(58, count2).Interior.Color = vbMagenta
       
        count2 = count2 + 1 'C
    Next
    
    ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Columns("C:F").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Columns("C:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Columns("D:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Columns("E:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_Funding_Liquidez_Ativos").Columns("F").EntireColumn.Hidden = True
    End If
    
End Sub

Public Sub Limpa_Planilha_PDD()
    
    countR2 = 3

    For i = 1 To 7
        If countR2 = 3 Then
            ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(5, countR2).Value = ""
        End If
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(6, countR2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(7, countR2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(11, countR2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(12, countR2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(13, countR2).Value = ""
        
        If countR2 = 3 Then
            ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(5, countR2).Interior.Color = vbWhite
        End If
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(6, countR2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(7, countR2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(11, countR2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(12, countR2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(13, countR2).Interior.Color = vbWhite

        countR2 = countR2 + 1 'C
    Next i

End Sub

Public Sub Planilha_Bancos_PDD()

    countR2 = 3

    For Each blc In colBalanco
        If countR2 = 3 Then
            ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(5, countR2).Value = blc.BCO_SALDO_INICIAL
        End If
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(6, countR2).Value = blc.BCO_CONST
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(7, countR2).Value = blc.BCO_REVERSAO
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(11, countR2).Value = blc.BCO_RENEG_FLUXO
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(12, countR2).Value = blc.BCO_RENEG_ESTOQ
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(13, countR2).Value = blc.BCO_RECUP
        
        If countR2 = 3 Then
            ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(5, countR2).Interior.Color = vbGreen
        End If
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(6, countR2).Interior.Color = vbGreen
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(7, countR2).Interior.Color = vbGreen
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(11, countR2).Interior.Color = vbGreen
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(12, countR2).Interior.Color = vbGreen
        ActiveWorkbook.Worksheets("BANCOS_PDD").Cells(13, countR2).Interior.Color = vbGreen

        countR2 = countR2 + 1 'C
    Next

    ActiveWorkbook.Worksheets("BANCOS_PDD").Columns("C:F").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_PDD").Columns("C:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_PDD").Columns("D:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_PDD").Columns("E:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_PDD").Columns("F").EntireColumn.Hidden = True
    End If
   
End Sub

Public Sub Limpa_Planilha_Rentabilidada()

    countR2 = 3

    For i = 1 To 7

        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(22, countR2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(24, countR2).Value = ""
         
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(22, countR2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(24, countR2).Interior.Color = vbWhite
        
         countR2 = countR2 + 1 'C
    Next
  
End Sub

Public Sub Planilha_Bancos_Rentabilidada()

    countR2 = 3

    For Each blc In colBalanco

        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(22, countR2).Value = blc.BCO_NIM_AJUST_CLI
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(24, countR2).Value = blc.BCO_EFICIENCY_RATIO_AJUS_CLI
         
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(22, countR2).Interior.Color = vbCyan
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Cells(24, countR2).Interior.Color = vbCyan
        
         countR2 = countR2 + 1 'C
    Next

    ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Columns("C:F").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Columns("C:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Columns("D:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Columns("E:F").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_RENTABILIDADE").Columns("F").EntireColumn.Hidden = True
    End If
  
End Sub

Public Sub Limpa_Planilha_Carteira()
    
    count2 = 3
    
    For i = 1 To 7
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(6, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(7, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(8, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(9, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(10, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(11, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(12, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(13, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(14, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(18, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(19, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(20, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(23, count2).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(26, count2).Value = ""
        
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(6, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(7, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(8, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(9, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(10, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(11, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(12, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(13, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(14, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(18, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(19, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(20, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(23, count2).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(26, count2).Interior.Color = vbWhite
        
        count2 = count2 + 2 'C
    Next i

End Sub

Public Sub Planilha_Bancos_Carteira()
    
    count2 = 3
    
    For Each blc In colBalanco

        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(6, count2).Value = blc.BCO_AA
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(7, count2).Value = blc.BCO_A
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(8, count2).Value = blc.BCO_B
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(9, count2).Value = blc.BCO_C
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(10, count2).Value = blc.BCO_D
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(11, count2).Value = blc.BCO_E
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(12, count2).Value = blc.BCO_F
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(13, count2).Value = blc.BCO_G
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(14, count2).Value = blc.BCO_H
        
        'ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(15, count2).Value = blc.BCO_TOTAL_CART
        'ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(16, count2).Value = blc.BCO_D_H
        'ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(17, count2).Value = blc.BCO_PDD_EXIG
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(18, count2).Value = blc.BCO_PDD_CONST
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(19, count2).Value = blc.BCO_VENCD
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(20, count2).Value = blc.BCO_VENCD_90D
        
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(23, count2).Value = blc.BCO_PDD_CARACT_CRED
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(26, count2).Value = blc.BCO_PDD_AVAIS_FIANCAS
        'ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(29, count2).Value = blc.BCO_PDD_CART_EXPAND
        
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(6, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(7, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(8, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(9, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(10, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(11, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(12, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(13, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(14, count2).Interior.Color = vbBlue
        
        'ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(15, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(18, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(19, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(20, count2).Interior.Color = vbBlue
        
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(23, count2).Interior.Color = vbBlue
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Cells(26, count2).Interior.Color = vbBlue
        
        count2 = count2 + 2 'C
    Next

    ActiveWorkbook.Worksheets("BANCOS_Carteira").Columns("C:J").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Columns("C:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Columns("E:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Columns("G:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_Carteira").Columns("I:J").EntireColumn.Hidden = True
    End If

End Sub

Public Sub Bancos_Mil()
    
    Dim iColOrig As Integer
    Dim iLinOrig As Integer
    Dim iColAux As Integer
    Dim iLinAux As Integer
    
    iLinAux = 2
    Do While iLinAux <= 8
        Planilha2.Cells(iLinAux, 4) = IIf(cd_grupo <> "", cd_grupo, "") 'cd_grp
        Planilha2.Cells(iLinAux, 5) = IIf(cd_cli <> "", cd_cli, "") 'cd_cli
        Planilha2.Cells(iLinAux, 7) = IIf(FLG_GRP <> "", FLG_GRP, "") 'FLG_GRP
        Planilha2.Cells(iLinAux, 10) = IIf(Sheet6.ComboBox1.Text <> "", Sheet6.ComboBox1.Text, "") 'Auditor
        Planilha2.Cells(iLinAux, 11) = IIf(Sheet6.Cells(3, 19) <> "", Sheet6.Cells(3, 19), "") 'Planilhador
        Planilha2.Cells(iLinAux, 12) = IIf(Layout <> "", Layout, "") 'Layout
        Planilha2.Cells(iLinAux, 13) = IIf(CNPJ <> "", CNPJ, "") 'CNPJ
        iLinAux = iLinAux + 1
    Loop

    ModuloBanco.Bancos_Mil_Grp1
    ModuloBanco.Bancos_Mil_Grp2
    ModuloBanco.Bancos_Mil_Grp3
    ModuloBanco.Bancos_Mil_Grp4
    ModuloBanco.Bancos_Mil_Grp5
    ModuloBanco.Bancos_Mil_Grp6
    ModuloBanco.Bancos_Mil_Grp7

End Sub


Public Sub Bancos_Carteira()

    Dim iColOrig As Integer
    Dim iLinOrig As Integer
    Dim iColAux As Integer
    Dim iLinAux As Integer
    
'BCO_AA Orig (6, 3) ==> Aux (127, 2)
'BCO_A Orig (7, 3) ==> Aux (134, 2)

    iColOrig = 3
    iLinOrig = 6
    
    iColAux = 127
    iLinAux = 2
    Do While iLinAux <= 8
        Planilha2.Cells(iLinAux, iColAux) = IIf(Sheet10.Cells(iLinOrig, iColOrig) <> "", Sheet10.Cells(iLinOrig, iColOrig), 0)
        Planilha2.Cells(iLinAux, iColAux).Interior.Color = vbBlue
        iLinAux = iLinAux + 1
        iColOrig = iColOrig + 2
    Loop
    
    iColOrig = 3
    iLinOrig = 7
    iColAux = 134
    iLinAux = 2
    Do While iLinOrig <= 29
        iLinAux = 2
        Do While iLinAux <= 8
            Planilha2.Cells(iLinAux, iColAux) = IIf(Sheet10.Cells(iLinOrig, iColOrig) <> "", Sheet10.Cells(iLinOrig, iColOrig), 0)
            Planilha2.Cells(iLinAux, iColAux).Interior.Color = vbBlue
            iLinAux = iLinAux + 1
            iColOrig = iColOrig + 2
        Loop
        iColOrig = 3
        iColAux = iColAux + 1
        iLinOrig = iLinOrig + 1
        If iLinOrig = 15 Then
            iColAux = 128
        End If
        If iLinOrig >= 21 Then
            iLinOrig = iLinOrig + 2
        End If
        If iLinOrig = 23 Then
            iColAux = 211
        End If
    Loop

End Sub


Public Sub Bancos_Contingencia()

    'Bco_Civeis_Conting_Provs
     Planilha2.Cells(2, 114) = IIf(Sheet13.Cells(6, 3) <> "", Sheet13.Cells(6, 3), 0)
     Planilha2.Cells(3, 114) = IIf(Sheet13.Cells(6, 5) <> "", Sheet13.Cells(6, 5), 0)
     Planilha2.Cells(4, 114) = IIf(Sheet13.Cells(6, 7) <> "", Sheet13.Cells(6, 7), 0)
     Planilha2.Cells(5, 114) = IIf(Sheet13.Cells(6, 9) <> "", Sheet13.Cells(6, 9), 0)
     Planilha2.Cells(6, 114) = IIf(Sheet13.Cells(6, 11) <> "", Sheet13.Cells(6, 11), 0)
     Planilha2.Cells(7, 114) = IIf(Sheet13.Cells(6, 13) <> "", Sheet13.Cells(6, 13), 0)
     Planilha2.Cells(8, 114) = IIf(Sheet13.Cells(6, 15) <> "", Sheet13.Cells(6, 15), 0)
    
    'Bco_Trablstas_Conting_Provs
     Planilha2.Cells(2, 107) = IIf(Sheet13.Cells(7, 3) <> "", Sheet13.Cells(7, 3), 0)
     Planilha2.Cells(3, 107) = IIf(Sheet13.Cells(7, 5) <> "", Sheet13.Cells(7, 5), 0)
     Planilha2.Cells(4, 107) = IIf(Sheet13.Cells(7, 7) <> "", Sheet13.Cells(7, 7), 0)
     Planilha2.Cells(5, 107) = IIf(Sheet13.Cells(7, 9) <> "", Sheet13.Cells(7, 9), 0)
     Planilha2.Cells(6, 107) = IIf(Sheet13.Cells(7, 11) <> "", Sheet13.Cells(7, 11), 0)
     Planilha2.Cells(7, 107) = IIf(Sheet13.Cells(7, 13) <> "", Sheet13.Cells(7, 13), 0)
     Planilha2.Cells(8, 107) = IIf(Sheet13.Cells(7, 15) <> "", Sheet13.Cells(7, 15), 0)
    
    'Bco_Fiscais_Conting_Provs
     Planilha2.Cells(2, 108) = IIf(Sheet13.Cells(8, 3) <> "", Sheet13.Cells(8, 3), 0)
     Planilha2.Cells(3, 108) = IIf(Sheet13.Cells(8, 5) <> "", Sheet13.Cells(8, 5), 0)
     Planilha2.Cells(4, 108) = IIf(Sheet13.Cells(8, 7) <> "", Sheet13.Cells(8, 7), 0)
     Planilha2.Cells(5, 108) = IIf(Sheet13.Cells(8, 9) <> "", Sheet13.Cells(8, 9), 0)
     Planilha2.Cells(6, 108) = IIf(Sheet13.Cells(8, 11) <> "", Sheet13.Cells(8, 11), 0)
     Planilha2.Cells(7, 108) = IIf(Sheet13.Cells(8, 13) <> "", Sheet13.Cells(8, 13), 0)
     Planilha2.Cells(8, 108) = IIf(Sheet13.Cells(8, 15) <> "", Sheet13.Cells(8, 15), 0)

    'Bco_Total_Conting_Provs
     Planilha2.Cells(2, 109) = IIf(Sheet13.Cells(5, 3) <> "", Sheet13.Cells(5, 3), 0)
     Planilha2.Cells(3, 109) = IIf(Sheet13.Cells(5, 5) <> "", Sheet13.Cells(5, 5), 0)
     Planilha2.Cells(4, 109) = IIf(Sheet13.Cells(5, 7) <> "", Sheet13.Cells(5, 7), 0)
     Planilha2.Cells(5, 109) = IIf(Sheet13.Cells(5, 9) <> "", Sheet13.Cells(5, 9), 0)
     Planilha2.Cells(6, 109) = IIf(Sheet13.Cells(5, 11) <> "", Sheet13.Cells(5, 11), 0)
     Planilha2.Cells(7, 109) = IIf(Sheet13.Cells(5, 13) <> "", Sheet13.Cells(5, 13), 0)
     Planilha2.Cells(8, 109) = IIf(Sheet13.Cells(5, 15) <> "", Sheet13.Cells(5, 15), 0)

    'blc.Bco_Civeis_Depos_Judc
     Planilha2.Cells(2, 110) = IIf(Sheet13.Cells(10, 3) <> "", Sheet13.Cells(10, 3), 0)
     Planilha2.Cells(3, 110) = IIf(Sheet13.Cells(10, 5) <> "", Sheet13.Cells(10, 5), 0)
     Planilha2.Cells(4, 110) = IIf(Sheet13.Cells(10, 7) <> "", Sheet13.Cells(10, 7), 0)
     Planilha2.Cells(5, 110) = IIf(Sheet13.Cells(10, 9) <> "", Sheet13.Cells(10, 9), 0)
     Planilha2.Cells(6, 110) = IIf(Sheet13.Cells(10, 11) <> "", Sheet13.Cells(10, 11), 0)
     Planilha2.Cells(7, 110) = IIf(Sheet13.Cells(10, 13) <> "", Sheet13.Cells(10, 13), 0)
     Planilha2.Cells(8, 110) = IIf(Sheet13.Cells(10, 15) <> "", Sheet13.Cells(10, 15), 0)
    
    'blc.Bco_Trablstas_Depos_Judc
     Planilha2.Cells(2, 111) = IIf(Sheet13.Cells(11, 3) <> "", Sheet13.Cells(11, 3), 0)
     Planilha2.Cells(3, 111) = IIf(Sheet13.Cells(11, 5) <> "", Sheet13.Cells(11, 5), 0)
     Planilha2.Cells(4, 111) = IIf(Sheet13.Cells(11, 7) <> "", Sheet13.Cells(11, 7), 0)
     Planilha2.Cells(5, 111) = IIf(Sheet13.Cells(11, 9) <> "", Sheet13.Cells(11, 9), 0)
     Planilha2.Cells(6, 111) = IIf(Sheet13.Cells(11, 11) <> "", Sheet13.Cells(11, 11), 0)
     Planilha2.Cells(7, 111) = IIf(Sheet13.Cells(11, 13) <> "", Sheet13.Cells(11, 13), 0)
     Planilha2.Cells(8, 111) = IIf(Sheet13.Cells(11, 15) <> "", Sheet13.Cells(11, 15), 0)
           
    'blc.Bco_Fiscais_Depos_Judc
     Planilha2.Cells(2, 112) = IIf(Sheet13.Cells(12, 3) <> "", Sheet13.Cells(12, 3), 0)
     Planilha2.Cells(3, 112) = IIf(Sheet13.Cells(12, 5) <> "", Sheet13.Cells(12, 5), 0)
     Planilha2.Cells(4, 112) = IIf(Sheet13.Cells(12, 7) <> "", Sheet13.Cells(12, 7), 0)
     Planilha2.Cells(5, 112) = IIf(Sheet13.Cells(12, 9) <> "", Sheet13.Cells(12, 9), 0)
     Planilha2.Cells(6, 112) = IIf(Sheet13.Cells(12, 11) <> "", Sheet13.Cells(12, 11), 0)
     Planilha2.Cells(7, 112) = IIf(Sheet13.Cells(12, 13) <> "", Sheet13.Cells(12, 13), 0)
     Planilha2.Cells(8, 112) = IIf(Sheet13.Cells(12, 15) <> "", Sheet13.Cells(12, 15), 0)
    
    'blc.Bco_Total_Depos_Judc
     Planilha2.Cells(2, 113) = IIf(Sheet13.Cells(9, 3) <> "", Sheet13.Cells(9, 3), 0)
     Planilha2.Cells(3, 113) = IIf(Sheet13.Cells(9, 5) <> "", Sheet13.Cells(9, 5), 0)
     Planilha2.Cells(4, 113) = IIf(Sheet13.Cells(9, 7) <> "", Sheet13.Cells(9, 7), 0)
     Planilha2.Cells(5, 113) = IIf(Sheet13.Cells(9, 9) <> "", Sheet13.Cells(9, 9), 0)
     Planilha2.Cells(6, 113) = IIf(Sheet13.Cells(9, 11) <> "", Sheet13.Cells(9, 11), 0)
     Planilha2.Cells(7, 113) = IIf(Sheet13.Cells(9, 13) <> "", Sheet13.Cells(9, 13), 0)
     Planilha2.Cells(8, 113) = IIf(Sheet13.Cells(9, 15) <> "", Sheet13.Cells(9, 15), 0)

    'blc.Bco_Civeis_Conting_Nao_Provs
     Planilha2.Cells(2, 103) = IIf(Sheet13.Cells(14, 3) <> "", Sheet13.Cells(14, 3), 0)
     Planilha2.Cells(3, 103) = IIf(Sheet13.Cells(14, 5) <> "", Sheet13.Cells(14, 5), 0)
     Planilha2.Cells(4, 103) = IIf(Sheet13.Cells(14, 7) <> "", Sheet13.Cells(14, 7), 0)
     Planilha2.Cells(5, 103) = IIf(Sheet13.Cells(14, 9) <> "", Sheet13.Cells(14, 9), 0)
     Planilha2.Cells(6, 103) = IIf(Sheet13.Cells(14, 11) <> "", Sheet13.Cells(14, 11), 0)
     Planilha2.Cells(7, 103) = IIf(Sheet13.Cells(14, 13) <> "", Sheet13.Cells(14, 13), 0)
     Planilha2.Cells(8, 103) = IIf(Sheet13.Cells(14, 15) <> "", Sheet13.Cells(14, 15), 0)

    'blc.Bco_Trablstas_Conting_Nao_Provs
     Planilha2.Cells(2, 104) = IIf(Sheet13.Cells(15, 3) <> "", Sheet13.Cells(15, 3), 0)
     Planilha2.Cells(3, 104) = IIf(Sheet13.Cells(15, 5) <> "", Sheet13.Cells(15, 5), 0)
     Planilha2.Cells(4, 104) = IIf(Sheet13.Cells(15, 7) <> "", Sheet13.Cells(15, 7), 0)
     Planilha2.Cells(5, 104) = IIf(Sheet13.Cells(15, 9) <> "", Sheet13.Cells(15, 9), 0)
     Planilha2.Cells(6, 104) = IIf(Sheet13.Cells(15, 11) <> "", Sheet13.Cells(15, 11), 0)
     Planilha2.Cells(7, 104) = IIf(Sheet13.Cells(15, 13) <> "", Sheet13.Cells(15, 13), 0)
     Planilha2.Cells(8, 104) = IIf(Sheet13.Cells(15, 15) <> "", Sheet13.Cells(15, 15), 0)

    'blc.Bco_Fiscais_Conting_Nao_Provs
     Planilha2.Cells(2, 105) = IIf(Sheet13.Cells(16, 3) <> "", Sheet13.Cells(16, 3), 0)
     Planilha2.Cells(3, 105) = IIf(Sheet13.Cells(16, 5) <> "", Sheet13.Cells(16, 5), 0)
     Planilha2.Cells(4, 105) = IIf(Sheet13.Cells(16, 7) <> "", Sheet13.Cells(16, 7), 0)
     Planilha2.Cells(5, 105) = IIf(Sheet13.Cells(16, 9) <> "", Sheet13.Cells(16, 9), 0)
     Planilha2.Cells(6, 105) = IIf(Sheet13.Cells(16, 11) <> "", Sheet13.Cells(16, 11), 0)
     Planilha2.Cells(7, 105) = IIf(Sheet13.Cells(16, 13) <> "", Sheet13.Cells(16, 13), 0)
     Planilha2.Cells(8, 105) = IIf(Sheet13.Cells(16, 15) <> "", Sheet13.Cells(16, 15), 0)
    
    'blc.Bco_Total_Conting_Nao_Provs
     Planilha2.Cells(2, 106) = IIf(Sheet13.Cells(13, 3) <> "", Sheet13.Cells(13, 3), 0)
     Planilha2.Cells(3, 106) = IIf(Sheet13.Cells(13, 5) <> "", Sheet13.Cells(13, 5), 0)
     Planilha2.Cells(4, 106) = IIf(Sheet13.Cells(13, 7) <> "", Sheet13.Cells(13, 7), 0)
     Planilha2.Cells(5, 106) = IIf(Sheet13.Cells(13, 9) <> "", Sheet13.Cells(13, 9), 0)
     Planilha2.Cells(6, 106) = IIf(Sheet13.Cells(13, 11) <> "", Sheet13.Cells(13, 11), 0)
     Planilha2.Cells(7, 106) = IIf(Sheet13.Cells(13, 13) <> "", Sheet13.Cells(13, 13), 0)
     Planilha2.Cells(8, 106) = IIf(Sheet13.Cells(13, 15) <> "", Sheet13.Cells(13, 15), 0)

    Planilha2.Cells(2, 114).Interior.Color = vbYellow
     Planilha2.Cells(3, 114).Interior.Color = vbYellow
     Planilha2.Cells(4, 114).Interior.Color = vbYellow
     Planilha2.Cells(5, 114).Interior.Color = vbYellow
     Planilha2.Cells(6, 114).Interior.Color = vbYellow
     Planilha2.Cells(7, 114).Interior.Color = vbYellow
     Planilha2.Cells(8, 114).Interior.Color = vbYellow
     Planilha2.Cells(2, 107).Interior.Color = vbYellow
     Planilha2.Cells(3, 107).Interior.Color = vbYellow
     Planilha2.Cells(4, 107).Interior.Color = vbYellow
     Planilha2.Cells(5, 107).Interior.Color = vbYellow
     Planilha2.Cells(6, 107).Interior.Color = vbYellow
     Planilha2.Cells(7, 107).Interior.Color = vbYellow
     Planilha2.Cells(8, 107).Interior.Color = vbYellow
    
    'Bco_Fiscais_Conting_Provs
     Planilha2.Cells(2, 108).Interior.Color = vbYellow
     Planilha2.Cells(3, 108).Interior.Color = vbYellow
     Planilha2.Cells(4, 108).Interior.Color = vbYellow
     Planilha2.Cells(5, 108).Interior.Color = vbYellow
     Planilha2.Cells(6, 108).Interior.Color = vbYellow
     Planilha2.Cells(7, 108).Interior.Color = vbYellow
     Planilha2.Cells(8, 108).Interior.Color = vbYellow

    'Bco_Total_Conting_Provs
     Planilha2.Cells(2, 109).Interior.Color = vbYellow
     Planilha2.Cells(3, 109).Interior.Color = vbYellow
     Planilha2.Cells(4, 109).Interior.Color = vbYellow
     Planilha2.Cells(5, 109).Interior.Color = vbYellow
     Planilha2.Cells(6, 109).Interior.Color = vbYellow
     Planilha2.Cells(7, 109).Interior.Color = vbYellow
     Planilha2.Cells(8, 109).Interior.Color = vbYellow

    'blc.Bco_Civeis_Depos_Judc
     Planilha2.Cells(2, 110).Interior.Color = vbYellow
     Planilha2.Cells(3, 110).Interior.Color = vbYellow
     Planilha2.Cells(4, 110).Interior.Color = vbYellow
     Planilha2.Cells(5, 110).Interior.Color = vbYellow
     Planilha2.Cells(6, 110).Interior.Color = vbYellow
     Planilha2.Cells(7, 110).Interior.Color = vbYellow
     Planilha2.Cells(8, 110).Interior.Color = vbYellow
    
    'blc.Bco_Trablstas_Depos_Judc
     Planilha2.Cells(2, 111).Interior.Color = vbYellow
     Planilha2.Cells(3, 111).Interior.Color = vbYellow
     Planilha2.Cells(4, 111).Interior.Color = vbYellow
     Planilha2.Cells(5, 111).Interior.Color = vbYellow
     Planilha2.Cells(6, 111).Interior.Color = vbYellow
     Planilha2.Cells(7, 111).Interior.Color = vbYellow
     Planilha2.Cells(8, 111).Interior.Color = vbYellow
           
    'blc.Bco_Fiscais_Depos_Judc
     Planilha2.Cells(2, 112).Interior.Color = vbYellow
     Planilha2.Cells(3, 112).Interior.Color = vbYellow
     Planilha2.Cells(4, 112).Interior.Color = vbYellow
     Planilha2.Cells(5, 112).Interior.Color = vbYellow
     Planilha2.Cells(6, 112).Interior.Color = vbYellow
     Planilha2.Cells(7, 112).Interior.Color = vbYellow
     Planilha2.Cells(8, 112).Interior.Color = vbYellow
    
    'blc.Bco_Total_Depos_Judc
     Planilha2.Cells(2, 113).Interior.Color = vbYellow
     Planilha2.Cells(3, 113).Interior.Color = vbYellow
     Planilha2.Cells(4, 113).Interior.Color = vbYellow
     Planilha2.Cells(5, 113).Interior.Color = vbYellow
     Planilha2.Cells(6, 113).Interior.Color = vbYellow
     Planilha2.Cells(7, 113).Interior.Color = vbYellow
     Planilha2.Cells(8, 113).Interior.Color = vbYellow

    'blc.Bco_Civeis_Conting_Nao_Provs
     Planilha2.Cells(2, 103).Interior.Color = vbYellow
     Planilha2.Cells(3, 103).Interior.Color = vbYellow
     Planilha2.Cells(4, 103).Interior.Color = vbYellow
     Planilha2.Cells(5, 103).Interior.Color = vbYellow
     Planilha2.Cells(6, 103).Interior.Color = vbYellow
     Planilha2.Cells(7, 103).Interior.Color = vbYellow
     Planilha2.Cells(8, 103).Interior.Color = vbYellow

    'blc.Bco_Trablstas_Conting_Nao_Provs
     Planilha2.Cells(2, 104).Interior.Color = vbYellow
     Planilha2.Cells(3, 104).Interior.Color = vbYellow
     Planilha2.Cells(4, 104).Interior.Color = vbYellow
     Planilha2.Cells(5, 104).Interior.Color = vbYellow
     Planilha2.Cells(6, 104).Interior.Color = vbYellow
     Planilha2.Cells(7, 104).Interior.Color = vbYellow
     Planilha2.Cells(8, 104).Interior.Color = vbYellow

    'blc.Bco_Fiscais_Conting_Nao_Provs
     Planilha2.Cells(2, 105).Interior.Color = vbYellow
     Planilha2.Cells(3, 105).Interior.Color = vbYellow
     Planilha2.Cells(4, 105).Interior.Color = vbYellow
     Planilha2.Cells(5, 105).Interior.Color = vbYellow
     Planilha2.Cells(6, 105).Interior.Color = vbYellow
     Planilha2.Cells(7, 105).Interior.Color = vbYellow
     Planilha2.Cells(8, 105).Interior.Color = vbYellow
    
    'blc.Bco_Total_Conting_Nao_Provs
     Planilha2.Cells(2, 106).Interior.Color = vbYellow
     Planilha2.Cells(3, 106).Interior.Color = vbYellow
     Planilha2.Cells(4, 106).Interior.Color = vbYellow
     Planilha2.Cells(5, 106).Interior.Color = vbYellow
     Planilha2.Cells(6, 106).Interior.Color = vbYellow
     Planilha2.Cells(7, 106).Interior.Color = vbYellow
     Planilha2.Cells(8, 106).Interior.Color = vbYellow

End Sub

Public Sub Bancos_Rentabilidade()

    ' BCO_NIM_AJUST_CLI
    Planilha2.Cells(2, 239) = IIf(Sheet14.Cells(22, 3) <> "", Sheet14.Cells(22, 3), 0)
    Planilha2.Cells(3, 239) = IIf(Sheet14.Cells(22, 4) <> "", Sheet14.Cells(22, 4), 0)
    Planilha2.Cells(4, 239) = IIf(Sheet14.Cells(22, 5) <> "", Sheet14.Cells(22, 5), 0)
    Planilha2.Cells(5, 239) = IIf(Sheet14.Cells(22, 6) <> "", Sheet6.Cells(22, 6), 0)
    Planilha2.Cells(6, 239) = IIf(Sheet14.Cells(22, 7) <> "", Sheet14.Cells(22, 7), 0)
    Planilha2.Cells(7, 239) = IIf(Sheet14.Cells(22, 8) <> "", Sheet14.Cells(22, 8), 0)
    Planilha2.Cells(8, 239) = IIf(Sheet14.Cells(22, 9) <> "", Sheet14.Cells(22, 9), 0)
    
    ' BCO_EFICIENCY_RATIO_AJUS_CLI
    Planilha2.Cells(2, 238) = IIf(Sheet14.Cells(24, 3) <> "", Sheet14.Cells(24, 3), 0)
    Planilha2.Cells(3, 238) = IIf(Sheet14.Cells(24, 4) <> "", Sheet14.Cells(24, 4), 0)
    Planilha2.Cells(4, 238) = IIf(Sheet14.Cells(24, 5) <> "", Sheet14.Cells(24, 5), 0)
    Planilha2.Cells(5, 238) = IIf(Sheet14.Cells(24, 6) <> "", Sheet6.Cells(24, 6), 0)
    Planilha2.Cells(6, 238) = IIf(Sheet14.Cells(24, 7) <> "", Sheet14.Cells(24, 7), 0)
    Planilha2.Cells(7, 238) = IIf(Sheet14.Cells(24, 8) <> "", Sheet14.Cells(24, 8), 0)
    Planilha2.Cells(8, 238) = IIf(Sheet14.Cells(24, 9) <> "", Sheet14.Cells(24, 9), 0)
    
    Planilha2.Cells(2, 239).Interior.Color = vbCyan
    Planilha2.Cells(3, 239).Interior.Color = vbCyan
    Planilha2.Cells(4, 239).Interior.Color = vbCyan
    Planilha2.Cells(5, 239).Interior.Color = vbCyan
    Planilha2.Cells(6, 239).Interior.Color = vbCyan
    Planilha2.Cells(7, 239).Interior.Color = vbCyan
    Planilha2.Cells(8, 239).Interior.Color = vbCyan
    
    ' BCO_EFICIENCY_RATIO_AJUS_CLI
    Planilha2.Cells(2, 238).Interior.Color = vbCyan
    Planilha2.Cells(3, 238).Interior.Color = vbCyan
    Planilha2.Cells(4, 238).Interior.Color = vbCyan
    Planilha2.Cells(5, 238).Interior.Color = vbCyan
    Planilha2.Cells(6, 238).Interior.Color = vbCyan
    Planilha2.Cells(7, 238).Interior.Color = vbCyan
    Planilha2.Cells(8, 238).Interior.Color = vbCyan

End Sub


Public Sub Limpa_Planilha_TVM()

    count = 3

    For i = 1 To 7
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(7, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(8, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(9, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(10, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(12, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(13, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(14, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(15, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(17, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(18, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(19, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(20, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(22, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(23, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(24, count).Value = ""
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(25, count).Value = ""

        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(22, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(23, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(24, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(25, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(17, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(18, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(19, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(20, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(12, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(13, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(14, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(15, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(7, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(8, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(9, count).Interior.Color = vbWhite
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(10, count).Interior.Color = vbWhite
        
        count = count + 2
    Next

End Sub

Public Sub Planilha_Bancos_TVM()

    count = 3

    For Each blc In colBalanco
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(7, count).Value = blc.Bco_P_negoc_Valor_Custo
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(8, count).Value = blc.Bco_Disp_Venda_Valor_Custo
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(9, count).Value = blc.Bco_Mtdos_Vcto_Valor_Custo
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(10, count).Value = blc.Bco_Instr_Financ_Deriv_Vlr_custo

        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(12, count).Value = blc.Bco_P_negoc_Valor_Contab
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(13, count).Value = blc.Bco_Disp_Venda_Valor_Contab
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(14, count).Value = blc.Bco_Mtdos_Vcto_Valor_Contab
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(15, count).Value = blc.Bco_Instr_Fin_Deriv_Vlr_Contab

        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(17, count).Value = blc.Bco_P_negoc_Mtm
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(18, count).Value = blc.Bco_Disp_Venda_Mtm
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(19, count).Value = blc.Bco_Mtdos_Vcto_Mtm
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(20, count).Value = blc.Bco_Instr_Financ_Deriv_Mtm

        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(22, count).Value = blc.Bco_P_Negoc_Prov_P_Desv
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(23, count).Value = blc.Bco_Disp_Venda_Prov_P_Desv
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(24, count).Value = blc.Bco_Mtdos_Vcto_Prov_P_Desv
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(25, count).Value = blc.Bco_Instr_Fin_Deriv_Prov_P_Desv

        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(22, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(23, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(24, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(25, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(17, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(18, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(19, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(20, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(12, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(13, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(14, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(15, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(7, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(8, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(9, count).Interior.Color = vbRed
        ActiveWorkbook.Worksheets("BANCOS_TVM").Cells(10, count).Interior.Color = vbRed
        
        count = count + 2
    Next
    
    ActiveWorkbook.Worksheets("BANCOS_TVM").Columns("C:J").EntireColumn.Hidden = False
    If colBalanco.count = 0 Then
        ActiveWorkbook.Worksheets("BANCOS_TVM").Columns("C:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 1 Then
        ActiveWorkbook.Worksheets("BANCOS_TVM").Columns("E:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 2 Then
        ActiveWorkbook.Worksheets("BANCOS_TVM").Columns("G:J").EntireColumn.Hidden = True
    ElseIf colBalanco.count = 3 Then
        ActiveWorkbook.Worksheets("BANCOS_TVM").Columns("I:J").EntireColumn.Hidden = True
    End If

End Sub

Public Sub Bancos_PDD()

    Dim count1 As Integer 'linha planilha aux
    Dim count2 As Integer 'linha planilha origem
    Dim count3 As Integer 'coluna planilha origem

    count1 = 2
    count2 = 5
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 161) = IIf(Sheet11.Cells(count2, count3) <> "", Sheet11.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 161).Interior.Color = vbGreen
        count3 = count3 + 1
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For kCrt = 0 To 6
        Planilha2.Cells(count1, 162) = IIf(Sheet11.Cells(count2, count3) <> "", Sheet11.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 162).Interior.Color = vbGreen
        count3 = count3 + 1
        count1 = count1 + 1
    Next kCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For kCrt = 0 To 6
        Planilha2.Cells(count1, 163) = IIf(Sheet11.Cells(count2, count3) <> "", Sheet11.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 163).Interior.Color = vbGreen
        count3 = count3 + 1
        count1 = count1 + 1
    Next kCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 4
    For mCrt = 0 To 6
        Planilha2.Cells(count1, 166) = IIf(Sheet11.Cells(count2, count3) <> "", Sheet11.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 166).Interior.Color = vbGreen
        count3 = count3 + 1
        count1 = count1 + 1
    Next mCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For nCrt = 0 To 6
        Planilha2.Cells(count1, 216) = IIf(Sheet11.Cells(count2, count3) <> "", Sheet11.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 216).Interior.Color = vbGreen
        count3 = count3 + 1
        count1 = count1 + 1
    Next nCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For oCrt = 0 To 6
        Planilha2.Cells(count1, 167) = IIf(Sheet11.Cells(count2, count3) <> "", Sheet11.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 167).Interior.Color = vbGreen
        count3 = count3 + 1
        count1 = count1 + 1
    Next oCrt

End Sub


Public Sub Bancos_Founding()

    Dim count1 As Integer 'linha planilha aux
    Dim count2 As Integer 'linha planilha origem
    Dim count3 As Integer 'coluna planilha origem

    count1 = 2
    count2 = 11
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 168) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 168).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 169) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 169).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
     For jCrt = 0 To 6
        Planilha2.Cells(count1, 170) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 170).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 208) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 208).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 207) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 207).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 5
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 234) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 234).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 200) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 200).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 4
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 173) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 173).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 10
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 146) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 146).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 147) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 147).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 209) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 209).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 20
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 142) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 142).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
    count1 = 2
    count2 = count2 + 1
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 143) = IIf(Sheet8.Cells(count2, count3) <> "", Sheet8.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 143).Interior.Color = vbMagenta
        count1 = count1 + 1
        count3 = count3 + 1
    Next jCrt
    
End Sub


Public Sub Bancos_TVM()
    Dim count1 As Integer 'linha planilha aux
    Dim count2 As Integer 'linha planilha origem
    Dim count3 As Integer 'coluna planilha origem

    count1 = 2
    count2 = 7
    count3 = 3
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 115) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 115).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 118) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 118).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 121) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 121).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 124) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 124).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 2
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 116) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 116).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 119) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 119).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 122) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 122).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 125) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 125).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 2
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 117) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 117).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 120) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 120).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 123) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 123).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 126) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 126).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 2
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 220) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 220).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 217) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 217).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 219) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 219).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt
    
    count3 = 3
    count1 = 2
    count2 = count2 + 1
    For jCrt = 0 To 6
        Planilha2.Cells(count1, 218) = IIf(Sheet12.Cells(count2, count3) <> "", Sheet12.Cells(count2, count3), 0)
        Planilha2.Cells(count1, 218).Interior.Color = vbRed
        count3 = count3 + 2
        count1 = count1 + 1
    Next jCrt

End Sub

Public Sub Bancos_Mil_Grp1()

    Planilha2.Cells(2, 1) = 0
    Planilha2.Cells(3, 1) = 0
    Planilha2.Cells(4, 1) = 0
    Planilha2.Cells(5, 1) = 0
    Planilha2.Cells(6, 1) = 0
    Planilha2.Cells(7, 1) = 0
    Planilha2.Cells(8, 1) = 0
    
    'Mes fechamento
    Planilha2.Cells(2, 6) = Sheet6.Cells(5, 3)
    Planilha2.Cells(3, 6) = Sheet6.Cells(5, 5)
    Planilha2.Cells(4, 6) = Sheet6.Cells(5, 7)
    Planilha2.Cells(5, 6) = Sheet6.Cells(5, 9)
    Planilha2.Cells(6, 6) = Sheet6.Cells(5, 11)
    Planilha2.Cells(7, 6) = Sheet6.Cells(5, 13)
    Planilha2.Cells(8, 6) = Sheet6.Cells(5, 15)
    
    'DT_EXERC
    Planilha2.Cells(2, 2) = Sheet6.Cells(6, 3)
    Planilha2.Cells(3, 2) = Sheet6.Cells(6, 5)
    Planilha2.Cells(4, 2) = Sheet6.Cells(6, 7)
    Planilha2.Cells(5, 2) = Sheet6.Cells(6, 9)
    Planilha2.Cells(6, 2) = Sheet6.Cells(6, 11)
    Planilha2.Cells(7, 2) = Sheet6.Cells(6, 13)
    Planilha2.Cells(8, 2) = Sheet6.Cells(6, 15)
        
    'blc.Bco_Ativo_Disp
    Planilha2.Cells(2, 14) = IIf(Sheet6.Cells(7, 3) <> "", Sheet6.Cells(7, 3), 0)
    Planilha2.Cells(3, 14) = IIf(Sheet6.Cells(7, 5) <> "", Sheet6.Cells(7, 5), 0)
    Planilha2.Cells(4, 14) = IIf(Sheet6.Cells(7, 7) <> "", Sheet6.Cells(7, 7), 0)
    Planilha2.Cells(5, 14) = IIf(Sheet6.Cells(7, 9) <> "", Sheet6.Cells(7, 9), 0)
    Planilha2.Cells(6, 14) = IIf(Sheet6.Cells(7, 11) <> "", Sheet6.Cells(7, 11), 0)
    Planilha2.Cells(7, 14) = IIf(Sheet6.Cells(7, 13) <> "", Sheet6.Cells(7, 13), 0)
    Planilha2.Cells(8, 14) = IIf(Sheet6.Cells(7, 15) <> "", Sheet6.Cells(7, 15), 0)
    
    'CDI
    Planilha2.Cells(2, 15) = IIf(Sheet6.Cells(8, 3) <> "", Sheet6.Cells(8, 3), 0)
    Planilha2.Cells(3, 15) = IIf(Sheet6.Cells(8, 5) <> "", Sheet6.Cells(8, 5), 0)
    Planilha2.Cells(4, 15) = IIf(Sheet6.Cells(8, 7) <> "", Sheet6.Cells(8, 7), 0)
    Planilha2.Cells(5, 15) = IIf(Sheet6.Cells(8, 9) <> "", Sheet6.Cells(8, 9), 0)
    Planilha2.Cells(6, 15) = IIf(Sheet6.Cells(8, 11) <> "", Sheet6.Cells(8, 11), 0)
    Planilha2.Cells(7, 15) = IIf(Sheet6.Cells(8, 13) <> "", Sheet6.Cells(8, 13), 0)
    Planilha2.Cells(8, 15) = IIf(Sheet6.Cells(8, 15) <> "", Sheet6.Cells(8, 15), 0)
    
    'BCO_ATIVO_TITULO_MERC_ABERT
    Planilha2.Cells(2, 16) = IIf(Sheet6.Cells(9, 3) <> "", Sheet6.Cells(9, 3), 0)
    Planilha2.Cells(3, 16) = IIf(Sheet6.Cells(9, 5) <> "", Sheet6.Cells(9, 5), 0)
    Planilha2.Cells(4, 16) = IIf(Sheet6.Cells(9, 7) <> "", Sheet6.Cells(9, 7), 0)
    Planilha2.Cells(5, 16) = IIf(Sheet6.Cells(9, 9) <> "", Sheet6.Cells(9, 9), 0)
    Planilha2.Cells(6, 16) = IIf(Sheet6.Cells(9, 11) <> "", Sheet6.Cells(9, 11), 0)
    Planilha2.Cells(7, 16) = IIf(Sheet6.Cells(9, 13) <> "", Sheet6.Cells(9, 13), 0)
    Planilha2.Cells(8, 16) = IIf(Sheet6.Cells(9, 15) <> "", Sheet6.Cells(9, 15), 0)
    
    'BCO_ATIVO_TVM
    Planilha2.Cells(2, 17) = IIf(Sheet6.Cells(10, 3) <> "", Sheet6.Cells(10, 3), 0)
    Planilha2.Cells(3, 17) = IIf(Sheet6.Cells(10, 5) <> "", Sheet6.Cells(10, 5), 0)
    Planilha2.Cells(4, 17) = IIf(Sheet6.Cells(10, 7) <> "", Sheet6.Cells(10, 7), 0)
    Planilha2.Cells(5, 17) = IIf(Sheet6.Cells(10, 9) <> "", Sheet6.Cells(10, 9), 0)
    Planilha2.Cells(6, 17) = IIf(Sheet6.Cells(10, 11) <> "", Sheet6.Cells(10, 11), 0)
    Planilha2.Cells(7, 17) = IIf(Sheet6.Cells(10, 13) <> "", Sheet6.Cells(10, 13), 0)
    Planilha2.Cells(8, 17) = IIf(Sheet6.Cells(10, 15) <> "", Sheet6.Cells(10, 15), 0)
    
    'BCO_ATIVO_OPERAC_CRED
    Planilha2.Cells(2, 18) = IIf(Sheet6.Cells(11, 3) <> "", Sheet6.Cells(11, 3), 0)
    Planilha2.Cells(3, 18) = IIf(Sheet6.Cells(11, 5) <> "", Sheet6.Cells(11, 5), 0)
    Planilha2.Cells(4, 18) = IIf(Sheet6.Cells(11, 7) <> "", Sheet6.Cells(11, 7), 0)
    Planilha2.Cells(5, 18) = IIf(Sheet6.Cells(11, 9) <> "", Sheet6.Cells(11, 9), 0)
    Planilha2.Cells(6, 18) = IIf(Sheet6.Cells(11, 11) <> "", Sheet6.Cells(11, 11), 0)
    Planilha2.Cells(7, 18) = IIf(Sheet6.Cells(11, 13) <> "", Sheet6.Cells(11, 13), 0)
    Planilha2.Cells(8, 18) = IIf(Sheet6.Cells(11, 15) <> "", Sheet6.Cells(11, 15), 0)
      
    'BCO_ATIVO_PDD
    Planilha2.Cells(2, 19) = IIf(Sheet6.Cells(12, 3) <> "", Sheet6.Cells(12, 3), 0)
    Planilha2.Cells(3, 19) = IIf(Sheet6.Cells(12, 5) <> "", Sheet6.Cells(12, 5), 0)
    Planilha2.Cells(4, 19) = IIf(Sheet6.Cells(12, 7) <> "", Sheet6.Cells(12, 7), 0)
    Planilha2.Cells(5, 19) = IIf(Sheet6.Cells(12, 9) <> "", Sheet6.Cells(12, 9), 0)
    Planilha2.Cells(6, 19) = IIf(Sheet6.Cells(12, 11) <> "", Sheet6.Cells(12, 11), 0)
    Planilha2.Cells(7, 19) = IIf(Sheet6.Cells(12, 13) <> "", Sheet6.Cells(12, 13), 0)
    Planilha2.Cells(8, 19) = IIf(Sheet6.Cells(12, 15) <> "", Sheet6.Cells(12, 15), 0)
      
    'BCO_ATIVO_OP_ARREND_MERCATL
    Planilha2.Cells(2, 20) = IIf(Sheet6.Cells(13, 3) <> "", Sheet6.Cells(13, 3), 0)
    Planilha2.Cells(3, 20) = IIf(Sheet6.Cells(13, 5) <> "", Sheet6.Cells(13, 5), 0)
    Planilha2.Cells(4, 20) = IIf(Sheet6.Cells(13, 7) <> "", Sheet6.Cells(13, 7), 0)
    Planilha2.Cells(5, 20) = IIf(Sheet6.Cells(13, 9) <> "", Sheet6.Cells(13, 9), 0)
    Planilha2.Cells(6, 20) = IIf(Sheet6.Cells(13, 11) <> "", Sheet6.Cells(13, 11), 0)
    Planilha2.Cells(7, 20) = IIf(Sheet6.Cells(13, 13) <> "", Sheet6.Cells(13, 13), 0)
    Planilha2.Cells(8, 20) = IIf(Sheet6.Cells(13, 15) <> "", Sheet6.Cells(13, 15), 0)
    
    'BCO_ATIVO_PDD_ARREND_MERCATL
    Planilha2.Cells(2, 21) = IIf(Sheet6.Cells(14, 3) <> "", Sheet6.Cells(14, 3), 0)
    Planilha2.Cells(3, 21) = IIf(Sheet6.Cells(14, 5) <> "", Sheet6.Cells(14, 5), 0)
    Planilha2.Cells(4, 21) = IIf(Sheet6.Cells(14, 7) <> "", Sheet6.Cells(14, 7), 0)
    Planilha2.Cells(5, 21) = IIf(Sheet6.Cells(14, 9) <> "", Sheet6.Cells(14, 9), 0)
    Planilha2.Cells(6, 21) = IIf(Sheet6.Cells(14, 11) <> "", Sheet6.Cells(14, 11), 0)
    Planilha2.Cells(7, 21) = IIf(Sheet6.Cells(14, 13) <> "", Sheet6.Cells(14, 13), 0)
    Planilha2.Cells(8, 21) = IIf(Sheet6.Cells(14, 15) <> "", Sheet6.Cells(14, 15), 0)
      
    'BCO_ATIVO_DESP_ANTEC
    Planilha2.Cells(2, 22) = IIf(Sheet6.Cells(15, 3) <> "", Sheet6.Cells(15, 3), 0)
    Planilha2.Cells(3, 22) = IIf(Sheet6.Cells(15, 5) <> "", Sheet6.Cells(15, 5), 0)
    Planilha2.Cells(4, 22) = IIf(Sheet6.Cells(15, 7) <> "", Sheet6.Cells(15, 7), 0)
    Planilha2.Cells(5, 22) = IIf(Sheet6.Cells(15, 9) <> "", Sheet6.Cells(15, 9), 0)
    Planilha2.Cells(6, 22) = IIf(Sheet6.Cells(15, 11) <> "", Sheet6.Cells(15, 11), 0)
    Planilha2.Cells(7, 22) = IIf(Sheet6.Cells(15, 13) <> "", Sheet6.Cells(15, 13), 0)
    Planilha2.Cells(8, 22) = IIf(Sheet6.Cells(15, 15) <> "", Sheet6.Cells(15, 15), 0)
      
    'BCO_ATIVO_CART_CAMB
    Planilha2.Cells(2, 23) = IIf(Sheet6.Cells(16, 3) <> "", Sheet6.Cells(16, 3), 0)
    Planilha2.Cells(3, 23) = IIf(Sheet6.Cells(16, 5) <> "", Sheet6.Cells(16, 5), 0)
    Planilha2.Cells(4, 23) = IIf(Sheet6.Cells(16, 7) <> "", Sheet6.Cells(16, 7), 0)
    Planilha2.Cells(5, 23) = IIf(Sheet6.Cells(16, 9) <> "", Sheet6.Cells(16, 9), 0)
    Planilha2.Cells(6, 23) = IIf(Sheet6.Cells(16, 11) <> "", Sheet6.Cells(16, 11), 0)
    Planilha2.Cells(7, 23) = IIf(Sheet6.Cells(16, 13) <> "", Sheet6.Cells(16, 13), 0)
    Planilha2.Cells(8, 23) = IIf(Sheet6.Cells(16, 15) <> "", Sheet6.Cells(16, 15), 0)

    'BCO_ATIVO_OUTROS_CREDS
    Planilha2.Cells(2, 24) = IIf(Sheet6.Cells(17, 3) <> "", Sheet6.Cells(17, 3), 0)
    Planilha2.Cells(3, 24) = IIf(Sheet6.Cells(17, 5) <> "", Sheet6.Cells(17, 5), 0)
    Planilha2.Cells(4, 24) = IIf(Sheet6.Cells(17, 7) <> "", Sheet6.Cells(17, 7), 0)
    Planilha2.Cells(5, 24) = IIf(Sheet6.Cells(17, 9) <> "", Sheet6.Cells(17, 9), 0)
    Planilha2.Cells(6, 24) = IIf(Sheet6.Cells(17, 11) <> "", Sheet6.Cells(17, 11), 0)
    Planilha2.Cells(7, 24) = IIf(Sheet6.Cells(17, 13) <> "", Sheet6.Cells(17, 13), 0)
    Planilha2.Cells(8, 24) = IIf(Sheet6.Cells(17, 15) <> "", Sheet6.Cells(17, 15), 0)

    'BCO_ATIVO_CIRC
    Planilha2.Cells(2, 25) = IIf(Sheet6.Cells(18, 3) <> "", Sheet6.Cells(18, 3), 0)
    Planilha2.Cells(3, 25) = IIf(Sheet6.Cells(18, 5) <> "", Sheet6.Cells(18, 5), 0)
    Planilha2.Cells(4, 25) = IIf(Sheet6.Cells(18, 7) <> "", Sheet6.Cells(18, 7), 0)
    Planilha2.Cells(5, 25) = IIf(Sheet6.Cells(18, 9) <> "", Sheet6.Cells(18, 9), 0)
    Planilha2.Cells(6, 25) = IIf(Sheet6.Cells(18, 11) <> "", Sheet6.Cells(18, 11), 0)
    Planilha2.Cells(7, 25) = IIf(Sheet6.Cells(18, 13) <> "", Sheet6.Cells(18, 13), 0)
    Planilha2.Cells(8, 25) = IIf(Sheet6.Cells(18, 15) <> "", Sheet6.Cells(18, 15), 0)
    
End Sub

Public Sub Bancos_Mil_Grp2()

    'Nao Circulante

    'BCO_ATIVO_CIRC_TVM

    Planilha2.Cells(2, 26) = IIf(Sheet6.Cells(19, 3) <> "", Sheet6.Cells(19, 3), 0)
    Planilha2.Cells(3, 26) = IIf(Sheet6.Cells(19, 5) <> "", Sheet6.Cells(19, 5), 0)
    Planilha2.Cells(4, 26) = IIf(Sheet6.Cells(19, 7) <> "", Sheet6.Cells(19, 7), 0)
    Planilha2.Cells(5, 26) = IIf(Sheet6.Cells(19, 9) <> "", Sheet6.Cells(19, 9), 0)
    Planilha2.Cells(6, 26) = IIf(Sheet6.Cells(19, 11) <> "", Sheet6.Cells(19, 11), 0)
    Planilha2.Cells(7, 26) = IIf(Sheet6.Cells(19, 13) <> "", Sheet6.Cells(19, 13), 0)
    Planilha2.Cells(8, 26) = IIf(Sheet6.Cells(19, 15) <> "", Sheet6.Cells(19, 15), 0)

    'BCO_ATIVO_CIRC_OPERAC_CRED
    Planilha2.Cells(2, 27) = IIf(Sheet6.Cells(20, 3) <> "", Sheet6.Cells(20, 3), 0)
    Planilha2.Cells(3, 27) = IIf(Sheet6.Cells(20, 5) <> "", Sheet6.Cells(20, 5), 0)
    Planilha2.Cells(4, 27) = IIf(Sheet6.Cells(20, 7) <> "", Sheet6.Cells(20, 7), 0)
    Planilha2.Cells(5, 27) = IIf(Sheet6.Cells(20, 9) <> "", Sheet6.Cells(20, 9), 0)
    Planilha2.Cells(6, 27) = IIf(Sheet6.Cells(20, 11) <> "", Sheet6.Cells(20, 11), 0)
    Planilha2.Cells(7, 27) = IIf(Sheet6.Cells(20, 13) <> "", Sheet6.Cells(20, 13), 0)
    Planilha2.Cells(8, 27) = IIf(Sheet6.Cells(20, 15) <> "", Sheet6.Cells(20, 15), 0)
       
    'BCO_ATIVO_CIRC_PDD_OP_CRED
    Planilha2.Cells(2, 28) = IIf(Sheet6.Cells(21, 3) <> "", Sheet6.Cells(21, 3), 0)
    Planilha2.Cells(3, 28) = IIf(Sheet6.Cells(21, 5) <> "", Sheet6.Cells(21, 5), 0)
    Planilha2.Cells(4, 28) = IIf(Sheet6.Cells(21, 7) <> "", Sheet6.Cells(21, 7), 0)
    Planilha2.Cells(5, 28) = IIf(Sheet6.Cells(21, 9) <> "", Sheet6.Cells(21, 9), 0)
    Planilha2.Cells(6, 28) = IIf(Sheet6.Cells(21, 11) <> "", Sheet6.Cells(21, 11), 0)
    Planilha2.Cells(7, 28) = IIf(Sheet6.Cells(21, 13) <> "", Sheet6.Cells(21, 13), 0)
    Planilha2.Cells(8, 28) = IIf(Sheet6.Cells(21, 15) <> "", Sheet6.Cells(21, 15), 0)
    
    'BCO_ATIVO_CIRC_OP_ARREND_MERC
    Planilha2.Cells(2, 29) = IIf(Sheet6.Cells(22, 3) <> "", Sheet6.Cells(22, 3), 0)
    Planilha2.Cells(3, 29) = IIf(Sheet6.Cells(22, 5) <> "", Sheet6.Cells(22, 5), 0)
    Planilha2.Cells(4, 29) = IIf(Sheet6.Cells(22, 7) <> "", Sheet6.Cells(22, 7), 0)
    Planilha2.Cells(5, 29) = IIf(Sheet6.Cells(22, 9) <> "", Sheet6.Cells(22, 9), 0)
    Planilha2.Cells(6, 29) = IIf(Sheet6.Cells(22, 11) <> "", Sheet6.Cells(22, 11), 0)
    Planilha2.Cells(7, 29) = IIf(Sheet6.Cells(22, 13) <> "", Sheet6.Cells(22, 13), 0)
    Planilha2.Cells(8, 29) = IIf(Sheet6.Cells(22, 15) <> "", Sheet6.Cells(22, 15), 0)

    'BCO_ATIVO_CIRC_PDD_OP_ARR_MERC
    Planilha2.Cells(2, 30) = IIf(Sheet6.Cells(23, 3) <> "", Sheet6.Cells(23, 3), 0)
    Planilha2.Cells(3, 30) = IIf(Sheet6.Cells(23, 5) <> "", Sheet6.Cells(23, 5), 0)
    Planilha2.Cells(4, 30) = IIf(Sheet6.Cells(23, 7) <> "", Sheet6.Cells(23, 7), 0)
    Planilha2.Cells(5, 30) = IIf(Sheet6.Cells(23, 9) <> "", Sheet6.Cells(23, 9), 0)
    Planilha2.Cells(6, 30) = IIf(Sheet6.Cells(23, 11) <> "", Sheet6.Cells(23, 11), 0)
    Planilha2.Cells(7, 30) = IIf(Sheet6.Cells(23, 13) <> "", Sheet6.Cells(23, 13), 0)
    Planilha2.Cells(8, 30) = IIf(Sheet6.Cells(23, 15) <> "", Sheet6.Cells(23, 15), 0)

    'BCO_ATIVO_CIRC_OUTROS_CRED
    Planilha2.Cells(2, 31) = IIf(Sheet6.Cells(24, 3) <> "", Sheet6.Cells(24, 3), 0)
    Planilha2.Cells(3, 31) = IIf(Sheet6.Cells(24, 5) <> "", Sheet6.Cells(24, 5), 0)
    Planilha2.Cells(4, 31) = IIf(Sheet6.Cells(24, 7) <> "", Sheet6.Cells(24, 7), 0)
    Planilha2.Cells(5, 31) = IIf(Sheet6.Cells(24, 9) <> "", Sheet6.Cells(24, 9), 0)
    Planilha2.Cells(6, 31) = IIf(Sheet6.Cells(24, 11) <> "", Sheet6.Cells(24, 11), 0)
    Planilha2.Cells(7, 31) = IIf(Sheet6.Cells(24, 13) <> "", Sheet6.Cells(24, 13), 0)
    Planilha2.Cells(8, 31) = IIf(Sheet6.Cells(24, 15) <> "", Sheet6.Cells(24, 15), 0)
    
    'BCO_ATV_N_CIRC
    Planilha2.Cells(2, 32) = IIf(Sheet6.Cells(25, 3) <> "", Sheet6.Cells(25, 3), 0)
    Planilha2.Cells(3, 32) = IIf(Sheet6.Cells(25, 5) <> "", Sheet6.Cells(25, 5), 0)
    Planilha2.Cells(4, 32) = IIf(Sheet6.Cells(25, 7) <> "", Sheet6.Cells(25, 7), 0)
    Planilha2.Cells(5, 32) = IIf(Sheet6.Cells(25, 9) <> "", Sheet6.Cells(25, 9), 0)
    Planilha2.Cells(6, 32) = IIf(Sheet6.Cells(25, 11) <> "", Sheet6.Cells(25, 11), 0)
    Planilha2.Cells(7, 32) = IIf(Sheet6.Cells(25, 13) <> "", Sheet6.Cells(25, 13), 0)
    Planilha2.Cells(8, 32) = IIf(Sheet6.Cells(25, 15) <> "", Sheet6.Cells(25, 15), 0)

End Sub

Public Sub Bancos_Mil_Grp3()

    'Passivo Total

    'BCO_ATV_N_CIRC_PART_CTRL_COLIG

    Planilha2.Cells(2, 33) = IIf(Sheet6.Cells(26, 3) <> "", Sheet6.Cells(26, 3), 0)
    Planilha2.Cells(3, 33) = IIf(Sheet6.Cells(26, 5) <> "", Sheet6.Cells(26, 5), 0)
    Planilha2.Cells(4, 33) = IIf(Sheet6.Cells(26, 7) <> "", Sheet6.Cells(26, 7), 0)
    Planilha2.Cells(5, 33) = IIf(Sheet6.Cells(26, 9) <> "", Sheet6.Cells(26, 9), 0)
    Planilha2.Cells(6, 33) = IIf(Sheet6.Cells(26, 11) <> "", Sheet6.Cells(26, 11), 0)
    Planilha2.Cells(7, 33) = IIf(Sheet6.Cells(26, 13) <> "", Sheet6.Cells(26, 13), 0)
    Planilha2.Cells(8, 33) = IIf(Sheet6.Cells(26, 15) <> "", Sheet6.Cells(26, 15), 0)

   'BCO_ATV_N_CIRC_OUTROS_INVEST
    Planilha2.Cells(2, 34) = IIf(Sheet6.Cells(27, 3) <> "", Sheet6.Cells(27, 3), 0)
    Planilha2.Cells(3, 34) = IIf(Sheet6.Cells(27, 5) <> "", Sheet6.Cells(27, 5), 0)
    Planilha2.Cells(4, 34) = IIf(Sheet6.Cells(27, 7) <> "", Sheet6.Cells(27, 7), 0)
    Planilha2.Cells(5, 34) = IIf(Sheet6.Cells(27, 9) <> "", Sheet6.Cells(27, 9), 0)
    Planilha2.Cells(6, 34) = IIf(Sheet6.Cells(27, 11) <> "", Sheet6.Cells(27, 11), 0)
    Planilha2.Cells(7, 34) = IIf(Sheet6.Cells(27, 13) <> "", Sheet6.Cells(27, 13), 0)
    Planilha2.Cells(8, 34) = IIf(Sheet6.Cells(27, 15) <> "", Sheet6.Cells(27, 15), 0)
   
   'BCO_ATV_N_CIRC_INVEST
    Planilha2.Cells(2, 35) = IIf(Sheet6.Cells(28, 3) <> "", Sheet6.Cells(28, 3), 0)
    Planilha2.Cells(3, 35) = IIf(Sheet6.Cells(28, 5) <> "", Sheet6.Cells(28, 5), 0)
    Planilha2.Cells(4, 35) = IIf(Sheet6.Cells(28, 7) <> "", Sheet6.Cells(28, 7), 0)
    Planilha2.Cells(5, 35) = IIf(Sheet6.Cells(28, 9) <> "", Sheet6.Cells(28, 9), 0)
    Planilha2.Cells(6, 35) = IIf(Sheet6.Cells(28, 11) <> "", Sheet6.Cells(28, 11), 0)
    Planilha2.Cells(7, 35) = IIf(Sheet6.Cells(28, 13) <> "", Sheet6.Cells(28, 13), 0)
    Planilha2.Cells(8, 35) = IIf(Sheet6.Cells(28, 15) <> "", Sheet6.Cells(28, 15), 0)
   
   'BCO_ATV_N_CIRC_IMOB_TEC_LIQ
    Planilha2.Cells(2, 36) = IIf(Sheet6.Cells(29, 3) <> "", Sheet6.Cells(29, 3), 0)
    Planilha2.Cells(3, 36) = IIf(Sheet6.Cells(29, 5) <> "", Sheet6.Cells(29, 5), 0)
    Planilha2.Cells(4, 36) = IIf(Sheet6.Cells(29, 7) <> "", Sheet6.Cells(29, 7), 0)
    Planilha2.Cells(5, 36) = IIf(Sheet6.Cells(29, 9) <> "", Sheet6.Cells(29, 9), 0)
    Planilha2.Cells(6, 36) = IIf(Sheet6.Cells(29, 11) <> "", Sheet6.Cells(29, 11), 0)
    Planilha2.Cells(7, 36) = IIf(Sheet6.Cells(29, 13) <> "", Sheet6.Cells(29, 13), 0)
    Planilha2.Cells(8, 36) = IIf(Sheet6.Cells(29, 15) <> "", Sheet6.Cells(29, 15), 0)

   'BCO_ATV_N_CIRC_ATV_INTANG
    Planilha2.Cells(2, 37) = IIf(Sheet6.Cells(30, 3) <> "", Sheet6.Cells(30, 3), 0)
    Planilha2.Cells(3, 37) = IIf(Sheet6.Cells(30, 5) <> "", Sheet6.Cells(30, 5), 0)
    Planilha2.Cells(4, 37) = IIf(Sheet6.Cells(30, 7) <> "", Sheet6.Cells(30, 7), 0)
    Planilha2.Cells(5, 37) = IIf(Sheet6.Cells(30, 9) <> "", Sheet6.Cells(30, 9), 0)
    Planilha2.Cells(6, 37) = IIf(Sheet6.Cells(30, 11) <> "", Sheet6.Cells(30, 11), 0)
    Planilha2.Cells(7, 37) = IIf(Sheet6.Cells(30, 13) <> "", Sheet6.Cells(30, 13), 0)
    Planilha2.Cells(8, 37) = IIf(Sheet6.Cells(30, 15) <> "", Sheet6.Cells(30, 15), 0)

   'BCO_ATV_N_CIRC_ATV_PERMAN
    Planilha2.Cells(2, 214) = IIf(Sheet6.Cells(31, 3) <> "", Sheet6.Cells(31, 3), 0)
    Planilha2.Cells(3, 214) = IIf(Sheet6.Cells(31, 5) <> "", Sheet6.Cells(31, 5), 0)
    Planilha2.Cells(4, 214) = IIf(Sheet6.Cells(31, 7) <> "", Sheet6.Cells(31, 7), 0)
    Planilha2.Cells(5, 214) = IIf(Sheet6.Cells(31, 9) <> "", Sheet6.Cells(31, 9), 0)
    Planilha2.Cells(6, 214) = IIf(Sheet6.Cells(31, 11) <> "", Sheet6.Cells(31, 11), 0)
    Planilha2.Cells(7, 214) = IIf(Sheet6.Cells(31, 13) <> "", Sheet6.Cells(31, 13), 0)
    Planilha2.Cells(8, 214) = IIf(Sheet6.Cells(31, 15) <> "", Sheet6.Cells(31, 15), 0)

    ' Bco_Atv_N_Circ_Atv_Total
    Planilha2.Cells(2, 202) = IIf(Sheet6.Cells(32, 3) <> "", Sheet6.Cells(32, 3), 0)
    Planilha2.Cells(3, 202) = IIf(Sheet6.Cells(32, 5) <> "", Sheet6.Cells(32, 5), 0)
    Planilha2.Cells(4, 202) = IIf(Sheet6.Cells(32, 7) <> "", Sheet6.Cells(32, 7), 0)
    Planilha2.Cells(5, 202) = IIf(Sheet6.Cells(32, 9) <> "", Sheet6.Cells(32, 9), 0)
    Planilha2.Cells(6, 202) = IIf(Sheet6.Cells(32, 11) <> "", Sheet6.Cells(32, 11), 0)
    Planilha2.Cells(7, 202) = IIf(Sheet6.Cells(32, 13) <> "", Sheet6.Cells(32, 13), 0)
    Planilha2.Cells(8, 202) = IIf(Sheet6.Cells(32, 15) <> "", Sheet6.Cells(32, 15), 0)

End Sub

Public Sub Bancos_Mil_Grp4()

    '------2 - DEMONSTRATIVO DE RESULTADO-------------

   'BCO_DRE_REC_INTERM_FINANC
    Planilha2.Cells(2, 38) = IIf(Sheet6.Cells(37, 3) <> "", Sheet6.Cells(37, 3), 0)
    Planilha2.Cells(3, 38) = IIf(Sheet6.Cells(37, 5) <> "", Sheet6.Cells(37, 5), 0)
    Planilha2.Cells(4, 38) = IIf(Sheet6.Cells(37, 7) <> "", Sheet6.Cells(37, 7), 0)
    Planilha2.Cells(5, 38) = IIf(Sheet6.Cells(37, 9) <> "", Sheet6.Cells(37, 9), 0)
    Planilha2.Cells(6, 38) = IIf(Sheet6.Cells(37, 11) <> "", Sheet6.Cells(37, 11), 0)
    Planilha2.Cells(7, 38) = IIf(Sheet6.Cells(37, 13) <> "", Sheet6.Cells(37, 13), 0)
    Planilha2.Cells(8, 38) = IIf(Sheet6.Cells(37, 15) <> "", Sheet6.Cells(37, 15), 0)

    ' Bco_Dre_Operac_Cred
    Planilha2.Cells(2, 39) = IIf(Sheet6.Cells(38, 3) <> "", Sheet6.Cells(38, 3), 0)
    Planilha2.Cells(3, 39) = IIf(Sheet6.Cells(38, 5) <> "", Sheet6.Cells(38, 5), 0)
    Planilha2.Cells(4, 39) = IIf(Sheet6.Cells(38, 7) <> "", Sheet6.Cells(38, 7), 0)
    Planilha2.Cells(5, 39) = IIf(Sheet6.Cells(38, 9) <> "", Sheet6.Cells(38, 9), 0)
    Planilha2.Cells(6, 39) = IIf(Sheet6.Cells(38, 11) <> "", Sheet6.Cells(38, 11), 0)
    Planilha2.Cells(7, 39) = IIf(Sheet6.Cells(38, 13) <> "", Sheet6.Cells(38, 13), 0)
    Planilha2.Cells(8, 39) = IIf(Sheet6.Cells(38, 15) <> "", Sheet6.Cells(38, 15), 0)

    ' BCO_DRE_TVM
    Planilha2.Cells(2, 40) = IIf(Sheet6.Cells(39, 3) <> "", Sheet6.Cells(39, 3), 0)
    Planilha2.Cells(3, 40) = IIf(Sheet6.Cells(39, 5) <> "", Sheet6.Cells(39, 5), 0)
    Planilha2.Cells(4, 40) = IIf(Sheet6.Cells(39, 7) <> "", Sheet6.Cells(39, 7), 0)
    Planilha2.Cells(5, 40) = IIf(Sheet6.Cells(39, 9) <> "", Sheet6.Cells(39, 9), 0)
    Planilha2.Cells(6, 40) = IIf(Sheet6.Cells(39, 11) <> "", Sheet6.Cells(39, 11), 0)
    Planilha2.Cells(7, 40) = IIf(Sheet6.Cells(39, 13) <> "", Sheet6.Cells(39, 13), 0)
    Planilha2.Cells(8, 40) = IIf(Sheet6.Cells(39, 15) <> "", Sheet6.Cells(39, 15), 0)

    ' BCO_DRE_OUTRAS_REC_INTERM
    Planilha2.Cells(2, 41) = IIf(Sheet6.Cells(40, 3) <> "", Sheet6.Cells(40, 3), 0)
    Planilha2.Cells(3, 41) = IIf(Sheet6.Cells(40, 5) <> "", Sheet6.Cells(40, 5), 0)
    Planilha2.Cells(4, 41) = IIf(Sheet6.Cells(40, 7) <> "", Sheet6.Cells(40, 7), 0)
    Planilha2.Cells(5, 41) = IIf(Sheet6.Cells(40, 9) <> "", Sheet6.Cells(40, 9), 0)
    Planilha2.Cells(6, 41) = IIf(Sheet6.Cells(40, 11) <> "", Sheet6.Cells(40, 11), 0)
    Planilha2.Cells(7, 41) = IIf(Sheet6.Cells(40, 13) <> "", Sheet6.Cells(40, 13), 0)
    Planilha2.Cells(8, 41) = IIf(Sheet6.Cells(40, 15) <> "", Sheet6.Cells(40, 15), 0)
    
    'BCO_DRE_DESP_INTERM_FINANC
    Planilha2.Cells(2, 42) = IIf(Sheet6.Cells(41, 3) <> "", Sheet6.Cells(41, 3), 0)
    Planilha2.Cells(3, 42) = IIf(Sheet6.Cells(41, 5) <> "", Sheet6.Cells(41, 5), 0)
    Planilha2.Cells(4, 42) = IIf(Sheet6.Cells(41, 7) <> "", Sheet6.Cells(41, 7), 0)
    Planilha2.Cells(5, 42) = IIf(Sheet6.Cells(41, 9) <> "", Sheet6.Cells(41, 9), 0)
    Planilha2.Cells(6, 42) = IIf(Sheet6.Cells(41, 11) <> "", Sheet6.Cells(41, 11), 0)
    Planilha2.Cells(7, 42) = IIf(Sheet6.Cells(41, 13) <> "", Sheet6.Cells(41, 13), 0)
    Planilha2.Cells(8, 42) = IIf(Sheet6.Cells(41, 15) <> "", Sheet6.Cells(41, 15), 0)
    
    ' BCO_DRE_CAPT_MERC
    Planilha2.Cells(2, 43) = IIf(Sheet6.Cells(42, 3) <> "", Sheet6.Cells(42, 3), 0)
    Planilha2.Cells(3, 43) = IIf(Sheet6.Cells(42, 5) <> "", Sheet6.Cells(42, 5), 0)
    Planilha2.Cells(4, 43) = IIf(Sheet6.Cells(42, 7) <> "", Sheet6.Cells(42, 7), 0)
    Planilha2.Cells(5, 43) = IIf(Sheet6.Cells(42, 9) <> "", Sheet6.Cells(42, 9), 0)
    Planilha2.Cells(6, 43) = IIf(Sheet6.Cells(42, 11) <> "", Sheet6.Cells(42, 11), 0)
    Planilha2.Cells(7, 43) = IIf(Sheet6.Cells(42, 13) <> "", Sheet6.Cells(42, 13), 0)
    Planilha2.Cells(8, 43) = IIf(Sheet6.Cells(42, 15) <> "", Sheet6.Cells(42, 15), 0)
        
    ' BCO_DRE_CAPT_MERC
    Planilha2.Cells(2, 44) = IIf(Sheet6.Cells(43, 3) <> "", Sheet6.Cells(43, 3), 0)
    Planilha2.Cells(3, 44) = IIf(Sheet6.Cells(43, 5) <> "", Sheet6.Cells(43, 5), 0)
    Planilha2.Cells(4, 44) = IIf(Sheet6.Cells(43, 7) <> "", Sheet6.Cells(43, 7), 0)
    Planilha2.Cells(5, 44) = IIf(Sheet6.Cells(43, 9) <> "", Sheet6.Cells(43, 9), 0)
    Planilha2.Cells(6, 44) = IIf(Sheet6.Cells(43, 11) <> "", Sheet6.Cells(43, 11), 0)
    Planilha2.Cells(7, 44) = IIf(Sheet6.Cells(43, 13) <> "", Sheet6.Cells(43, 13), 0)
    Planilha2.Cells(8, 44) = IIf(Sheet6.Cells(43, 15) <> "", Sheet6.Cells(43, 15), 0)
    
    ' BCO_DRE_OUTRAS_DESP_INTERM
    Planilha2.Cells(2, 45) = IIf(Sheet6.Cells(44, 3) <> "", Sheet6.Cells(44, 3), 0)
    Planilha2.Cells(3, 45) = IIf(Sheet6.Cells(44, 5) <> "", Sheet6.Cells(44, 5), 0)
    Planilha2.Cells(4, 45) = IIf(Sheet6.Cells(44, 7) <> "", Sheet6.Cells(44, 7), 0)
    Planilha2.Cells(5, 45) = IIf(Sheet6.Cells(44, 9) <> "", Sheet6.Cells(44, 9), 0)
    Planilha2.Cells(6, 45) = IIf(Sheet6.Cells(44, 11) <> "", Sheet6.Cells(44, 11), 0)
    Planilha2.Cells(7, 45) = IIf(Sheet6.Cells(44, 13) <> "", Sheet6.Cells(44, 13), 0)
    Planilha2.Cells(8, 45) = IIf(Sheet6.Cells(44, 15) <> "", Sheet6.Cells(44, 15), 0)
    
    ' BCO_DRE_RES_BRUTO_INTERM
    Planilha2.Cells(2, 46) = IIf(Sheet6.Cells(45, 3) <> "", Sheet6.Cells(45, 3), 0)
    Planilha2.Cells(3, 46) = IIf(Sheet6.Cells(45, 5) <> "", Sheet6.Cells(45, 5), 0)
    Planilha2.Cells(4, 46) = IIf(Sheet6.Cells(45, 7) <> "", Sheet6.Cells(45, 7), 0)
    Planilha2.Cells(5, 46) = IIf(Sheet6.Cells(45, 9) <> "", Sheet6.Cells(45, 9), 0)
    Planilha2.Cells(6, 46) = IIf(Sheet6.Cells(45, 11) <> "", Sheet6.Cells(45, 11), 0)
    Planilha2.Cells(7, 46) = IIf(Sheet6.Cells(45, 13) <> "", Sheet6.Cells(45, 13), 0)
    Planilha2.Cells(8, 46) = IIf(Sheet6.Cells(45, 15) <> "", Sheet6.Cells(45, 15), 0)
    
    ' BCO_DRE_CONST_PDD
    Planilha2.Cells(2, 47) = IIf(Sheet6.Cells(46, 3) <> "", Sheet6.Cells(46, 3), 0)
    Planilha2.Cells(3, 47) = IIf(Sheet6.Cells(46, 5) <> "", Sheet6.Cells(46, 5), 0)
    Planilha2.Cells(4, 47) = IIf(Sheet6.Cells(46, 7) <> "", Sheet6.Cells(46, 7), 0)
    Planilha2.Cells(5, 47) = IIf(Sheet6.Cells(46, 9) <> "", Sheet6.Cells(46, 9), 0)
    Planilha2.Cells(6, 47) = IIf(Sheet6.Cells(46, 11) <> "", Sheet6.Cells(46, 11), 0)
    Planilha2.Cells(7, 47) = IIf(Sheet6.Cells(46, 13) <> "", Sheet6.Cells(46, 13), 0)
    Planilha2.Cells(8, 47) = IIf(Sheet6.Cells(46, 15) <> "", Sheet6.Cells(46, 15), 0)
    
    ' BCO_DRE_RES_INTERM_APOS_PDD
    Planilha2.Cells(2, 48) = IIf(Sheet6.Cells(47, 3) <> "", Sheet6.Cells(47, 3), 0)
    Planilha2.Cells(3, 48) = IIf(Sheet6.Cells(47, 5) <> "", Sheet6.Cells(47, 5), 0)
    Planilha2.Cells(4, 48) = IIf(Sheet6.Cells(47, 7) <> "", Sheet6.Cells(47, 7), 0)
    Planilha2.Cells(5, 48) = IIf(Sheet6.Cells(47, 9) <> "", Sheet6.Cells(47, 9), 0)
    Planilha2.Cells(6, 48) = IIf(Sheet6.Cells(47, 11) <> "", Sheet6.Cells(47, 11), 0)
    Planilha2.Cells(7, 48) = IIf(Sheet6.Cells(47, 13) <> "", Sheet6.Cells(47, 13), 0)
    Planilha2.Cells(8, 48) = IIf(Sheet6.Cells(47, 15) <> "", Sheet6.Cells(47, 15), 0)
        
    ' BCO_DRE_RECT_PREST_SERV
    Planilha2.Cells(2, 49) = IIf(Sheet6.Cells(48, 3) <> "", Sheet6.Cells(48, 3), 0)
    Planilha2.Cells(3, 49) = IIf(Sheet6.Cells(48, 5) <> "", Sheet6.Cells(48, 5), 0)
    Planilha2.Cells(4, 49) = IIf(Sheet6.Cells(48, 7) <> "", Sheet6.Cells(48, 7), 0)
    Planilha2.Cells(5, 49) = IIf(Sheet6.Cells(48, 9) <> "", Sheet6.Cells(48, 9), 0)
    Planilha2.Cells(6, 49) = IIf(Sheet6.Cells(48, 11) <> "", Sheet6.Cells(48, 11), 0)
    Planilha2.Cells(7, 49) = IIf(Sheet6.Cells(48, 13) <> "", Sheet6.Cells(48, 13), 0)
    Planilha2.Cells(8, 49) = IIf(Sheet6.Cells(48, 15) <> "", Sheet6.Cells(48, 15), 0)

    ' BCO_DRE_CUSTO_OPERAC
    Planilha2.Cells(2, 50) = IIf(Sheet6.Cells(49, 3) <> "", Sheet6.Cells(49, 3), 0)
    Planilha2.Cells(3, 50) = IIf(Sheet6.Cells(49, 5) <> "", Sheet6.Cells(49, 5), 0)
    Planilha2.Cells(4, 50) = IIf(Sheet6.Cells(49, 7) <> "", Sheet6.Cells(49, 7), 0)
    Planilha2.Cells(5, 50) = IIf(Sheet6.Cells(49, 9) <> "", Sheet6.Cells(49, 9), 0)
    Planilha2.Cells(6, 50) = IIf(Sheet6.Cells(49, 11) <> "", Sheet6.Cells(49, 11), 0)
    Planilha2.Cells(7, 50) = IIf(Sheet6.Cells(49, 13) <> "", Sheet6.Cells(49, 13), 0)
    Planilha2.Cells(8, 50) = IIf(Sheet6.Cells(49, 15) <> "", Sheet6.Cells(49, 15), 0)

    ' BCO_DRE_DESP_TRIBUT
    Planilha2.Cells(2, 51) = IIf(Sheet6.Cells(50, 3) <> "", Sheet6.Cells(50, 3), 0)
    Planilha2.Cells(3, 51) = IIf(Sheet6.Cells(50, 5) <> "", Sheet6.Cells(50, 5), 0)
    Planilha2.Cells(4, 51) = IIf(Sheet6.Cells(50, 7) <> "", Sheet6.Cells(50, 7), 0)
    Planilha2.Cells(5, 51) = IIf(Sheet6.Cells(50, 9) <> "", Sheet6.Cells(50, 9), 0)
    Planilha2.Cells(6, 51) = IIf(Sheet6.Cells(50, 11) <> "", Sheet6.Cells(50, 11), 0)
    Planilha2.Cells(7, 51) = IIf(Sheet6.Cells(50, 13) <> "", Sheet6.Cells(50, 13), 0)
    Planilha2.Cells(8, 51) = IIf(Sheet6.Cells(50, 15) <> "", Sheet6.Cells(50, 15), 0)
 
    ' Bco_Dre_Outras_Rect_Desp_Operac
    Planilha2.Cells(2, 52) = IIf(Sheet6.Cells(51, 3) <> "", Sheet6.Cells(51, 3), 0)
    Planilha2.Cells(3, 52) = IIf(Sheet6.Cells(51, 5) <> "", Sheet6.Cells(51, 5), 0)
    Planilha2.Cells(4, 52) = IIf(Sheet6.Cells(51, 7) <> "", Sheet6.Cells(51, 7), 0)
    Planilha2.Cells(5, 52) = IIf(Sheet6.Cells(51, 9) <> "", Sheet6.Cells(51, 9), 0)
    Planilha2.Cells(6, 52) = IIf(Sheet6.Cells(51, 11) <> "", Sheet6.Cells(51, 11), 0)
    Planilha2.Cells(7, 52) = IIf(Sheet6.Cells(51, 13) <> "", Sheet6.Cells(51, 13), 0)
    Planilha2.Cells(8, 52) = IIf(Sheet6.Cells(51, 15) <> "", Sheet6.Cells(51, 15), 0)
    
    ' Bco_Dre_Res_Operac
    Planilha2.Cells(2, 53) = IIf(Sheet6.Cells(52, 3) <> "", Sheet6.Cells(52, 3), 0)
    Planilha2.Cells(3, 53) = IIf(Sheet6.Cells(52, 5) <> "", Sheet6.Cells(52, 5), 0)
    Planilha2.Cells(4, 53) = IIf(Sheet6.Cells(52, 7) <> "", Sheet6.Cells(52, 7), 0)
    Planilha2.Cells(5, 53) = IIf(Sheet6.Cells(52, 9) <> "", Sheet6.Cells(52, 9), 0)
    Planilha2.Cells(6, 53) = IIf(Sheet6.Cells(52, 11) <> "", Sheet6.Cells(52, 11), 0)
    Planilha2.Cells(7, 53) = IIf(Sheet6.Cells(52, 13) <> "", Sheet6.Cells(52, 13), 0)
    Planilha2.Cells(8, 53) = IIf(Sheet6.Cells(52, 15) <> "", Sheet6.Cells(52, 15), 0)
       
    ' Bco_Dre_Equiv_Patrim
    Planilha2.Cells(2, 54) = IIf(Sheet6.Cells(53, 3) <> "", Sheet6.Cells(53, 3), 0)
    Planilha2.Cells(3, 54) = IIf(Sheet6.Cells(53, 5) <> "", Sheet6.Cells(53, 5), 0)
    Planilha2.Cells(4, 54) = IIf(Sheet6.Cells(53, 7) <> "", Sheet6.Cells(53, 7), 0)
    Planilha2.Cells(5, 54) = IIf(Sheet6.Cells(53, 9) <> "", Sheet6.Cells(53, 9), 0)
    Planilha2.Cells(6, 54) = IIf(Sheet6.Cells(53, 11) <> "", Sheet6.Cells(53, 11), 0)
    Planilha2.Cells(7, 54) = IIf(Sheet6.Cells(53, 13) <> "", Sheet6.Cells(53, 13), 0)
    Planilha2.Cells(8, 54) = IIf(Sheet6.Cells(53, 15) <> "", Sheet6.Cells(53, 15), 0)
    
    ' Bco_Dre_Res_Apos_Equiv_Patrim
    Planilha2.Cells(2, 55) = IIf(Sheet6.Cells(54, 3) <> "", Sheet6.Cells(54, 3), 0)
    Planilha2.Cells(3, 55) = IIf(Sheet6.Cells(54, 5) <> "", Sheet6.Cells(54, 5), 0)
    Planilha2.Cells(4, 55) = IIf(Sheet6.Cells(54, 7) <> "", Sheet6.Cells(54, 7), 0)
    Planilha2.Cells(5, 55) = IIf(Sheet6.Cells(54, 9) <> "", Sheet6.Cells(54, 9), 0)
    Planilha2.Cells(6, 55) = IIf(Sheet6.Cells(54, 11) <> "", Sheet6.Cells(54, 11), 0)
    Planilha2.Cells(7, 55) = IIf(Sheet6.Cells(54, 13) <> "", Sheet6.Cells(54, 13), 0)
    Planilha2.Cells(8, 55) = IIf(Sheet6.Cells(54, 15) <> "", Sheet6.Cells(54, 15), 0)
    
    ' Bco_Dre_Rect_Desp_N_Operac
    Planilha2.Cells(2, 56) = IIf(Sheet6.Cells(55, 3) <> "", Sheet6.Cells(55, 3), 0)
    Planilha2.Cells(3, 56) = IIf(Sheet6.Cells(55, 5) <> "", Sheet6.Cells(55, 5), 0)
    Planilha2.Cells(4, 56) = IIf(Sheet6.Cells(55, 7) <> "", Sheet6.Cells(55, 7), 0)
    Planilha2.Cells(5, 56) = IIf(Sheet6.Cells(55, 9) <> "", Sheet6.Cells(55, 9), 0)
    Planilha2.Cells(6, 56) = IIf(Sheet6.Cells(55, 11) <> "", Sheet6.Cells(55, 11), 0)
    Planilha2.Cells(7, 56) = IIf(Sheet6.Cells(55, 13) <> "", Sheet6.Cells(55, 13), 0)
    Planilha2.Cells(8, 56) = IIf(Sheet6.Cells(55, 15) <> "", Sheet6.Cells(55, 15), 0)
    
    ' Bco_Dre_Lucro_Antes_Ir
    Planilha2.Cells(2, 57) = IIf(Sheet6.Cells(56, 3) <> "", Sheet6.Cells(56, 3), 0)
    Planilha2.Cells(3, 57) = IIf(Sheet6.Cells(56, 5) <> "", Sheet6.Cells(56, 5), 0)
    Planilha2.Cells(4, 57) = IIf(Sheet6.Cells(56, 7) <> "", Sheet6.Cells(56, 7), 0)
    Planilha2.Cells(5, 57) = IIf(Sheet6.Cells(56, 9) <> "", Sheet6.Cells(56, 9), 0)
    Planilha2.Cells(6, 57) = IIf(Sheet6.Cells(56, 11) <> "", Sheet6.Cells(56, 11), 0)
    Planilha2.Cells(7, 57) = IIf(Sheet6.Cells(56, 13) <> "", Sheet6.Cells(56, 13), 0)
    Planilha2.Cells(8, 57) = IIf(Sheet6.Cells(56, 15) <> "", Sheet6.Cells(56, 15), 0)
    
    ' Bco_Dre_Impst_Renda_Ctrl_Soc
    Planilha2.Cells(2, 58) = IIf(Sheet6.Cells(57, 3) <> "", Sheet6.Cells(57, 3), 0)
    Planilha2.Cells(3, 58) = IIf(Sheet6.Cells(57, 5) <> "", Sheet6.Cells(57, 5), 0)
    Planilha2.Cells(4, 58) = IIf(Sheet6.Cells(57, 7) <> "", Sheet6.Cells(57, 7), 0)
    Planilha2.Cells(5, 58) = IIf(Sheet6.Cells(57, 9) <> "", Sheet6.Cells(57, 9), 0)
    Planilha2.Cells(6, 58) = IIf(Sheet6.Cells(57, 11) <> "", Sheet6.Cells(57, 11), 0)
    Planilha2.Cells(7, 58) = IIf(Sheet6.Cells(57, 13) <> "", Sheet6.Cells(57, 13), 0)
    Planilha2.Cells(8, 58) = IIf(Sheet6.Cells(57, 15) <> "", Sheet6.Cells(57, 15), 0)
    
    ' Bco_Dre_Part
    Planilha2.Cells(2, 59) = IIf(Sheet6.Cells(58, 3) <> "", Sheet6.Cells(58, 3), 0)
    Planilha2.Cells(3, 59) = IIf(Sheet6.Cells(58, 5) <> "", Sheet6.Cells(58, 5), 0)
    Planilha2.Cells(4, 59) = IIf(Sheet6.Cells(58, 7) <> "", Sheet6.Cells(58, 7), 0)
    Planilha2.Cells(5, 59) = IIf(Sheet6.Cells(58, 9) <> "", Sheet6.Cells(58, 9), 0)
    Planilha2.Cells(6, 59) = IIf(Sheet6.Cells(58, 11) <> "", Sheet6.Cells(58, 11), 0)
    Planilha2.Cells(7, 59) = IIf(Sheet6.Cells(58, 13) <> "", Sheet6.Cells(58, 13), 0)
    Planilha2.Cells(8, 59) = IIf(Sheet6.Cells(58, 15) <> "", Sheet6.Cells(58, 15), 0)
    
    ' BCO_DRE_LUCRO_LIQ
    Planilha2.Cells(2, 60) = IIf(Sheet6.Cells(59, 3) <> "", Sheet6.Cells(59, 3), 0)
    Planilha2.Cells(3, 60) = IIf(Sheet6.Cells(59, 5) <> "", Sheet6.Cells(59, 5), 0)
    Planilha2.Cells(4, 60) = IIf(Sheet6.Cells(59, 7) <> "", Sheet6.Cells(59, 7), 0)
    Planilha2.Cells(5, 60) = IIf(Sheet6.Cells(59, 9) <> "", Sheet6.Cells(59, 9), 0)
    Planilha2.Cells(6, 60) = IIf(Sheet6.Cells(59, 11) <> "", Sheet6.Cells(59, 11), 0)
    Planilha2.Cells(7, 60) = IIf(Sheet6.Cells(59, 13) <> "", Sheet6.Cells(59, 13), 0)
    Planilha2.Cells(8, 60) = IIf(Sheet6.Cells(59, 15) <> "", Sheet6.Cells(59, 15), 0)

End Sub

Public Sub Bancos_Mil_Grp6()

    '---------------- Passivo  ----------------------

    'Bco_Pass_Depos_Avista
    Planilha2.Cells(2, 61) = IIf(Sheet6.Cells(7, 19) <> "", Sheet6.Cells(7, 19), 0)
    Planilha2.Cells(3, 61) = IIf(Sheet6.Cells(7, 21) <> "", Sheet6.Cells(7, 21), 0)
    Planilha2.Cells(4, 61) = IIf(Sheet6.Cells(7, 23) <> "", Sheet6.Cells(7, 23), 0)
    Planilha2.Cells(5, 61) = IIf(Sheet6.Cells(7, 25) <> "", Sheet6.Cells(7, 25), 0)
    Planilha2.Cells(6, 61) = IIf(Sheet6.Cells(7, 27) <> "", Sheet6.Cells(7, 27), 0)
    Planilha2.Cells(7, 61) = IIf(Sheet6.Cells(7, 29) <> "", Sheet6.Cells(7, 29), 0)
    Planilha2.Cells(8, 61) = IIf(Sheet6.Cells(7, 31) <> "", Sheet6.Cells(7, 31), 0)

    'Bco_Pass_Poupanca
    Planilha2.Cells(2, 62) = IIf(Sheet6.Cells(8, 19) <> "", Sheet6.Cells(8, 19), 0)
    Planilha2.Cells(3, 62) = IIf(Sheet6.Cells(8, 21) <> "", Sheet6.Cells(8, 21), 0)
    Planilha2.Cells(4, 62) = IIf(Sheet6.Cells(8, 23) <> "", Sheet6.Cells(8, 23), 0)
    Planilha2.Cells(5, 62) = IIf(Sheet6.Cells(8, 25) <> "", Sheet6.Cells(8, 25), 0)
    Planilha2.Cells(6, 62) = IIf(Sheet6.Cells(8, 27) <> "", Sheet6.Cells(8, 27), 0)
    Planilha2.Cells(7, 62) = IIf(Sheet6.Cells(8, 29) <> "", Sheet6.Cells(8, 29), 0)
    Planilha2.Cells(8, 62) = IIf(Sheet6.Cells(8, 31) <> "", Sheet6.Cells(8, 31), 0)

    'Bco_Pass_Depos_Interfinan
    Planilha2.Cells(2, 63) = IIf(Sheet6.Cells(9, 19) <> "", Sheet6.Cells(9, 19), 0)
    Planilha2.Cells(3, 63) = IIf(Sheet6.Cells(9, 21) <> "", Sheet6.Cells(9, 21), 0)
    Planilha2.Cells(4, 63) = IIf(Sheet6.Cells(9, 23) <> "", Sheet6.Cells(9, 23), 0)
    Planilha2.Cells(5, 63) = IIf(Sheet6.Cells(9, 25) <> "", Sheet6.Cells(9, 25), 0)
    Planilha2.Cells(6, 63) = IIf(Sheet6.Cells(9, 27) <> "", Sheet6.Cells(9, 27), 0)
    Planilha2.Cells(7, 63) = IIf(Sheet6.Cells(9, 29) <> "", Sheet6.Cells(9, 29), 0)
    Planilha2.Cells(8, 63) = IIf(Sheet6.Cells(9, 31) <> "", Sheet6.Cells(9, 31), 0)

    'Bco_Pass_Depos_Aprazo
    Planilha2.Cells(2, 64) = IIf(Sheet6.Cells(10, 19) <> "", Sheet6.Cells(10, 19), 0)
    Planilha2.Cells(3, 64) = IIf(Sheet6.Cells(10, 21) <> "", Sheet6.Cells(10, 21), 0)
    Planilha2.Cells(4, 64) = IIf(Sheet6.Cells(10, 23) <> "", Sheet6.Cells(10, 23), 0)
    Planilha2.Cells(5, 64) = IIf(Sheet6.Cells(10, 25) <> "", Sheet6.Cells(10, 25), 0)
    Planilha2.Cells(6, 64) = IIf(Sheet6.Cells(10, 27) <> "", Sheet6.Cells(10, 27), 0)
    Planilha2.Cells(7, 64) = IIf(Sheet6.Cells(10, 29) <> "", Sheet6.Cells(10, 29), 0)
    Planilha2.Cells(8, 64) = IIf(Sheet6.Cells(10, 31) <> "", Sheet6.Cells(10, 31), 0)

    'Bco_Pass_Capt_Merc_Abert
    Planilha2.Cells(2, 65) = IIf(Sheet6.Cells(11, 19) <> "", Sheet6.Cells(11, 19), 0)
    Planilha2.Cells(3, 65) = IIf(Sheet6.Cells(11, 21) <> "", Sheet6.Cells(11, 21), 0)
    Planilha2.Cells(4, 65) = IIf(Sheet6.Cells(11, 23) <> "", Sheet6.Cells(11, 23), 0)
    Planilha2.Cells(5, 65) = IIf(Sheet6.Cells(11, 25) <> "", Sheet6.Cells(11, 25), 0)
    Planilha2.Cells(6, 65) = IIf(Sheet6.Cells(11, 27) <> "", Sheet6.Cells(11, 27), 0)
    Planilha2.Cells(7, 65) = IIf(Sheet6.Cells(11, 29) <> "", Sheet6.Cells(11, 29), 0)
    Planilha2.Cells(8, 65) = IIf(Sheet6.Cells(11, 31) <> "", Sheet6.Cells(11, 31), 0)

    'Bco_Pass_Emprest_Pais
    Planilha2.Cells(2, 66) = IIf(Sheet6.Cells(12, 19) <> "", Sheet6.Cells(12, 19), 0)
    Planilha2.Cells(3, 66) = IIf(Sheet6.Cells(12, 21) <> "", Sheet6.Cells(12, 21), 0)
    Planilha2.Cells(4, 66) = IIf(Sheet6.Cells(12, 23) <> "", Sheet6.Cells(12, 23), 0)
    Planilha2.Cells(5, 66) = IIf(Sheet6.Cells(12, 25) <> "", Sheet6.Cells(12, 25), 0)
    Planilha2.Cells(6, 66) = IIf(Sheet6.Cells(12, 27) <> "", Sheet6.Cells(12, 27), 0)
    Planilha2.Cells(7, 66) = IIf(Sheet6.Cells(12, 29) <> "", Sheet6.Cells(12, 29), 0)
    Planilha2.Cells(8, 66) = IIf(Sheet6.Cells(12, 31) <> "", Sheet6.Cells(12, 31), 0)

    'Bco_Pass_Repass_Pais
    Planilha2.Cells(2, 67) = IIf(Sheet6.Cells(13, 19) <> "", Sheet6.Cells(13, 19), 0)
    Planilha2.Cells(3, 67) = IIf(Sheet6.Cells(13, 21) <> "", Sheet6.Cells(13, 21), 0)
    Planilha2.Cells(4, 67) = IIf(Sheet6.Cells(13, 23) <> "", Sheet6.Cells(13, 23), 0)
    Planilha2.Cells(5, 67) = IIf(Sheet6.Cells(13, 25) <> "", Sheet6.Cells(13, 25), 0)
    Planilha2.Cells(6, 67) = IIf(Sheet6.Cells(13, 27) <> "", Sheet6.Cells(13, 27), 0)
    Planilha2.Cells(7, 67) = IIf(Sheet6.Cells(13, 29) <> "", Sheet6.Cells(13, 29), 0)
    Planilha2.Cells(8, 67) = IIf(Sheet6.Cells(13, 31) <> "", Sheet6.Cells(13, 31), 0)

    'Bco_Pass_Emprest_Exterior
    Planilha2.Cells(2, 68) = IIf(Sheet6.Cells(14, 19) <> "", Sheet6.Cells(14, 19), 0)
    Planilha2.Cells(3, 68) = IIf(Sheet6.Cells(14, 21) <> "", Sheet6.Cells(14, 21), 0)
    Planilha2.Cells(4, 68) = IIf(Sheet6.Cells(14, 23) <> "", Sheet6.Cells(14, 23), 0)
    Planilha2.Cells(5, 68) = IIf(Sheet6.Cells(14, 25) <> "", Sheet6.Cells(14, 25), 0)
    Planilha2.Cells(6, 68) = IIf(Sheet6.Cells(14, 27) <> "", Sheet6.Cells(14, 27), 0)
    Planilha2.Cells(7, 68) = IIf(Sheet6.Cells(14, 29) <> "", Sheet6.Cells(14, 29), 0)
    Planilha2.Cells(8, 68) = IIf(Sheet6.Cells(14, 31) <> "", Sheet6.Cells(14, 31), 0)

    'Bco_Pass_Repass_Exterior
    Planilha2.Cells(2, 69) = IIf(Sheet6.Cells(15, 19) <> "", Sheet6.Cells(15, 19), 0)
    Planilha2.Cells(3, 69) = IIf(Sheet6.Cells(15, 21) <> "", Sheet6.Cells(15, 21), 0)
    Planilha2.Cells(4, 69) = IIf(Sheet6.Cells(15, 23) <> "", Sheet6.Cells(15, 23), 0)
    Planilha2.Cells(5, 69) = IIf(Sheet6.Cells(15, 25) <> "", Sheet6.Cells(15, 25), 0)
    Planilha2.Cells(6, 69) = IIf(Sheet6.Cells(15, 27) <> "", Sheet6.Cells(15, 27), 0)
    Planilha2.Cells(7, 69) = IIf(Sheet6.Cells(15, 29) <> "", Sheet6.Cells(15, 29), 0)
    Planilha2.Cells(8, 69) = IIf(Sheet6.Cells(15, 31) <> "", Sheet6.Cells(15, 31), 0)

    'Bco_Pass_Cart_Camb
    Planilha2.Cells(2, 70) = IIf(Sheet6.Cells(16, 19) <> "", Sheet6.Cells(16, 19), 0)
    Planilha2.Cells(3, 70) = IIf(Sheet6.Cells(16, 21) <> "", Sheet6.Cells(16, 21), 0)
    Planilha2.Cells(4, 70) = IIf(Sheet6.Cells(16, 23) <> "", Sheet6.Cells(16, 23), 0)
    Planilha2.Cells(5, 70) = IIf(Sheet6.Cells(16, 25) <> "", Sheet6.Cells(16, 25), 0)
    Planilha2.Cells(6, 70) = IIf(Sheet6.Cells(16, 27) <> "", Sheet6.Cells(16, 27), 0)
    Planilha2.Cells(7, 70) = IIf(Sheet6.Cells(16, 29) <> "", Sheet6.Cells(16, 29), 0)
    Planilha2.Cells(8, 70) = IIf(Sheet6.Cells(16, 31) <> "", Sheet6.Cells(16, 31), 0)
    
    'Bco_Pass_Outras_Contas
    Planilha2.Cells(2, 71) = IIf(Sheet6.Cells(17, 19) <> "", Sheet6.Cells(17, 19), 0)
    Planilha2.Cells(3, 71) = IIf(Sheet6.Cells(17, 21) <> "", Sheet6.Cells(17, 21), 0)
    Planilha2.Cells(4, 71) = IIf(Sheet6.Cells(17, 23) <> "", Sheet6.Cells(17, 23), 0)
    Planilha2.Cells(5, 71) = IIf(Sheet6.Cells(17, 25) <> "", Sheet6.Cells(17, 25), 0)
    Planilha2.Cells(6, 71) = IIf(Sheet6.Cells(17, 27) <> "", Sheet6.Cells(17, 27), 0)
    Planilha2.Cells(7, 71) = IIf(Sheet6.Cells(17, 29) <> "", Sheet6.Cells(17, 29), 0)
    Planilha2.Cells(8, 71) = IIf(Sheet6.Cells(17, 31) <> "", Sheet6.Cells(17, 31), 0)

    'BCO_PASS_CIRC
    Planilha2.Cells(2, 72) = IIf(Sheet6.Cells(18, 19) <> "", Sheet6.Cells(18, 19), 0)
    Planilha2.Cells(3, 72) = IIf(Sheet6.Cells(18, 21) <> "", Sheet6.Cells(18, 21), 0)
    Planilha2.Cells(4, 72) = IIf(Sheet6.Cells(18, 23) <> "", Sheet6.Cells(18, 23), 0)
    Planilha2.Cells(5, 72) = IIf(Sheet6.Cells(18, 25) <> "", Sheet6.Cells(18, 25), 0)
    Planilha2.Cells(6, 72) = IIf(Sheet6.Cells(18, 27) <> "", Sheet6.Cells(18, 27), 0)
    Planilha2.Cells(7, 72) = IIf(Sheet6.Cells(18, 29) <> "", Sheet6.Cells(18, 29), 0)
    Planilha2.Cells(8, 72) = IIf(Sheet6.Cells(18, 31) <> "", Sheet6.Cells(18, 31), 0)

End Sub

Public Sub Bancos_Mil_Grp5()

    ' BCO_PASS_DEPOS
    Planilha2.Cells(2, 73) = IIf(Sheet6.Cells(19, 19) <> "", Sheet6.Cells(19, 19), 0)
    Planilha2.Cells(3, 73) = IIf(Sheet6.Cells(19, 21) <> "", Sheet6.Cells(19, 21), 0)
    Planilha2.Cells(4, 73) = IIf(Sheet6.Cells(19, 23) <> "", Sheet6.Cells(19, 23), 0)
    Planilha2.Cells(5, 73) = IIf(Sheet6.Cells(19, 25) <> "", Sheet6.Cells(19, 25), 0)
    Planilha2.Cells(6, 73) = IIf(Sheet6.Cells(19, 27) <> "", Sheet6.Cells(18, 27), 0)
    Planilha2.Cells(7, 73) = IIf(Sheet6.Cells(19, 29) <> "", Sheet6.Cells(19, 29), 0)
    Planilha2.Cells(8, 73) = IIf(Sheet6.Cells(19, 31) <> "", Sheet6.Cells(19, 31), 0)

    ' Bco_Pass_Circ_Emprest_Pais
    Planilha2.Cells(2, 74) = IIf(Sheet6.Cells(20, 19) <> "", Sheet6.Cells(20, 19), 0)
    Planilha2.Cells(3, 74) = IIf(Sheet6.Cells(20, 21) <> "", Sheet6.Cells(20, 21), 0)
    Planilha2.Cells(4, 74) = IIf(Sheet6.Cells(20, 23) <> "", Sheet6.Cells(20, 23), 0)
    Planilha2.Cells(5, 74) = IIf(Sheet6.Cells(20, 25) <> "", Sheet6.Cells(20, 25), 0)
    Planilha2.Cells(6, 74) = IIf(Sheet6.Cells(20, 27) <> "", Sheet6.Cells(20, 27), 0)
    Planilha2.Cells(7, 74) = IIf(Sheet6.Cells(20, 29) <> "", Sheet6.Cells(20, 29), 0)
    Planilha2.Cells(8, 74) = IIf(Sheet6.Cells(20, 31) <> "", Sheet6.Cells(20, 31), 0)
    
    ' Bco_Pass_Circ_Repass_Pais
    Planilha2.Cells(2, 75) = IIf(Sheet6.Cells(21, 19) <> "", Sheet6.Cells(21, 19), 0)
    Planilha2.Cells(3, 75) = IIf(Sheet6.Cells(21, 21) <> "", Sheet6.Cells(21, 21), 0)
    Planilha2.Cells(4, 75) = IIf(Sheet6.Cells(21, 23) <> "", Sheet6.Cells(21, 23), 0)
    Planilha2.Cells(5, 75) = IIf(Sheet6.Cells(21, 25) <> "", Sheet6.Cells(21, 25), 0)
    Planilha2.Cells(6, 75) = IIf(Sheet6.Cells(21, 27) <> "", Sheet6.Cells(21, 27), 0)
    Planilha2.Cells(7, 75) = IIf(Sheet6.Cells(21, 29) <> "", Sheet6.Cells(21, 29), 0)
    Planilha2.Cells(8, 75) = IIf(Sheet6.Cells(21, 31) <> "", Sheet6.Cells(21, 31), 0)
    
    ' Bco_Pass_Circ_Emprest_Exterior
    Planilha2.Cells(2, 76) = IIf(Sheet6.Cells(22, 19) <> "", Sheet6.Cells(22, 19), 0)
    Planilha2.Cells(3, 76) = IIf(Sheet6.Cells(22, 21) <> "", Sheet6.Cells(22, 21), 0)
    Planilha2.Cells(4, 76) = IIf(Sheet6.Cells(22, 23) <> "", Sheet6.Cells(22, 23), 0)
    Planilha2.Cells(5, 76) = IIf(Sheet6.Cells(22, 25) <> "", Sheet6.Cells(22, 25), 0)
    Planilha2.Cells(6, 76) = IIf(Sheet6.Cells(22, 27) <> "", Sheet6.Cells(22, 27), 0)
    Planilha2.Cells(7, 76) = IIf(Sheet6.Cells(22, 29) <> "", Sheet6.Cells(22, 29), 0)
    Planilha2.Cells(8, 76) = IIf(Sheet6.Cells(22, 31) <> "", Sheet6.Cells(22, 31), 0)
    
    ' Bco_Pass_Circ_Repass_Exterior
    Planilha2.Cells(2, 77) = IIf(Sheet6.Cells(23, 19) <> "", Sheet6.Cells(23, 19), 0)
    Planilha2.Cells(3, 77) = IIf(Sheet6.Cells(23, 21) <> "", Sheet6.Cells(23, 21), 0)
    Planilha2.Cells(4, 77) = IIf(Sheet6.Cells(23, 23) <> "", Sheet6.Cells(23, 23), 0)
    Planilha2.Cells(5, 77) = IIf(Sheet6.Cells(23, 25) <> "", Sheet6.Cells(23, 25), 0)
    Planilha2.Cells(6, 77) = IIf(Sheet6.Cells(23, 27) <> "", Sheet6.Cells(23, 27), 0)
    Planilha2.Cells(7, 77) = IIf(Sheet6.Cells(23, 29) <> "", Sheet6.Cells(23, 29), 0)
    Planilha2.Cells(8, 77) = IIf(Sheet6.Cells(23, 31) <> "", Sheet6.Cells(23, 31), 0)
    
    ' Bco_Pass_Circ_Outras_Contas
    Planilha2.Cells(2, 78) = IIf(Sheet6.Cells(24, 19) <> "", Sheet6.Cells(24, 19), 0)
    Planilha2.Cells(3, 78) = IIf(Sheet6.Cells(24, 21) <> "", Sheet6.Cells(24, 21), 0)
    Planilha2.Cells(4, 78) = IIf(Sheet6.Cells(24, 23) <> "", Sheet6.Cells(24, 23), 0)
    Planilha2.Cells(5, 78) = IIf(Sheet6.Cells(24, 25) <> "", Sheet6.Cells(24, 25), 0)
    Planilha2.Cells(6, 78) = IIf(Sheet6.Cells(24, 27) <> "", Sheet6.Cells(24, 27), 0)
    Planilha2.Cells(7, 78) = IIf(Sheet6.Cells(24, 29) <> "", Sheet6.Cells(24, 29), 0)
    Planilha2.Cells(8, 78) = IIf(Sheet6.Cells(24, 31) <> "", Sheet6.Cells(24, 31), 0)
    
    ' Bco_Pass_N_Circ
    Planilha2.Cells(2, 79) = IIf(Sheet6.Cells(25, 19) <> "", Sheet6.Cells(25, 19), 0)
    Planilha2.Cells(3, 79) = IIf(Sheet6.Cells(25, 21) <> "", Sheet6.Cells(25, 21), 0)
    Planilha2.Cells(4, 79) = IIf(Sheet6.Cells(25, 23) <> "", Sheet6.Cells(25, 23), 0)
    Planilha2.Cells(5, 79) = IIf(Sheet6.Cells(25, 25) <> "", Sheet6.Cells(25, 25), 0)
    Planilha2.Cells(6, 79) = IIf(Sheet6.Cells(25, 27) <> "", Sheet6.Cells(25, 27), 0)
    Planilha2.Cells(7, 79) = IIf(Sheet6.Cells(25, 29) <> "", Sheet6.Cells(25, 29), 0)
    Planilha2.Cells(8, 79) = IIf(Sheet6.Cells(25, 31) <> "", Sheet6.Cells(25, 31), 0)
    
        
    ' Bco_Pass_N_Circ_Capit_Soc
    Planilha2.Cells(2, 80) = IIf(Sheet6.Cells(26, 19) <> "", Sheet6.Cells(26, 19), 0)
    Planilha2.Cells(3, 80) = IIf(Sheet6.Cells(26, 21) <> "", Sheet6.Cells(26, 21), 0)
    Planilha2.Cells(4, 80) = IIf(Sheet6.Cells(26, 23) <> "", Sheet6.Cells(26, 23), 0)
    Planilha2.Cells(5, 80) = IIf(Sheet6.Cells(26, 25) <> "", Sheet6.Cells(26, 25), 0)
    Planilha2.Cells(6, 80) = IIf(Sheet6.Cells(26, 27) <> "", Sheet6.Cells(26, 27), 0)
    Planilha2.Cells(7, 80) = IIf(Sheet6.Cells(26, 29) <> "", Sheet6.Cells(26, 29), 0)
    Planilha2.Cells(8, 80) = IIf(Sheet6.Cells(26, 31) <> "", Sheet6.Cells(26, 31), 0)
     
    ' Bco_Pass_N_Circ_Reserv_Capt
    Planilha2.Cells(2, 81) = IIf(Sheet6.Cells(27, 19) <> "", Sheet6.Cells(27, 19), 0)
    Planilha2.Cells(3, 81) = IIf(Sheet6.Cells(27, 21) <> "", Sheet6.Cells(27, 21), 0)
    Planilha2.Cells(4, 81) = IIf(Sheet6.Cells(27, 23) <> "", Sheet6.Cells(27, 23), 0)
    Planilha2.Cells(5, 81) = IIf(Sheet6.Cells(27, 25) <> "", Sheet6.Cells(27, 25), 0)
    Planilha2.Cells(6, 81) = IIf(Sheet6.Cells(27, 27) <> "", Sheet6.Cells(27, 27), 0)
    Planilha2.Cells(7, 81) = IIf(Sheet6.Cells(27, 29) <> "", Sheet6.Cells(27, 29), 0)
    Planilha2.Cells(8, 81) = IIf(Sheet6.Cells(27, 31) <> "", Sheet6.Cells(27, 31), 0)
    
    ' Bco_Pass_N_Circ_Part_Minor
    Planilha2.Cells(2, 82) = IIf(Sheet6.Cells(28, 19) <> "", Sheet6.Cells(28, 19), 0)
    Planilha2.Cells(3, 82) = IIf(Sheet6.Cells(28, 21) <> "", Sheet6.Cells(28, 21), 0)
    Planilha2.Cells(4, 82) = IIf(Sheet6.Cells(28, 23) <> "", Sheet6.Cells(28, 23), 0)
    Planilha2.Cells(5, 82) = IIf(Sheet6.Cells(28, 25) <> "", Sheet6.Cells(28, 25), 0)
    Planilha2.Cells(6, 82) = IIf(Sheet6.Cells(28, 27) <> "", Sheet6.Cells(28, 27), 0)
    Planilha2.Cells(7, 82) = IIf(Sheet6.Cells(28, 29) <> "", Sheet6.Cells(28, 29), 0)
    Planilha2.Cells(8, 82) = IIf(Sheet6.Cells(28, 31) <> "", Sheet6.Cells(28, 31), 0)
    
    ' Bco_Pass_N_Circ_Ajust_Vlr_Merc
    Planilha2.Cells(2, 83) = IIf(Sheet6.Cells(29, 19) <> "", Sheet6.Cells(29, 19), 0)
    Planilha2.Cells(3, 83) = IIf(Sheet6.Cells(29, 21) <> "", Sheet6.Cells(29, 21), 0)
    Planilha2.Cells(4, 83) = IIf(Sheet6.Cells(29, 23) <> "", Sheet6.Cells(29, 23), 0)
    Planilha2.Cells(5, 83) = IIf(Sheet6.Cells(29, 25) <> "", Sheet6.Cells(29, 25), 0)
    Planilha2.Cells(6, 83) = IIf(Sheet6.Cells(29, 27) <> "", Sheet6.Cells(29, 27), 0)
    Planilha2.Cells(7, 83) = IIf(Sheet6.Cells(29, 29) <> "", Sheet6.Cells(29, 29), 0)
    Planilha2.Cells(8, 83) = IIf(Sheet6.Cells(29, 31) <> "", Sheet6.Cells(29, 31), 0)
    
    ' Bco_Pass_N_Circ_Lcr_Prej_Acml
    Planilha2.Cells(2, 84) = IIf(Sheet6.Cells(30, 19) <> "", Sheet6.Cells(30, 19), 0)
    Planilha2.Cells(3, 84) = IIf(Sheet6.Cells(30, 21) <> "", Sheet6.Cells(30, 21), 0)
    Planilha2.Cells(4, 84) = IIf(Sheet6.Cells(30, 23) <> "", Sheet6.Cells(30, 23), 0)
    Planilha2.Cells(5, 84) = IIf(Sheet6.Cells(30, 25) <> "", Sheet6.Cells(30, 25), 0)
    Planilha2.Cells(6, 84) = IIf(Sheet6.Cells(30, 27) <> "", Sheet6.Cells(30, 27), 0)
    Planilha2.Cells(7, 84) = IIf(Sheet6.Cells(30, 29) <> "", Sheet6.Cells(30, 29), 0)
    Planilha2.Cells(8, 84) = IIf(Sheet6.Cells(30, 31) <> "", Sheet6.Cells(30, 31), 0)
    
    ' Bco_Pass_N_Circ_Patrim_Liq
    Planilha2.Cells(2, 85) = IIf(Sheet6.Cells(31, 19) <> "", Sheet6.Cells(31, 19), 0)
    Planilha2.Cells(3, 85) = IIf(Sheet6.Cells(31, 21) <> "", Sheet6.Cells(31, 21), 0)
    Planilha2.Cells(4, 85) = IIf(Sheet6.Cells(31, 23) <> "", Sheet6.Cells(31, 23), 0)
    Planilha2.Cells(5, 85) = IIf(Sheet6.Cells(31, 25) <> "", Sheet6.Cells(31, 25), 0)
    Planilha2.Cells(6, 85) = IIf(Sheet6.Cells(31, 27) <> "", Sheet6.Cells(31, 27), 0)
    Planilha2.Cells(7, 85) = IIf(Sheet6.Cells(31, 29) <> "", Sheet6.Cells(31, 29), 0)
    Planilha2.Cells(8, 85) = IIf(Sheet6.Cells(31, 31) <> "", Sheet6.Cells(31, 31), 0)
    
    ' Bco_Pass_N_Circ_Pass_Total
    Planilha2.Cells(2, 86) = IIf(Sheet6.Cells(32, 19) <> "", Sheet6.Cells(32, 19), 0)
    Planilha2.Cells(3, 86) = IIf(Sheet6.Cells(32, 21) <> "", Sheet6.Cells(32, 21), 0)
    Planilha2.Cells(4, 86) = IIf(Sheet6.Cells(32, 23) <> "", Sheet6.Cells(32, 23), 0)
    Planilha2.Cells(5, 86) = IIf(Sheet6.Cells(32, 25) <> "", Sheet6.Cells(32, 25), 0)
    Planilha2.Cells(6, 86) = IIf(Sheet6.Cells(32, 27) <> "", Sheet6.Cells(32, 27), 0)
    Planilha2.Cells(7, 86) = IIf(Sheet6.Cells(32, 29) <> "", Sheet6.Cells(32, 29), 0)
    Planilha2.Cells(8, 86) = IIf(Sheet6.Cells(32, 31) <> "", Sheet6.Cells(32, 31), 0)

End Sub

Public Sub Bancos_Mil_Grp7()

    'Bco_Bxdos_Sectzdos
    Planilha2.Cells(2, 87) = IIf(Sheet6.Cells(37, 19) <> "", Sheet6.Cells(37, 19), 0)
    Planilha2.Cells(3, 87) = IIf(Sheet6.Cells(37, 21) <> "", Sheet6.Cells(37, 21), 0)
    Planilha2.Cells(4, 87) = IIf(Sheet6.Cells(37, 23) <> "", Sheet6.Cells(37, 23), 0)
    Planilha2.Cells(5, 87) = IIf(Sheet6.Cells(37, 25) <> "", Sheet6.Cells(37, 25), 0)
    Planilha2.Cells(6, 87) = IIf(Sheet6.Cells(37, 27) <> "", Sheet6.Cells(37, 27), 0)
    Planilha2.Cells(7, 87) = IIf(Sheet6.Cells(37, 29) <> "", Sheet6.Cells(37, 29), 0)
    Planilha2.Cells(8, 87) = IIf(Sheet6.Cells(37, 31) <> "", Sheet6.Cells(37, 31), 0)
    
    'Bco_Ind_Basileia_Br
    Planilha2.Cells(2, 88) = IIf(Sheet6.Cells(38, 19) <> "", Sheet6.Cells(38, 19), 0)
    Planilha2.Cells(3, 88) = IIf(Sheet6.Cells(38, 21) <> "", Sheet6.Cells(38, 21), 0)
    Planilha2.Cells(4, 88) = IIf(Sheet6.Cells(38, 23) <> "", Sheet6.Cells(38, 23), 0)
    Planilha2.Cells(5, 88) = IIf(Sheet6.Cells(38, 25) <> "", Sheet6.Cells(38, 25), 0)
    Planilha2.Cells(6, 88) = IIf(Sheet6.Cells(38, 27) <> "", Sheet6.Cells(38, 27), 0)
    Planilha2.Cells(7, 88) = IIf(Sheet6.Cells(38, 29) <> "", Sheet6.Cells(38, 29), 0)
    Planilha2.Cells(8, 88) = IIf(Sheet6.Cells(38, 31) <> "", Sheet6.Cells(38, 31), 0)
    
    'Bco_Basileia_Tier_I
    Planilha2.Cells(2, 89) = IIf(Sheet6.Cells(39, 19) <> "", Sheet6.Cells(39, 19), 0)
    Planilha2.Cells(3, 89) = IIf(Sheet6.Cells(39, 21) <> "", Sheet6.Cells(39, 21), 0)
    Planilha2.Cells(4, 89) = IIf(Sheet6.Cells(39, 23) <> "", Sheet6.Cells(39, 23), 0)
    Planilha2.Cells(5, 89) = IIf(Sheet6.Cells(39, 25) <> "", Sheet6.Cells(39, 25), 0)
    Planilha2.Cells(6, 89) = IIf(Sheet6.Cells(39, 27) <> "", Sheet6.Cells(39, 27), 0)
    Planilha2.Cells(7, 89) = IIf(Sheet6.Cells(39, 29) <> "", Sheet6.Cells(39, 29), 0)
    Planilha2.Cells(8, 89) = IIf(Sheet6.Cells(39, 31) <> "", Sheet6.Cells(39, 31), 0)
    
    'Bco_Dpge_I
    Planilha2.Cells(2, 90) = IIf(Sheet6.Cells(40, 19) <> "", Sheet6.Cells(40, 19), 0)
    Planilha2.Cells(3, 90) = IIf(Sheet6.Cells(40, 21) <> "", Sheet6.Cells(40, 21), 0)
    Planilha2.Cells(4, 90) = IIf(Sheet6.Cells(40, 23) <> "", Sheet6.Cells(40, 23), 0)
    Planilha2.Cells(5, 90) = IIf(Sheet6.Cells(40, 25) <> "", Sheet6.Cells(40, 25), 0)
    Planilha2.Cells(6, 90) = IIf(Sheet6.Cells(40, 27) <> "", Sheet6.Cells(40, 27), 0)
    Planilha2.Cells(7, 90) = IIf(Sheet6.Cells(40, 29) <> "", Sheet6.Cells(40, 29), 0)
    Planilha2.Cells(8, 90) = IIf(Sheet6.Cells(40, 31) <> "", Sheet6.Cells(40, 31), 0)
    
    'Bco_Dpge_II
    Planilha2.Cells(2, 91) = IIf(Sheet6.Cells(41, 19) <> "", Sheet6.Cells(41, 19), 0)
    Planilha2.Cells(3, 91) = IIf(Sheet6.Cells(41, 21) <> "", Sheet6.Cells(41, 21), 0)
    Planilha2.Cells(4, 91) = IIf(Sheet6.Cells(41, 23) <> "", Sheet6.Cells(41, 23), 0)
    Planilha2.Cells(5, 91) = IIf(Sheet6.Cells(41, 25) <> "", Sheet6.Cells(41, 25), 0)
    Planilha2.Cells(6, 91) = IIf(Sheet6.Cells(41, 27) <> "", Sheet6.Cells(41, 27), 0)
    Planilha2.Cells(7, 91) = IIf(Sheet6.Cells(41, 29) <> "", Sheet6.Cells(41, 29), 0)
    Planilha2.Cells(8, 91) = IIf(Sheet6.Cells(41, 31) <> "", Sheet6.Cells(41, 31), 0)
    
    'Bco_Avais_Fiancas_Prestdos
    Planilha2.Cells(2, 92) = IIf(Sheet6.Cells(42, 19) <> "", Sheet6.Cells(42, 19), 0)
    Planilha2.Cells(3, 92) = IIf(Sheet6.Cells(42, 21) <> "", Sheet6.Cells(42, 21), 0)
    Planilha2.Cells(4, 92) = IIf(Sheet6.Cells(42, 23) <> "", Sheet6.Cells(42, 23), 0)
    Planilha2.Cells(5, 92) = IIf(Sheet6.Cells(42, 25) <> "", Sheet6.Cells(42, 25), 0)
    Planilha2.Cells(6, 92) = IIf(Sheet6.Cells(42, 27) <> "", Sheet6.Cells(42, 27), 0)
    Planilha2.Cells(7, 92) = IIf(Sheet6.Cells(42, 29) <> "", Sheet6.Cells(42, 29), 0)
    Planilha2.Cells(8, 92) = IIf(Sheet6.Cells(42, 31) <> "", Sheet6.Cells(42, 31), 0)
    
    'Bco_Ag
    Planilha2.Cells(2, 93) = IIf(Sheet6.Cells(43, 19) <> "", Sheet6.Cells(43, 19), 0)
    Planilha2.Cells(3, 93) = IIf(Sheet6.Cells(43, 21) <> "", Sheet6.Cells(43, 21), 0)
    Planilha2.Cells(4, 93) = IIf(Sheet6.Cells(43, 23) <> "", Sheet6.Cells(43, 23), 0)
    Planilha2.Cells(5, 93) = IIf(Sheet6.Cells(43, 25) <> "", Sheet6.Cells(43, 25), 0)
    Planilha2.Cells(6, 93) = IIf(Sheet6.Cells(43, 27) <> "", Sheet6.Cells(43, 27), 0)
    Planilha2.Cells(7, 93) = IIf(Sheet6.Cells(43, 29) <> "", Sheet6.Cells(43, 29), 0)
    Planilha2.Cells(8, 93) = IIf(Sheet6.Cells(43, 31) <> "", Sheet6.Cells(43, 31), 0)

    'Bco_Func
    Planilha2.Cells(2, 94) = IIf(Sheet6.Cells(44, 19) <> "", Sheet6.Cells(44, 19), 0)
    Planilha2.Cells(3, 94) = IIf(Sheet6.Cells(44, 21) <> "", Sheet6.Cells(44, 21), 0)
    Planilha2.Cells(4, 94) = IIf(Sheet6.Cells(44, 23) <> "", Sheet6.Cells(44, 23), 0)
    Planilha2.Cells(5, 94) = IIf(Sheet6.Cells(44, 25) <> "", Sheet6.Cells(44, 25), 0)
    Planilha2.Cells(6, 94) = IIf(Sheet6.Cells(44, 27) <> "", Sheet6.Cells(44, 27), 0)
    Planilha2.Cells(7, 94) = IIf(Sheet6.Cells(44, 29) <> "", Sheet6.Cells(44, 29), 0)
    Planilha2.Cells(8, 94) = IIf(Sheet6.Cells(44, 31) <> "", Sheet6.Cells(44, 31), 0)

    'Bco_Fnds_Admn
    Planilha2.Cells(2, 95) = IIf(Sheet6.Cells(45, 19) <> "", Sheet6.Cells(45, 19), 0)
    Planilha2.Cells(3, 95) = IIf(Sheet6.Cells(45, 21) <> "", Sheet6.Cells(45, 21), 0)
    Planilha2.Cells(4, 95) = IIf(Sheet6.Cells(45, 23) <> "", Sheet6.Cells(45, 23), 0)
    Planilha2.Cells(5, 95) = IIf(Sheet6.Cells(45, 25) <> "", Sheet6.Cells(45, 25), 0)
    Planilha2.Cells(6, 95) = IIf(Sheet6.Cells(45, 27) <> "", Sheet6.Cells(45, 27), 0)
    Planilha2.Cells(7, 95) = IIf(Sheet6.Cells(45, 29) <> "", Sheet6.Cells(45, 29), 0)
    Planilha2.Cells(8, 95) = IIf(Sheet6.Cells(45, 31) <> "", Sheet6.Cells(45, 31), 0)
    
    'Bco_Cred_Trib
    Planilha2.Cells(2, 96) = IIf(Sheet6.Cells(46, 19) <> "", Sheet6.Cells(46, 19), 0)
    Planilha2.Cells(3, 96) = IIf(Sheet6.Cells(46, 21) <> "", Sheet6.Cells(46, 21), 0)
    Planilha2.Cells(4, 96) = IIf(Sheet6.Cells(46, 23) <> "", Sheet6.Cells(46, 23), 0)
    Planilha2.Cells(5, 96) = IIf(Sheet6.Cells(46, 25) <> "", Sheet6.Cells(46, 25), 0)
    Planilha2.Cells(6, 96) = IIf(Sheet6.Cells(46, 27) <> "", Sheet6.Cells(46, 27), 0)
    Planilha2.Cells(7, 96) = IIf(Sheet6.Cells(46, 29) <> "", Sheet6.Cells(46, 29), 0)
    Planilha2.Cells(8, 96) = IIf(Sheet6.Cells(46, 31) <> "", Sheet6.Cells(46, 31), 0)
    
    'Bco_Cdi_liqdz_Dia
    Planilha2.Cells(2, 97) = IIf(Sheet6.Cells(47, 19) <> "", Sheet6.Cells(47, 19), 0)
    Planilha2.Cells(3, 97) = IIf(Sheet6.Cells(47, 21) <> "", Sheet6.Cells(47, 21), 0)
    Planilha2.Cells(4, 97) = IIf(Sheet6.Cells(47, 23) <> "", Sheet6.Cells(47, 23), 0)
    Planilha2.Cells(5, 97) = IIf(Sheet6.Cells(47, 25) <> "", Sheet6.Cells(47, 25), 0)
    Planilha2.Cells(6, 97) = IIf(Sheet6.Cells(47, 27) <> "", Sheet6.Cells(47, 27), 0)
    Planilha2.Cells(7, 97) = IIf(Sheet6.Cells(47, 29) <> "", Sheet6.Cells(47, 29), 0)
    Planilha2.Cells(8, 97) = IIf(Sheet6.Cells(47, 31) <> "", Sheet6.Cells(47, 31), 0)
    
    'BCO_CAPT_MERC_ABER_NEG
    Planilha2.Cells(2, 98) = IIf(Sheet6.Cells(7, 36) <> "", Sheet6.Cells(7, 36), 0)
    Planilha2.Cells(3, 98) = IIf(Sheet6.Cells(7, 38) <> "", Sheet6.Cells(7, 38), 0)
    Planilha2.Cells(4, 98) = IIf(Sheet6.Cells(7, 40) <> "", Sheet6.Cells(7, 40), 0)
    Planilha2.Cells(5, 98) = IIf(Sheet6.Cells(7, 42) <> "", Sheet6.Cells(7, 42), 0)
    Planilha2.Cells(6, 98) = IIf(Sheet6.Cells(7, 44) <> "", Sheet6.Cells(7, 44), 0)
    Planilha2.Cells(7, 98) = IIf(Sheet6.Cells(7, 46) <> "", Sheet6.Cells(7, 46), 0)
    Planilha2.Cells(8, 98) = IIf(Sheet6.Cells(7, 48) <> "", Sheet6.Cells(7, 48), 0)
     
    'Bco_Div_Subord
    Planilha2.Cells(2, 99) = IIf(Sheet6.Cells(8, 36) <> "", Sheet6.Cells(8, 36), 0)
    Planilha2.Cells(3, 99) = IIf(Sheet6.Cells(8, 38) <> "", Sheet6.Cells(8, 38), 0)
    Planilha2.Cells(4, 99) = IIf(Sheet6.Cells(8, 40) <> "", Sheet6.Cells(8, 40), 0)
    Planilha2.Cells(5, 99) = IIf(Sheet6.Cells(8, 42) <> "", Sheet6.Cells(8, 42), 0)
    Planilha2.Cells(6, 99) = IIf(Sheet6.Cells(8, 44) <> "", Sheet6.Cells(8, 44), 0)
    Planilha2.Cells(7, 99) = IIf(Sheet6.Cells(8, 46) <> "", Sheet6.Cells(8, 46), 0)
    Planilha2.Cells(8, 99) = IIf(Sheet6.Cells(8, 48) <> "", Sheet6.Cells(8, 48), 0)
    
    'BCO_INSTRM_FIN_DERIV_PASS_NEG
    Planilha2.Cells(2, 100) = IIf(Sheet6.Cells(9, 36) <> "", Sheet6.Cells(9, 36), 0)
    Planilha2.Cells(3, 100) = IIf(Sheet6.Cells(9, 38) <> "", Sheet6.Cells(9, 38), 0)
    Planilha2.Cells(4, 100) = IIf(Sheet6.Cells(9, 40) <> "", Sheet6.Cells(9, 40), 0)
    Planilha2.Cells(5, 100) = IIf(Sheet6.Cells(9, 42) <> "", Sheet6.Cells(9, 42), 0)
    Planilha2.Cells(6, 100) = IIf(Sheet6.Cells(9, 44) <> "", Sheet6.Cells(9, 44), 0)
    Planilha2.Cells(7, 100) = IIf(Sheet6.Cells(9, 46) <> "", Sheet6.Cells(9, 46), 0)
    Planilha2.Cells(8, 100) = IIf(Sheet6.Cells(9, 48) <> "", Sheet6.Cells(9, 48), 0)
    
    'Bco_Depos_Aprazo
    Planilha2.Cells(2, 101) = IIf(Sheet6.Cells(10, 36) <> "", Sheet6.Cells(10, 36), 0)
    Planilha2.Cells(3, 101) = IIf(Sheet6.Cells(10, 38) <> "", Sheet6.Cells(10, 38), 0)
    Planilha2.Cells(4, 101) = IIf(Sheet6.Cells(10, 40) <> "", Sheet6.Cells(10, 40), 0)
    Planilha2.Cells(5, 101) = IIf(Sheet6.Cells(10, 42) <> "", Sheet6.Cells(10, 42), 0)
    Planilha2.Cells(6, 101) = IIf(Sheet6.Cells(10, 44) <> "", Sheet6.Cells(10, 44), 0)
    Planilha2.Cells(7, 101) = IIf(Sheet6.Cells(10, 46) <> "", Sheet6.Cells(10, 46), 0)
    Planilha2.Cells(8, 101) = IIf(Sheet6.Cells(10, 48) <> "", Sheet6.Cells(10, 48), 0)

    'Bco_Depos_Interfin
    Planilha2.Cells(2, 102) = IIf(Sheet6.Cells(11, 36) <> "", Sheet6.Cells(11, 36), 0)
    Planilha2.Cells(3, 102) = IIf(Sheet6.Cells(11, 38) <> "", Sheet6.Cells(11, 38), 0)
    Planilha2.Cells(4, 102) = IIf(Sheet6.Cells(11, 40) <> "", Sheet6.Cells(11, 40), 0)
    Planilha2.Cells(5, 102) = IIf(Sheet6.Cells(11, 42) <> "", Sheet6.Cells(11, 42), 0)
    Planilha2.Cells(6, 102) = IIf(Sheet6.Cells(11, 44) <> "", Sheet6.Cells(11, 44), 0)
    Planilha2.Cells(7, 102) = IIf(Sheet6.Cells(11, 46) <> "", Sheet6.Cells(11, 46), 0)
    Planilha2.Cells(8, 102) = IIf(Sheet6.Cells(11, 48) <> "", Sheet6.Cells(11, 48), 0)

End Sub

