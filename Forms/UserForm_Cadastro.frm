VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Cadastro 
   Caption         =   "Tela de Cadastro"
   ClientHeight    =   10932
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   18132
   OleObjectBlob   =   "UserForm_Cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Botao_limpa_Lista_grupo_selecionado_Click()
    Me.Lista_grupo_selecionado.Clear
End Sub

Private Sub Botao_LimpaList_Clientes_Consolidacao_Click()
    Me.List_Clientes_Consolidacao.Clear
End Sub

Private Sub Botao_Limpar_Click()

                
        'Limpa camppo tipo busca aba consolidado
                 
        Me.Tipo_Busca_Consolidado.Clear
                 
        Me.Tipo_Busca_Consolidado.AddItem "CNPJ/CPF"
        Me.Tipo_Busca_Consolidado.AddItem "CRC CLIENTE"
        Me.Tipo_Busca_Consolidado.AddItem "CRC GRUPO"
        Me.Tipo_Busca_Consolidado.AddItem "NOME"
                
        '---------------------------------------------------------------
        
        

End Sub

Private Sub Botao_pesquisa_cliente_Consolidado_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
   
    Set conn = getConnection()
       
    Dim i As Integer
    Dim j As Integer
    Dim l As Integer

    variable = ""
    variable2 = Me.Texto_Pesquisa_Consolidado.Value
    habilitaConsulta = False
    

    
   
   
    Select Case Tipo_Busca_Consolidado.ListIndex
    Case Is = 0
        variable = "CNPJ"
    Case Is = 1
        variable = "CD_CLI"
    Case Is = 2
        variable = "CD_GRP"
    Case Is = 3
        variable = "NM_EMP"
    End Select
    
    
    
     If Tipo_Busca_Consolidado.ListIndex = 1 Then
         If verificaNumeros(Me.Texto_Pesquisa_Consolidado.Value) Then
             MsgBox ("Não é permitido letras para esse tipo de pesquisa")
             Exit Sub
        End If
    End If
   
    Set rs = New ADODB.Recordset
   
    If variable = "" Then
        habilitaConsulta = False
        MsgBox "Favor selecionar o tipo de busca"
    Else
        habilitaConsulta = True
    End If
   
    If variable2 = "" Then
        habilitaConsulta = False
        MsgBox "Favor preecher dados para busca"
    Else
        habilitaConsulta = True
    End If
   
    If habilitaConsulta = True Then
           
        If Tipo_Busca_Consolidado.ListIndex = 3 Then
       
            qry = "select * from LB_PLANI.dim_grp_cli where " & variable & " like  '%" & UCase(variable2) & "%'"
                      
        ElseIf Tipo_Busca_Consolidado.ListIndex = 1 Then
        
            qry = "select * from LB_PLANI.dim_grp_cli where " & variable & " = " & variable2
        
        Else
           
            qry = "select * from LB_PLANI.dim_grp_cli where " & variable & " = '" & variable2 & "'"
       
        End If
       
        conn.Open
       
        rs.Open qry, conn, adOpenStatic
               
        If rs.RecordCount > 0 Then
          
            With Me.Lista_result_pesquisa_cliente_Consolidado
                   .Clear
               
                Do
                    .ColumnCount = 4
                    .ColumnWidths = "330"
                    .AddItem
                    .List(j, 0) = rs![cd_cli] & " - " & rs![Cd_grp] & " - " & rs![Flg_grp] & " - " & rs![nm_emp]
                                       
                    j = j + 1
                   
                    rs.MoveNext
                   
                Loop Until rs.EOF
            End With
           
        Else
       
            Me.Lista_result_pesquisa_cliente_Consolidado.Clear
            MsgBox "Nenhum resultado encontrado para essa pesquisa"
           
        End If
        
        rs.Close
        conn.Close
        
   End If
  
   
  
  
UserForm_Initialize_Exit:

    On Error Resume Next
   
    rs.Close
    Set rs = Nothing
    obConnection2.Close
    Set obConnection2 = Nothing
   
    Exit Sub
   
UserForm_Initialize_Err:

    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error!"
    Resume UserForm_Initialize_Exit
           
    rs.Close
    Set rs = Nothing
    obConnection.Close
    Set obConnection = Nothing


End Sub

Private Sub botao_sair_Click()
 ' fecha o useform de castro
  Unload Me
End Sub

Private Sub BT_Down_Cadastro_Click()

' On Error GoTo Erro

Dim Temp As String

For Item = ListBox4.ListCount - 2 To 0 Step -1

If ListBox4.Selected(Item) = True Then

With ListBox4

For x = 0 To (.ColumnCount - 1)

Temp = ListBox4.List(Item, x)
.List(Item, x) = .List(Item + 1, x)
.List(Item + 1, x) = Temp
.Selected(Item) = False
.Selected(Item + 1) = True

Next
End With

End If
Next

Exit Sub

'Erro:
'MsgBox "Teste, Ultima linha!!!!"



End Sub

Private Sub BT_PESQUISA_CADASTRO_P_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cd_cli As String
    Dim arr1(500)
    Dim count As Integer
   
       
    count = 0
   
    For iCtr = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(iCtr) = True Then
            arr1(count) = Me.ListBox2.List(iCtr)
            count = count + 1
        End If
    Next iCtr
       
    For iCtr = 0 To count - 1
    
        ax = Split(arr1(iCtr))
           
        If iCtr = 0 Then
           
            cd_cli = ax(0)
           
        Else
       
            cd_cli = cd_cli & "," & ax(0)
           
        End If
       
    Next iCtr
   
    Set conn = getConnection()
    Set rs = New ADODB.Recordset
        
    
    
    If cd_cli = "" Then
        
       MsgBox "Selecione um cliente para pesquisa"
       Exit Sub
        
    End If

    qry = "select distinct MAX(B.CD_GRP) AS CD_GRP, A.DT_EXERC, A.CD_CLI, B.nm_emp from LB_PLANI.FATO_BALANCO A join LB_PLANI.dim_grp_cli B on B.cd_cli = A.cd_cli where A.CD_CLI IN(" & cd_cli & ")"
   
    conn.Open
           
    rs.Open qry, conn, adOpenStatic
                   
            If rs.RecordCount > 0 Then
              
                With Me.ListBox_Cadastro_Periodo
                       .Clear
                   
                    Do
                        .ColumnCount = 1
                        .ColumnWidths = "60"
                        .AddItem
                        .List(j, 0) = rs![Dt_exerc] & " - " & rs![cd_cli] & " - " & rs![nm_emp]
                       
                        j = j + 1
                       
                        rs.MoveNext
                       
                    Loop Until rs.EOF
                End With
            Else
           
                Me.ListBox_Cadastro_Periodo.Clear
                MsgBox "Nenhum resultado encontrado para essa pesquisa"
               
            End If
           
    rs.Close
    conn.Close
      

End Sub

Private Sub BT_UP_Cadastro_Click()


'On Error GoTo Erro

Dim Temp As String

For Item = 1 To ListBox4.ListCount - 1

If ListBox4.Selected(Item) = True Then

With ListBox4

For x = 0 To (.ColumnCount - 1)

Temp = ListBox4.List(Item, x)
.List(Item, x) = .List(Item - 1, x)
.List(Item - 1, x) = Temp
.Selected(Item) = False
.Selected(Item - 1) = True

Next
End With

End If
Next

Exit Sub

'Erro:
'MsgBox "Teste, Primeira linha!!!!"



End Sub

Private Sub ComboBox1_Change()



End Sub

Private Sub BT_Clear_Cadastro_Click()

    'Botao de limpeza, limpar todos os campos
    Dim objeto As Control
   
    For Each objeto In Me.Controls 'faz o looping percorrendo todos os objetos do Userform
        If TypeName(objeto) = "TextBox" Or TypeName(objeto) = "ComboBox" Then  ' se o tipo do objeto encontrado tiver o nome TEXTBOX
               
               '----------- PROSPECT -----------
                 Me.TextBox_Nome_PROS = ""         'limpa o campo
                 Me.TextBox_CPF_CNPJ_PROS = "" 'limpa o campo
                'Me.ComboBox_Grupo_PROS.Clear
                'ComboBox_Grupo_PROS.AddItem "GRUPO"
                '----------- COMBINADO -----------
               
                 Me.TextBox_Nome_COMB = ""         'limpa o campo
                 Me.ListBox_Cadastro_Periodo.Clear
                 Me.ListBox1.Clear
                 Me.ListBox2.Clear
                 Me.ListBox3.Clear
                 Me.ListBox4.Clear
                 Me.Cbx_Cadastro_comb.Clear
                 Me.TxtBox_Inserir_Cadastro_Comb = "" 'limpa o campo
                 Me.ComboBox1 = "" 'Limpa o Campo
                 ComboBox1.ListIndex = 0
                 Me.ListBox_Cadastro_Periodo.Clear
                 'ComboBox_Grupo_COMB.AddItem "GRUPO"
                 Me.Cbx_Cadastro_comb.AddItem "CNPJ/CPF"
                 Me.Cbx_Cadastro_comb.AddItem "CRC CLIENTE"
                 Me.Cbx_Cadastro_comb.AddItem "CRC GRUPO"
                 Me.Cbx_Cadastro_comb.AddItem "NOME"
                 
                 ComboBox2.AddItem "CNPJ/CPF"
                 ComboBox2.AddItem "CRC CLIENTE"
                 ComboBox2.AddItem "CRC GRUPO"
                 ComboBox2.AddItem "NOME"
         End If
    Next objeto


End Sub

Private Sub BT_Clear_Cadastro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

'BT_Clear_Cadastro.BackColor = &HC0C0C0
'BT_Clear.MousePointer = fmMousePointerHourGlass
'BT_Clear.ForeColor = &HFFFFFF

End Sub

Private Sub BT_SALVAR_Click()

End Sub

Private Sub BT_PESQUISAR_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
   
    Set conn = getConnection()
       
    Dim i As Integer
    Dim j As Integer
    Dim l As Integer

    variable = ""
    variable2 = Me.TxtBox_Inserir_Cadastro_Comb
    habilitaConsulta = False
   
   
    Select Case Cbx_Cadastro_comb.ListIndex
    Case Is = 0
        variable = "CNPJ"
    Case Is = 1
        variable = "CD_CLI"
    Case Is = 2
        variable = "CD_GRP"
    Case Is = 3
        variable = "NM_EMP"
    End Select
      
    If Cbx_Cadastro_comb.ListIndex = 1 Then
       If verificaNumeros(Me.TxtBox_Inserir_Cadastro_Comb) Then
            MsgBox ("Não é permitido letras para esse tipo de pesquisa")
            Exit Sub
       End If
    End If
   
    Set rs = New ADODB.Recordset
   
    If variable = "" Then
        habilitaConsulta = False
        MsgBox "Favor selecionar o tipo de busca"
    Else
        habilitaConsulta = True
    End If
   
    If variable2 = "" Then
        habilitaConsulta = False
        MsgBox "Favor preecher dados para busca"
    Else
        habilitaConsulta = True
    End If
   
   
    If habilitaConsulta = True Then
           
        If Cbx_Cadastro_comb.ListIndex = 3 Then
       
            qry = "select * from LB_PLANI.dim_grp_cli where " & variable & " like  '%" & UCase(variable2) & "%'"
                      
        ElseIf Cbx_Cadastro_comb.ListIndex = 1 Then
        
            qry = "select * from LB_PLANI.dim_grp_cli where " & variable & " = " & variable2
        
        Else
           
            qry = "select * from LB_PLANI.dim_grp_cli where " & variable & " = '" & variable2 & "'"
       
        End If
       
        conn.Open
       
        rs.Open qry, conn, adOpenStatic
               
        If rs.RecordCount > 0 Then
          
            With Me.ListBox1
                   .Clear
               
                Do
                    .ColumnCount = 4
                    .ColumnWidths = "330"
                    .AddItem
                    .List(j, 0) = rs![cd_cli] & " - " & rs![Cd_grp] & " - " & rs![Flg_grp] & " - " & rs![nm_emp]
                                       
                    j = j + 1
                   
                    rs.MoveNext
                   
                Loop Until rs.EOF
            End With
           
        Else
       
            Me.ListBox1.Clear
            MsgBox "Nenhum resultado encontrado para essa pesquisa"
           
        End If
        
        rs.Close
        conn.Close
        
   End If
  
  
UserForm_Initialize_Exit:

    On Error Resume Next
   
    rs.Close
    Set rs = Nothing
    obConnection2.Close
    Set obConnection2 = Nothing
   
    Exit Sub
   
UserForm_Initialize_Err:

    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error!"
    Resume UserForm_Initialize_Exit
           
    rs.Close
    Set rs = Nothing
    obConnection.Close
    Set obConnection = Nothing

End Sub

Private Sub BT_SAIR_CADASTRO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

'BT_SAIR_CADASTRO.BackColor = &HC0C0C0
'BT_Planilhar.MousePointer = fmMousePointerHourGlass
'BT_Planilhar.ForeColor = &HFFFFFF

End Sub

Private Sub BT_SAIR_CADASTRO_Click()

' fecha o useform de castro
Unload Me

End Sub

Private Sub BT_SALVAR_CADASTRO_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cd_cli As String
    Dim cod_cli As String
    Dim nome_combinado As String
    Dim count As Integer
    Dim cod_grupo As String
    Dim cotacao As Double
    Dim mes_fech As String
    Dim moeda As String
    Dim dt_crg As String
       
        
    dt_crg = montaData(Now())
       
    nome_combinado = TextBox_Nome_COMB.Value
    
    mes_fech = Me.mes.Value
        
    If mes_fech = "" Then
    
        MsgBox ("Mês fechamento não preenchido")
        Exit Sub
        
    End If
    
    If CInt(mes_fech) = 0 Or CInt(mes_fech) > 12 Then
    
        MsgBox ("Mês de fechameto não existe")
        Exit Sub
    End If
    
    If nome_combinado = "" Then
        
        MsgBox ("Nome combimado não preenchido")
        Exit Sub
        
    End If
        
    For iCtr = 0 To Me.ListBox3.ListCount - 1
        If Me.ListBox3.Selected(iCtr) = True Then
            cd_cli = Me.ListBox3.List(iCtr)
        End If
    Next iCtr
    
    ax = Split(cd_cli)
    
    If Len(Join(ax)) = 0 Then
    
        MsgBox ("Chave de grupo não selecionado")
        Exit Sub
        
    End If
    
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
   
    If Me.ListBox4.ListCount = 0 Then
    
        MsgBox ("Nenhum planilhamento foi selecionado")
        Exit Sub
        
    End If
    
    
    If Me.ListBox2.ListCount > Me.ListBox4.ListCount Then
    
        MsgBox ("Numero de periodos inferior ao numeros de clientes combinados")
        Exit Sub
    
    End If
       
    
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    

  clis = ""
   
   For iCtr = 0 To Me.ListBox4.ListCount - 1
      If Me.ListBox4.Selected(iCtr) = True Then
      
        If iCtr = 0 Then
         clis = Split(Me.ListBox4.List(iCtr))(2)
         
        Else
        
          clis = clis & "," & Split(Me.ListBox4.List(iCtr))(2)
        
        End If
         
      End If
   Next iCtr
     
    Set conn = getConnection()
    
    'VALIDACAO MESMO LAYOUT
    
    qryValidacao = "select  COUNT(*) , layout from lb_plani.dim_grp_cli where cd_cli in (" & clis & ") group by layout"
    
    
    Set rs = New ADODB.Recordset
    conn.Open
    
    rs.Open qryValidacao, conn, adOpenStatic
    
    If rs.RecordCount > 1 Then
    
        MsgBox ("Cliente no pertencem ao mesmo layout")
        Exit Sub
        
    End If
    
    rs.Close
    conn.Close
    
    
    'CRIA E SALVA CLIENTE COMBINADO
    
    Set rs = New ADODB.Recordset
    
    cod_grupo = ax(2)
       
    qry = "select max(cd_cli)AS cd_cli from LB_PLANI.dim_grp_cli where cd_grp = '" & cod_grupo & "'"
    
    conn.Open
    
    rs.Open qry, conn, adOpenStatic
    
    cod_cli = Str(rs![cd_cli])
    
    cod_cli = geraCodigo2(Str(cod_cli), Str(cod_grupo))
    
    rs.Close
    conn.Close
    
    conn.Open
        
        qry2 = "inset into LB_PLANI.dim_grp_cli (cd_cli,cd_grp,nm_emp) values(" & cod_cli & ",'" & cod_grupo & "', '" & nome_combinado & "')"
                  
        conn.Execute (qry2)
    
    conn.Close
    
    MsgBox ("Cliente " & nome_combinado & " salvo com sucesso! Código: " & cod_cli)
    
   'FIM CRIA E SALVA CLIENTE COMBINADO
    
   '--------------------------------------------------------------------------------------
   'SALVA COMPOSIÇÃO CLIENTE COMBINADO NA BASE LOG_COMBINADO
    
    conn.Open
        
    For iCtr = 0 To Me.ListBox2.ListCount - 1
    
        cd_cli_transacao = Trim(Split(Me.ListBox2.List(iCtr))(0))
   
        qryLogCombinado = "ISERT INTO LB_PLANI.LOG_COMBINADO (CD_CLI_COMB,CD_CLI,DT_CRG) VALUES (" & cod_cli & ", " & cd_cli_transacao & ", '" & dt_crg & "'dt)"
           
        conn.Execute (qryLogCombinado)
        
    Next iCtr
    
    conn.Close
    
   'FIM SALVA COMPOSIÇÃO CLIENTE COMBINADO NA BASE LOG_COMBINADO
       
   
   '--------------------------------------------------------------------------------------
   'SALVA PLANILHAMENTO SELECIONADO NA TBELA FATO_BALANCO_AUX
   
   For iCtr = 0 To Me.ListBox4.ListCount - 1
      If Me.ListBox4.Selected(iCtr) = True Then
            count = count + 1
      End If
   Next iCtr
   
   
   ReDim arrPeriodos(count)
   
   
   For iCtr = 0 To Me.ListBox4.ListCount - 1
      If Me.ListBox4.Selected(iCtr) = True Then
      
         arrPeriodos(iCtr) = Me.ListBox4.List(iCtr)
      End If
   Next iCtr
   
   cd_cli_2 = ""
   
   For iCtr = 0 To count - 1
   
     aux2 = arrPeriodos(iCtr)
     aux3 = Split(aux2)
                  
     per = "'" & aux3(0) & "'"
     cd_cli_2 = aux3(2)
     
     qry3 = "insert into lb_plani.fato_balanco_aux select * from "
     qry3 = qry3 & "lb_plani.fato_balanco where cd_cli in (" & cd_cli_2 & ") "
     qry3 = qry3 & "and dt_exerc in (" & per & ") and dt_crg = (select max(dt_crg) from "
     qry3 = qry3 & "lb_plani.fato_balanco where cd_cli in (" & cd_cli_2 & ") and dt_exerc in (" & per & "))"
         
     conn.Open
     conn.Execute (qry3)
     conn.Close
     
     moeda = Trim(Split(Me.ComboBox1.Value)(0))
          
     qryMoeda = "select moeda from lb_plani.fato_balanco_aux where moeda <> '" & moeda & "' and dt_exerc = " & per & " and cd_cli = " & cd_cli_2 & ""
          
     Set rs = New ADODB.Recordset
     
     conn.Open
     rs.Open qryMoeda, conn, adOpenStatic
     
     If rs.RecordCount > 0 Then
     
        Do
        
           moedaBalanco = Trim(rs![moeda])
           
           rs.MoveNext
           
        Loop Until rs.EOF
        
             
        Select Case moeda
        
        Case Is = "BRL"
            variable = "VL_PTAX_COMPRA_BRL"
            
        Case Is = "USD"
            variable = "VL_PTAX_COMPRA_USD"
            
        Case Is = "EUR"
            variable = "VL_PTAX_COMPRA_EUR"
            
        End Select
     
        'rs.Close
        'conn.Close
        
        qryGetCotacao = "select " & variable & " AS cotacao from lb_plani.dim_moeda where moeda_orig = '" & moedaBalanco & "' and DT_COTACAO = " & per
        
        Set rs = New ADODB.Recordset
        'conn.Open
           
        rs.Open qryGetCotacao, conn, adOpenStatic
           
        If rs.RecordCount > 0 Then
           
           Do
           
             cotacao = rs![cotacao]
             rs.MoveNext
               
           Loop Until rs.EOF
           
        Else
           MsgBox ("Período selecionado não possue cotação correspondente")
        End If
           
        'rs.Close
        'conn.Close
            
        'conn.Open
               
           qryUpdateMoeda = updateMoeda(cotacao, Replace(per, "'", ""), Str(cd_cli_2))
            
           conn.Execute (qryUpdateMoeda)
            
         'conn.Close
      End If
      
      rs.Close
      conn.Close
          
   Next iCtr
     
   For iCtr = 0 To count - 1
   
       aux2 = arrPeriodos(iCtr)
       aux3 = Split(aux2)
      
        If iCtr = 0 Then
            
            per = "'" & aux3(0) & "'"
            cd_cli_2 = aux3(2)
            
        Else
        
            per = per & ", '" & aux3(0) & "'"
            cd_cli_2 = cd_cli_2 & "," & aux3(2)
           
        End If
        
   Next iCtr
 
   
   
   'FIM SALVA PLANILHAMENTO SELECIONADO NA TBELA FATO_BALANCO_AUX
   '--------------------------------------------------------------------------------------
     
   'MONTAR ARRAY COM AS DATAS SELECIONADAS PARA COMBIAR BALANCOS
      
   per_aux = Split(per, ",")
   

   For i = LBound(per_aux) To UBound(per_aux)
      
       perComb = Replace(per_aux(i), "'", "")
       
       a = Right(perComb, Len(Trim(perComb)) - 6)
       
       If i = 0 Then
           
           per_trat_conct = a
           per_aux_log = perComb
            
       Else
       
           per_trat_conct = per_trat_conct & "," & a
           per_aux_log = per_aux_log & "," & perComb
       
       End If
       
       Next i
    
        arr = Split(per_trat_conct, ",")
       
        arr_aux = arr
        
        per_log = Split(per_aux_log, ",")
    
        For i = LBound(arr) To UBound(arr)
        
        If i = 0 Then
        
            a = arr(i)
            fff = a
            dt_exerc_log = per_log(i)
        Else
            a = arr(i)
            
            If Not fff Like "*" & Trim(arr(i)) & "*" Then
                    
                    fff = fff & "," & a
                    dt_exerc_log = dt_exerc_log & "," & per_log(i)
             End If
            
        End If
    
    Next i
    
    arr3 = Split(fff, ",")
   

   'FIM MONTAR ARRAY COM AS DATAS SELECIONADAS PARA COMBIAR BALANCOS
   '--------------------------------------------------------------------------------------
   'LOOPING PARA COMBINAR E SALVAR BALANCOS AGRUPADOS POR DATAS

      
    For i = LBound(arr3) To UBound(arr3)
            
        mes_fech = Me.mes.Value
        
        qry4 = somaBalanco(Trim(Str(arr3(i))), cod_cli, cod_grupo, mes_fech, moeda, dt_crg)
        
        qryLogTransacao = "INSERT INTO LB_PLANI.LOG_TRANSACAO (DT_CRG,DATA_EXERC,LOGIN_USER,CD_CLI,ACAO,MAQUINA,IP)"
        qryLogTransacao = qryLogTransacao & " VALUES ('" & dt_crg & "'dt, '" & Split(dt_exerc_log)(i) & "', '"
        qryLogTransacao = qryLogTransacao & getLoginWindows() & "' , " & cod_cli & ", 'COBINAÇÃO', '"
        qryLogTransacao = qryLogTransacao & getMachineName() & "' , '" & GetIPAddress() & "')"
        
        conn.Open
            conn.Execute (qry4)
            conn.Execute (qryLogTransacao)
        conn.Close
                
    Next i
    
   ' --------------------------------------------------------------------------------------
    conn.Open
     
      qry5 = "delete from lb_plani.fato_balanco_aux where dt_exerc in (" & per & ") and cd_cli in (" & cd_cli_2 & ")"
      conn.Execute (qry5)
            
      conn.Close
   
     MsgBox ("Clientes combinados com sucesso")
   
End Sub

Private Sub BT_SALVAR_CADASTRO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

'BT_SALVAR_CADASTRO.BackColor = &HC0C0C0
'BT_Planilhar.MousePointer = fmMousePointerHourGlass
'BT_Planilhar.ForeColor = &HFFFFFF

End Sub

Private Sub ComboBox_Grupo_PROS_Change()

End Sub


Private Sub BTN_MoveSelectedRight_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox2.AddItem Me.ListBox1.List(iCtr)
            Me.ListBox3.AddItem Me.ListBox1.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox1.ListCount - 1 To 0 Step -1
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox1.RemoveItem iCtr
        End If
    Next iCtr

End Sub


Private Sub BTN_MoveSelectedLeft_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox2.ListCount - 1 To 0 Step -1
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox2.RemoveItem iCtr
            Me.ListBox3.RemoveItem iCtr
        End If
    Next iCtr
   
End Sub

Private Sub Cbx_Cadastro_comb_Change()

End Sub


Private Sub CommandButton_Salvar_P_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim cod_cli As String
    
    
    nome = Me.TextBox_Nome_PROS
    Cnpj = Me.TextBox_CPF_CNPJ_PROS
    
    Set conn = getConnection()
    
    Set rs = New ADODB.Recordset
    
    If nome = "" Then
        MsgBox ("Campo Nome é obrigatório e não foi preenchido")
        Exit Sub
    End If
    
    If Me.ListBox_p.ListCount = 0 Then
        MsgBox ("Selecione um grupo para cadastro")
        Exit Sub
    End If
    
    grupoSelecao = Me.ListBox_p.List(0)
    
    grupoArray = Split(grupoSelecao)
    
    grupo = grupoArray(0)
    
    conn.Open
    
    Set rs = New ADODB.Recordset
    
    qry = "select max(cd_cli) as cd_cli from LB_PLANI.dim_grp_cli where cd_grp = '" & grupo & "'"
    
    Set rs = conn.Execute(qry)
    
        cod_cli = Str(rs![cd_cli])
        
        cod_cli = geraCodigo2(cod_cli, Str(grupo))
        
    rs.Close
    conn.Close
    
        dt_crg = montaData(Now())
        
        qryLogTransacao = "INSERT INTO LB_PLANI.LOG_TRANSACAO (DT_CRG,DATA_EXERC,LOGIN_USER,CD_CLI,ACAO,MAQUINA,IP)"
        qryLogTransacao = qryLogTransacao & " VALUES ('" & dt_crg & "'dt, '', '"
        qryLogTransacao = qryLogTransacao & getLoginWindows() & "' , " & cod_cli & ", 'PROSPECT', '"
        qryLogTransacao = qryLogTransacao & getMachineName() & "' , '" & GetIPAddress() & "')"
    
    conn.Open
    
       qry = "insert into LB_PLANI.dim_grp_cli (cd_cli,cd_grp,nm_emp,cnpj) values(" & cod_cli & " , '" & grupo & "' , '" & nome & "','" & Cnpj & "')"
           
       conn.Execute (qry)
       conn.Execute (qryLogTransacao)
           
    conn.Close
    
    Me.TextBox_Nome_PROS.Value = ""
    
    Me.TextBox_CPF_CNPJ_PROS.Value = ""
    
    MsgBox ("Cliente com CRC " & cod_cli & " salvo com sucesso!")

End Sub

Private Sub CommandButton15_Click()



    Dim iCtr As Integer
    Dim count As Integer
    Dim periodos As String
    Dim cd_crp As String
    Dim codigos As String
    Dim moeda As String
    Dim mesFechamento As String
    Dim sas As SASExcelAddIn
    Dim prompts As SASPrompts
    Dim stp As SASStoredProcess
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    
    
    moeda = Me.Moeda_Consolidado.Value
           
    mesFechamento = Me.mes_fechamento_consolidado.Value
        
    If mesFechamento = "" Then
    
        MsgBox ("Mês fechamento não preenchido")
        Exit Sub
        
    End If
    
    If CInt(mesFechamento) = 0 Or CInt(mesFechamento) > 12 Then
    
        MsgBox ("Mês de fechameto não existe")
        Exit Sub
        
    End If
    
        
    count = 0
    
    For iCtr = 0 To Me.Lista_Periodos.ListCount - 1
        
        If Me.Lista_Periodos.Selected(iCtr) = True Then
            count = count + 1
        End If
    Next iCtr
    

    If count = 0 Then
    
        MsgBox ("Nenhum período selecionado para consolidação")
        Exit Sub
        
    End If
    
    cd_crp = ""
    
    For iCtr = 0 To Me.Lista_grupo_selecionado.ListCount - 1
        
        If Me.Lista_grupo_selecionado.Selected(iCtr) = True Then
            cd_crp = Split(Me.Lista_grupo_selecionado.List(iCtr))(2)
        End If
    Next iCtr
    

    If cd_crp = "" Then
    
        MsgBox ("Nenhum grupo principal selecionado para consolidação")
        Exit Sub
    End If
    
    
    periodos = ""
    count = 0

    For iCtr = 0 To Me.Lista_Periodos.ListCount - 1
        
        If Me.Lista_Periodos.Selected(iCtr) = True Then
        
            aux = Split(Me.Lista_Periodos.List(iCtr))
                
            periodos = periodos & """" & aux(0) & """" & "#" & aux(2) & "#" & aux(4) & "#" & ":"
            
            If count = 0 Then
            
                codigos = aux(2)
            Else
            
                codigos = codigos & "," & aux(2)
                
            End If
            
            count = count + 1
        End If
               
    Next iCtr
    
    'Validação layouts diferentes -----------------------------------------------
            
    qryValidacao = "select COUNT(*) , layout from lb_plani.dim_grp_cli where cd_cli in (" & codigos & ") group by layout"
    
    
    Set conn = getConnection()
    Set rs = New ADODB.Recordset
    
    conn.Open
    
    rs.Open qryValidacao, conn, adOpenStatic
    
    If rs.RecordCount > 1 Then
    
        MsgBox ("Cliente no pertencem ao mesmo layout")
        Exit Sub
        
    End If
    
    rs.Close
    
    conn.Close
        
    login = getLoginWindows()
    maquina = getMachineName()
    ip = GetIPAddress()
    
    '-----------------------------------------------------------------------------
        
    Set sas = Application.COMAddIns.Item("SAS.ExcelAddIn").Object
    
    Set prompts = sas.CreateSASPromptsObject
       
    
    prompts.Add "balancosSelecionados", periodos
    prompts.Add "cd_crp", cd_crp
    prompts.Add "moeda", moeda
    prompts.Add "mesFechamento", mesFechamento
    prompts.Add "login", login
    prompts.Add "maquina", maquina
    prompts.Add "ip", ip
                
    Set stp = sas.InsertStoredProcess("/My Folder/combinaPlanilhamento", Range("A1"), prompts)
            
    Debug.Print (periodos)
    Debug.Print (cd_crp)
    Debug.Print (moeda)
    Debug.Print (mesFechamento)
    Debug.Print (getLoginWindows())
    Debug.Print (getMachineName())
    Debug.Print (GetIPAddress())
        
End Sub

Private Sub CommandButton16_Click()
    Me.ListBox_Cadastro_Periodo.Clear
End Sub

Private Sub CommandButton17_Click()
 Me.ListBox4.Clear
End Sub

Private Sub CommandButton18_Click()
    
    List_Clientes_Consolidacao

End Sub

Private Sub CommandButton2_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox_Cadastro_Periodo.ListCount - 1
        If Me.ListBox_Cadastro_Periodo.Selected(iCtr) = True Then
            Me.ListBox4.AddItem Me.ListBox_Cadastro_Periodo.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox_Cadastro_Periodo.ListCount - 1 To 0 Step -1
        If Me.ListBox_Cadastro_Periodo.Selected(iCtr) = True Then
            Me.ListBox_Cadastro_Periodo.RemoveItem iCtr
        End If
    Next iCtr
   
End Sub



Private Sub BTN_moveAllLeft_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox2.ListCount - 1
        Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
    Next iCtr

    Me.ListBox2.Clear
    Me.ListBox3.Clear
   
End Sub



Private Sub CommandButton3_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox4.ListCount - 1
        If Me.ListBox4.Selected(iCtr) = True Then
            Me.ListBox_Cadastro_Periodo.AddItem Me.ListBox4.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox4.ListCount - 1 To 0 Step -1
        If Me.ListBox4.Selected(iCtr) = True Then
            Me.ListBox4.RemoveItem iCtr
        End If
    Next iCtr
   
End Sub

Private Sub CommandButton7_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim pesq As String
    
    Dim j As Integer
    
    variable = ""
    variable2 = Me.TextBox2
    habilitaConsulta = False
   
   
    Select Case ComboBox2.ListIndex
    Case Is = 0
        variable = "CNPJ"
    Case Is = 1
        variable = "CD_CLI"
    Case Is = 2
        variable = "CD_GRP"
    Case Is = 3
        variable = "NM_EMP"
    End Select
         
   
    If variable = "" Then
        habilitaConsulta = False
        MsgBox "Favor selecionar o tipo de busca"
    Else
        habilitaConsulta = True
    End If
   
    If variable2 = "" Then
        habilitaConsulta = False
        MsgBox "Favor preecher dados para busca"
    Else
        habilitaConsulta = True
    End If
   
   
    If habilitaConsulta = True Then
    
        Set conn = getConnection()
        Set rs = New ADODB.Recordset
           
        If ComboBox2.ListIndex = 3 Then
       
            qry = "select cd_grp, nm_emp from LB_PLANI.dim_grp_cli where FLG_GRP = 'G' AND  CD_GRP IN select cd_grp from LB_PLANI.dim_grp_cli where " & variable & " like  '%" & UCase(variable2) & "%'"
                      
        ElseIf ComboBox2.ListIndex = 1 Then
        
            qry = "select cd_grp, nm_emp from LB_PLANI.dim_grp_cli where flg_grp = 'G' AND  CD_GRP = select cd_grp from LB_PLANI.dim_grp_cli where " & variable & " = " & variable2
            
        ElseIf ComboBox2.ListIndex = 2 Then
        
            qry = "select cd_grp, nm_emp from LB_PLANI.dim_grp_cli where flg_grp = 'G' AND  CD_GRP = " & variable & " = '" & variable2 & "'"
        
        Else
           
            qry = "select cd_grp, nm_emp from LB_PLANI.dim_grp_cli where flg_grp = 'G' AND  CD_GRP = select cd_grp from LB_PLANI.dim_grp_cli where " & variable & " = '" & variable2 & "'"
       
        End If
        
        conn.Open
       
        rs.Open qry, conn, adOpenStatic
               
        If rs.RecordCount > 0 Then
          
            With Me.ListBox_p
                   .Clear
               
                Do
                    .ColumnCount = 4
                    .ColumnWidths = "330"
                    .AddItem
                    .List(j, 0) = rs![Cd_grp] & " - " & rs![nm_emp]
                                       
                    j = j + 1
                   
                    rs.MoveNext
                   
                Loop Until rs.EOF
            End With
           
        Else
       
            Me.ListBox_p.Clear
            MsgBox "Nenhum resultado encontrado para essa pesquisa"
           
        End If
        
        rs.Close
        conn.Close
        
   End If
  
   
  
  
UserForm_Initialize_Exit:

    On Error Resume Next
   
    rs.Close
    Set rs = Nothing
    obConnection2.Close
    Set obConnection2 = Nothing
   
    Exit Sub
   
UserForm_Initialize_Err:

    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error!"
    Resume UserForm_Initialize_Exit
           
    rs.Close
    Set rs = Nothing
    obConnection.Close
    Set obConnection = Nothing


End Sub


Private Sub Limpa_Lista_Periodos_Click()
    Me.Lista_Periodos.Clear
End Sub

Private Sub limpa_selecao_clientes_Click()
    Me.ListBox2.Clear
    Me.ListBox3.Clear
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()


Dim Size As Integer
Size = Me.List0.ListCount - 1
ReDim ListBoxContents(0 To Size) As String
Dim i As Integer

For i = 0 To Size
    ListBoxContents(i) = Me.List0.ItemData(i)
Next i

For i = 0 To Size
    MsgBox ListBoxContents(i)
Next i


End Sub



Private Sub Pesquisar_Periodos_Click()
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim iCtr As Long
    
    Dim codigos As String
    Dim count As Integer
    Dim j As Integer
    
    count = 0
    codigos = ""
    
    
    For iCtr = 0 To Me.List_Clientes_Consolidacao.ListCount - 1
        
        If Me.List_Clientes_Consolidacao.Selected(iCtr) = True Then
                    
            aux = Split(Me.List_Clientes_Consolidacao.List(iCtr))(0)
            
            If count = 0 Then
                codigos = aux
            Else
                codigos = codigos & "," & aux
            End If
        
        End If
        
        count = count + 1
    Next iCtr
    
    If codigos = "" Then
    
        MsgBox ("Selecione um cliente(s) para pesquisa")
        Exit Sub
    End If
    
    
    qryPeriodos = "SELECT  distinct MAX(B.CD_GRP) AS CD_GRP ,B.CD_CLI, B.DT_EXERC, C.NM_EMP FROM LB_PLANI.FATO_BALANCO B JOIN LB_PLANI.DIM_GRP_CLI C ON  C.CD_CLI = B.CD_CLI WHERE B.CD_CLI IN (" & codigos & ")"
    
    
    Set conn = getConnection()
    Set rs = New ADODB.Recordset
    
    conn.Open
    
    rs.Open qryPeriodos, conn, adOpenStatic
    
    If rs.RecordCount > 0 Then
    
                  
        With Me.Lista_Periodos
           .Clear
                   
           Do
              .ColumnCount = 1
              .ColumnWidths = "60"
              .AddItem
              .List(j, 0) = rs![Dt_exerc] & " - " & rs![cd_cli] & " - " & rs![Cd_grp] & " - " & rs![nm_emp]
                       
               j = j + 1
                       
               rs.MoveNext
                       
           Loop Until rs.EOF
        End With
    
    Else
    
        Lista_Periodos.Clear
        MsgBox ("Não exitem resultados para sua pesquisa")
        
    End If
    
    rs.Close
    conn.Close
    
End Sub

Private Sub Remove_Clientes_Consolidacao_Click()
    
    Dim iCtr As Long

    For iCtr = Me.List_Clientes_Consolidacao.ListCount - 1 To 0 Step -1
        If Me.List_Clientes_Consolidacao.Selected(iCtr) = True Then
            Me.List_Clientes_Consolidacao.RemoveItem iCtr
        End If
    Next iCtr
    
End Sub

Private Sub Remove_Grupo_Consolidado_Click()

    Dim iCtr As Long

    For iCtr = Me.Lista_grupo_selecionado.ListCount - 1 To 0 Step -1
        If Me.Lista_grupo_selecionado.Selected(iCtr) = True Then
            Me.Lista_grupo_selecionado.RemoveItem iCtr
        End If
    Next iCtr
   
End Sub

Private Sub Seleciona_Clientes_Consolidacao_Click()
    
    Dim iCtr As Long
    Dim jCtr As Long
    Dim adiciona As Boolean

    For iCtr = 0 To Me.Lista_result_pesquisa_cliente_Consolidado.ListCount - 1
        
        If Me.Lista_result_pesquisa_cliente_Consolidado.Selected(iCtr) = True Then
          
          If Me.List_Clientes_Consolidacao.ListCount = 0 Then
          
            Me.List_Clientes_Consolidacao.AddItem Me.Lista_result_pesquisa_cliente_Consolidado.List(iCtr)
            
          Else
              For jCrt = 0 To Me.List_Clientes_Consolidacao.ListCount - 1
              
              
               
               
               'MsgBox (Trim(Split(Me.Lista_result_pesquisa_cliente_Consolidado.List(iCtr), "-")(3)) & " <---> " & Trim(Split(Me.List_Clientes_Consolidacao.List(jCrt), "-")(3)))
                
                If Not Trim(Split(Me.Lista_result_pesquisa_cliente_Consolidado.List(iCtr), "-")(3)) Like "*" & Trim(Split(Me.List_Clientes_Consolidacao.List(jCrt), "-")(3)) & "*" Then
            
                   adiciona = True
                    
                Else
                
                    adiciona = False
                    
                
                End If
                
                If adiciona = False Then Exit For
                
                
              Next jCrt
              
              If adiciona Then
                Me.List_Clientes_Consolidacao.AddItem Me.Lista_result_pesquisa_cliente_Consolidado.List(iCtr)
              End If
              
          End If
        End If
    Next iCtr
    
End Sub

Private Sub Seleciona_grupo_consolidado_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.Lista_result_pesquisa_cliente_Consolidado.ListCount - 1
        If Me.Lista_result_pesquisa_cliente_Consolidado.Selected(iCtr) = True Then
        
            If Lista_grupo_selecionado.ListCount > 0 Then
            
               MsgBox ("Somente um grupo é permitido para esta seleção")
               Exit Sub
            
            Else
            
                aa = Split(Me.Lista_result_pesquisa_cliente_Consolidado.List(iCtr))
                
                If Trim(aa(4)) = "G" Then
            
                    Me.Lista_grupo_selecionado.AddItem Me.Lista_result_pesquisa_cliente_Consolidado.List(iCtr)
                
                Else
                    
                    MsgBox ("Este clienet não é um grupo")
                    Exit Sub
                End If
            End If
        End If
        
    Next iCtr

End Sub



Private Sub UserForm_Activate()

    With Me
        .ScrollBars = fmScrollBarsBoth
        .ScrollHeight = .InsideHeight * 1.1
        .ScrollWidth = .InsideWidth * 1.1
    End With



End Sub



Private Sub UserForm_Initialize()

    Me.ListBox1.MultiSelect = fmMultiSelectMulti
    Me.ListBox2.MultiSelect = fmMultiSelectMulti
    Me.ListBox3.MultiSelect = fmMultiSelectSingle
    Me.Lista_result_pesquisa_cliente_Consolidado.MultiSelect = fmMultiSelectMulti 'Painel Consolidado
    Me.List_Clientes_Consolidacao.MultiSelect = fmMultiSelectMulti 'Painel Consolidado
    
   
    Cbx_Cadastro_comb.AddItem "CNPJ/CPF"
    Cbx_Cadastro_comb.AddItem "CRC CLIENTE"
    Cbx_Cadastro_comb.AddItem "CRC GRUPO"
    Cbx_Cadastro_comb.AddItem "NOME"
    
    ComboBox2.AddItem "CNPJ/CPF"
    ComboBox2.AddItem "CRC CLIENTE"
    ComboBox2.AddItem "CRC GRUPO"
    ComboBox2.AddItem "NOME"
    
    'Monta comboBox para seleção de moeda
    
    ComboBox1.AddItem "BRL - Real"
    ComboBox1.AddItem "USD - Dolar"
    ComboBox1.AddItem "EUR - Euro"
    
    ComboBox1.ListIndex = 0
    
    'Monta Combo Tipo Busca Pesquisa no Cadastro de Consolidados
   
    Me.Tipo_Busca_Consolidado.AddItem "CNPJ/CPF"
    Me.Tipo_Busca_Consolidado.AddItem "CRC CLIENTE"
    Me.Tipo_Busca_Consolidado.AddItem "CRC GRUPO"
    Me.Tipo_Busca_Consolidado.AddItem "NOME"
   
    'Monta comboBox para seleção de moeda no painel de consolidados
    
    Moeda_Consolidado.AddItem "BRL - Real"
    Moeda_Consolidado.AddItem "USD - Dolar"
    Moeda_Consolidado.AddItem "EUR - Euro"
    
    Moeda_Consolidado.ListIndex = 0

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    'BT_SAIR_CADASTRO.BackColor = &H8000000F
    'BT_SAIR_CADASTRO.ForeColor = &H80000012
   
    'BT_Clear_Cadastro.BackColor = &H8000000F
    'BT_Clear_Cadastro.ForeColor = &H80000012
   
    'BT_SALVAR_CADASTRO.BackColor = &H8000000F
    'BT_SALVAR_CADASTRO.ForeColor = &H80000012
   
End Sub

Function removeAlpha(r As String) As String
    With CreateObject("vbscript.regexp")
     .Pattern = "[9 0]"
     .Global = True
     removeAlpha = .Replace(r, "")
    End With
End Function


Function completeZeroInLeft(id As String, qtd As Integer) As String

    id = Trim(id)
   
    While Len(id) <= qtd
   
       id = "0" & id
      
    Wend
     
    completeZeroInLeft = id
   
End Function

Function getConnection() As ADODB.Connection
   
    Dim obConnection2 As ADODB.Connection
   
    Set obConnection2 = New ADODB.Connection
   
    obConnection2.ConnectionString = "Provider=SAS.IOMProvider;Data Source=iom-bridge://localhost:8591;User ID=fabio;Password=1234"
    'obConnection2.ConnectionString = "Provider=SAS.IOMProvider;Data Source=""iom://svpsas06.abcbrasil.local:8591;Bridge;SECURITYPACKAGE=Negotiate"""
    
    Set getConnection = obConnection2
   
End Function

Function geraCodigo(cd_cli As String)

        cd_cli = Trim(cd_cli)
                    
        If Len(cd_cli) < 5 Then
        
            cd_cli = completeZeroInLeft(cd_cli, 4)
            cd_cli = "500" & cd_cli & "1"
        ElseIf Len(cd_cli) > 7 Then
        
            dig = Right(cd_cli, 1)
            
            If dig < 9 Then
            
                dig = dig + 1
                cd_cli = Left(cd_cli, Len(cd_cli) - 1)
                cd_cli = cd_cli & dig
            
            ElseIf dig = 9 Then
            
                dig = 1
                
                prefixo = Left(cd_cli, 1) + 1
                
                cd_cli = Right(cd_cli, Len(cd_cli) - 1)
                cd_cli = Left(cd_cli, Len(cd_cli) - 1)
                
                cd_cli = prefixo & cd_cli & dig
                
            End If
        
        End If
        
        geraCodigo = cd_cli
End Function

Function geraCodigo2(cd_cli As String, Cd_grp As String)

   cd_cli_aux = Trim(cd_cli)

   cd_cli_aux = Left(Trim(cd_cli), 3)
    
   If cd_cli_aux = "999" Then
   
        cd_cli = Right(cd_cli, 6)
        
        cd_cli = removeZeros(cd_cli) + 1
        
        cd_cli = completeZeroInLeft(cd_cli, 5)
        
        cd_cli = "999" & Trim(completeZeroInLeft(Cd_grp, 5)) & cd_cli
        
   Else
   
       cd_cli = "999" & completeZeroInLeft(Cd_grp, 5) + "000001"
        
   End If
   
   geraCodigo2 = cd_cli
   
End Function

Function removeNumbers(r As String) As String
    With CreateObject("vbscript.regexp")
     .Pattern = "[9 0]"
     .Global = True
     removeNumbers = .Replace(r, "")
    End With
End Function




Function removeZeros(r As String) As String
     removeZeros = CInt(Trim(r))
End Function

Function RemoveDupes(InputArray) As Variant
    Dim Array_2()
    Dim eleArr_1 As Variant
    Dim x As Integer
    x = 0
    On Error Resume Next
    For Each eleArr_1 In InputArray
        If UBound(Filter(InputArray, eleArr_1)) = 0 Then
            ReDim Preserve Array_2(x)
            Array_2(x) = eleArr_1
            x = x + 1
        End If
    Next
    RemoveDupes = Array_2
End Function

Function montaData(dt As Date) As String

  Dim mes As String
  
  Select Case Month(dt)
    Case Is = 1
        mes = "JAN"
    Case Is = 2
        mes = "FEB"
    Case Is = 3
        mes = "MAR"
    Case Is = 4
        mes = "APR"
    Case Is = 5
        mes = "MAY"
    Case Is = 6
        mes = "JUN"
    Case Is = 7
        mes = "JUL"
    Case Is = 8
        mes = "AGO"
    Case Is = 9
        mes = "SEP"
    Case Is = 10
        mes = "OCT"
    Case Is = 11
        mes = "NOV"
    Case Is = 12
        mes = "DEZ"
    End Select
  
  hora = completeZeroInLeft(Hour(dt), 1)
  minuto = completeZeroInLeft(Minute(dt), 1)
  segundo = completeZeroInLeft(Second(dt), 1)
  
  dt_result = completeZeroInLeft(Day(dt), 1) & mes & Year(dt) & ":" & hora & ":" & minuto & ":" & segundo

  montaData = dt_result

End Function

Function getLoginWindows() As String

    getLoginWindows = (Environ$("Username"))

End Function

Function verificaNumeros(r As String) As Boolean
    Set regex = CreateObject("vbscript.regexp")
    With regex
     .Pattern = "[a-z A-Z]"
    End With
    verificaNumeros = regex.Test(r)
End Function

Function populaBalanco(rs As ADODB.Recordset) As Balanco

    Dim blc As Balanco
    
    Set blc = New Balanco
    
    blc.dt_crg = rs![dt_crg]
    blc.Dt_exerc = rs![Dt_exerc]
    blc.Mod_demo = rs![Mod_demo]
    blc.Cd_grp = rs![Cd_grp]
    blc.cd_cli = rs![cd_cli]
    blc.Mes_de_fechamento = rs![Mes_de_fechamento]
    blc.Flg_grp = rs![Flg_grp]
    blc.Flg_fecham_balanc = rs![Flg_fecham_balanc]
    blc.moeda = rs![moeda]
    blc.Auditor = rs![Auditor]
    blc.Analista_planil = rs![Analista_planil]
    blc.Layout_planil = rs![Layout_planil]
    blc.Cnpj = rs![Cnpj]
    blc.Bco_ativo_cart_camb = rs![Bco_ativo_cart_camb]
    blc.Bco_ativo_outros_creds = rs![Bco_ativo_outros_creds]
    blc.Bco_atv_n_circ_part_ctrl_colig = rs![Bco_atv_n_circ_part_ctrl_colig]
    blc.Bco_atv_n_circ_outros_invest = rs![Bco_atv_n_circ_outros_invest]
    blc.Bco_atv_n_circ_invest = rs![Bco_atv_n_circ_invest]
    blc.Bco_atv_n_circ_imob_tec_liq = rs![Bco_atv_n_circ_imob_tec_liq]
    blc.Bco_atv_n_circ_atv_intang = rs![Bco_atv_n_circ_atv_intang]
    blc.Bco_atv_n_circ_atv_perman = rs![Bco_atv_n_circ_atv_perman]
    blc.Bco_atv_n_circ_atv_total = rs![Bco_atv_n_circ_atv_total]
    blc.Bco_ativo_op_arrend_mercatl = rs![Bco_ativo_op_arrend_mercatl]
    blc.Bco_ativo_disp = rs![Bco_ativo_disp]
    blc.Bco_ativo_pdd_arrend_mercatl = rs![Bco_ativo_pdd_arrend_mercatl]
    blc.Bco_pass_circ = rs![Bco_pass_circ]
    blc.Bco_pass_n_circ_part_minor = rs![Bco_pass_n_circ_part_minor]
    blc.Bco_pass_n_circ_ajust_vlr_merc = rs![Bco_pass_n_circ_ajust_vlr_merc]
    blc.Bco_pass_n_circ_lcr_prej_acml = rs![Bco_pass_n_circ_lcr_prej_acml]
    blc.Bco_pass_n_circ_patrim_liq = rs![Bco_pass_n_circ_patrim_liq]
    blc.Bco_pass_n_circ_pass_total = rs![Bco_pass_n_circ_pass_total]
    blc.Bco_pass_depos_aprazo = rs![Bco_pass_depos_aprazo]
    blc.Bco_pass_circ_repass_pais = rs![Bco_pass_circ_repass_pais]
    blc.Bco_dre_lucro_liq = rs![Bco_dre_lucro_liq]
    blc.Bco_dre_empr_cess_repass = rs![Bco_dre_empr_cess_repass]
    blc.Bco_civeis_conting_nao_provs = rs![Bco_civeis_conting_nao_provs]
    blc.Bco_trablstas_conting_nao_provs = rs![Bco_trablstas_conting_nao_provs]
    blc.Bco_fiscais_conting_nao_provs = rs![Bco_fiscais_conting_nao_provs]
    blc.Bco_total_conting_nao_provs = rs![Bco_total_conting_nao_provs]
    blc.Bco_trablstas_conting_provs = rs![Bco_trablstas_conting_provs]
    blc.Bco_fiscais_conting_provs = rs![Bco_fiscais_conting_provs]
    blc.Bco_total_conting_provs = rs![Bco_total_conting_provs]
    blc.Bco_civeis_depos_judc = rs![Bco_civeis_depos_judc]
    blc.Bco_trablstas_depos_judc = rs![Bco_trablstas_depos_judc]
    blc.Bco_fiscais_depos_judc = rs![Bco_fiscais_depos_judc]
    blc.Bco_total_depos_judc = rs![Bco_total_depos_judc]
    blc.Bco_civeis_conting_provs = rs![Bco_civeis_conting_provs]
    blc.Bco_p_negoc_valor_custo = rs![Bco_p_negoc_valor_custo]
    blc.Bco_p_negoc_valor_contab = rs![Bco_p_negoc_valor_contab]
    blc.Bco_p_negoc_mtm = rs![Bco_p_negoc_mtm]
    blc.Bco_disp_venda_valor_custo = rs![Bco_disp_venda_valor_custo]
    blc.Bco_disp_venda_valor_contab = rs![Bco_disp_venda_valor_contab]
    blc.Bco_disp_venda_mtm = rs![Bco_disp_venda_mtm]
    blc.Bco_mtdos_vcto_valor_custo = rs![Bco_mtdos_vcto_valor_custo]
    blc.Bco_dre_outras_rec_interm = rs![Bco_dre_outras_rec_interm]
    blc.Bco_mtdos_vcto_valor_contab = rs![Bco_mtdos_vcto_valor_contab]
    blc.Bco_mtdos_vcto_mtm = rs![Bco_mtdos_vcto_mtm]
    blc.Bco_instr_financ_deriv_vlr_custo = rs![Bco_instr_financ_deriv_vlr_custo]
    blc.Bco_instr_fin_deriv_vlr_contab = rs![Bco_instr_fin_deriv_vlr_contab]
    blc.Bco_instr_financ_deriv_mtm = rs![Bco_instr_financ_deriv_mtm]
    blc.Bco_aa = rs![Bco_aa]
    blc.Bco_total_cart = rs![Bco_total_cart]
    blc.Bco_d_h = rs![Bco_d_h]
    blc.Bco_pdd_exig = rs![Bco_pdd_exig]
    blc.Bco_pdd_const = rs![Bco_pdd_const]
    blc.Bco_vencd = rs![Bco_vencd]
    blc.Bco_vencd_90d = rs![Bco_vencd_90d]
    blc.Bco_a = rs![Bco_a]
    blc.Bco_b = rs![Bco_b]
    blc.Bco_c = rs![Bco_c]
    blc.Bco_d = rs![Bco_d]
    blc.Bco_e = rs![Bco_e]
    blc.Bco_f = rs![Bco_f]
    blc.Bco_g = rs![Bco_g]
    blc.Bco_h = rs![Bco_h]
    blc.Bco_bxdos_sectzdos = rs![Bco_bxdos_sectzdos]
    blc.Bco_ind_basileia_br = rs![Bco_ind_basileia_br]
    blc.Bco_avais_fiancas_prestdos = rs![Bco_avais_fiancas_prestdos]
    blc.Bco_ag = rs![Bco_ag]
    blc.Bco_func = rs![Bco_func]
    blc.Bco_fnds_admn = rs![Bco_fnds_admn]
    blc.Bco_depos_judc = rs![Bco_depos_judc]
    blc.Bco_bndu = rs![Bco_bndu]
    blc.Bco_part_contrlds_clgds = rs![Bco_part_contrlds_clgds]
    blc.Bco_ativo_intang = rs![Bco_ativo_intang]
    blc.Bco_cdi_liqdz_dia = rs![Bco_cdi_liqdz_dia]
    blc.Bco_tvm_vinc_prest_gar_neg = rs![Bco_tvm_vinc_prest_gar_neg]
    blc.Bco_tvm_baixa_liqdz = rs![Bco_tvm_baixa_liqdz]
    blc.Bco_instrm_fin_deriv_pass_neg = rs![Bco_instrm_fin_deriv_pass_neg]
    blc.Bco_capt_merc_aber_neg = rs![Bco_capt_merc_aber_neg]
    blc.Bco_caixa_dispnvl = rs![Bco_caixa_dispnvl]
    blc.Bco_caixa_dispnvl_pl = rs![Bco_caixa_dispnvl_pl]
    blc.Bco_caixa_dispnvl_cart_cred = rs![Bco_caixa_dispnvl_cart_cred]
    blc.Bco_oper_cred_arrend_merctl = rs![Bco_oper_cred_arrend_merctl]
    blc.Bco_pdd_neg = rs![Bco_pdd_neg]
    blc.Bco_ind_basileia = rs![Bco_ind_basileia]
    blc.Bco_cgp_ajust_atv_capt_merc_aber = rs![Bco_cgp_ajust_atv_capt_merc_aber]
    blc.Bco_alvcgem_passva = rs![Bco_alvcgem_passva]
    blc.Bco_alvcgem_cred = rs![Bco_alvcgem_cred]
    blc.Bco_alvcgem_oper = rs![Bco_alvcgem_oper]
    blc.Bco_capt_giro_prop = rs![Bco_capt_giro_prop]
    blc.Bco_capt_giro_prop_ajust = rs![Bco_capt_giro_prop_ajust]
    blc.Bco_cgp_ajust_pl = rs![Bco_cgp_ajust_pl]
    blc.Bco_saldo_inicial = rs![Bco_saldo_inicial]
    blc.Bco_const = rs![Bco_const]
    blc.Bco_reversao = rs![Bco_reversao]
    blc.Bco_baixas = rs![Bco_baixas]
    blc.Bco_saldo_final = rs![Bco_saldo_final]
    blc.Bco_reneg_fluxo = rs![Bco_reneg_fluxo]
    blc.Bco_recup = rs![Bco_recup]
    blc.Bco_lca = rs![Bco_lca]
    blc.Bco_lci = rs![Bco_lci]
    blc.Bco_lf = rs![Bco_lf]
    blc.Bco_total = rs![Bco_total]
    blc.Bco_patr_liq = rs![Bco_patr_liq]
    blc.Bco_pass_circ_emprest_exterior = rs![Bco_pass_circ_emprest_exterior]
    blc.Bco_pass_circ_repass_exterior = rs![Bco_pass_circ_repass_exterior]
    blc.Bco_pass_circ_outras_contas = rs![Bco_pass_circ_outras_contas]
    blc.Bco_parts_relac = rs![Bco_parts_relac]
    blc.Bco_rec_interfin = rs![Bco_rec_interfin]
    blc.Bco_result_nao_operac = rs![Bco_result_nao_operac]
    blc.Bco_lucro_antes_ir = rs![Bco_lucro_antes_ir]
    blc.Bco_ir_cs = rs![Bco_ir_cs]
    blc.Bco_lucro_liq = rs![Bco_lucro_liq]
    blc.Bco_nim = rs![Bco_nim]
    blc.Bco_eficiency_ratio = rs![Bco_eficiency_ratio]
    blc.Bco_roae = rs![Bco_roae]
    blc.Bco_roaa = rs![Bco_roaa]
    blc.Bco_desp_int_financ = rs![Bco_desp_int_financ]
    blc.Bco_res_brto_interm = rs![Bco_res_brto_interm]
    blc.Bco_outras_rec_desp_operac = rs![Bco_outras_rec_desp_operac]
    blc.Bco_res_oper = rs![Bco_res_oper]
    blc.Bco_basileia_tier_i = rs![Bco_basileia_tier_i]
    blc.Bco_dpge_i = rs![Bco_dpge_i]
    blc.Bco_dpge_ii = rs![Bco_dpge_ii]
    blc.Bco_cred_trib = rs![Bco_cred_trib]
    blc.Bco_outros_pass = rs![Bco_outros_pass]
    blc.Bco_aprazo = rs![Bco_aprazo]
    blc.Bco_ativo_cdi = rs![Bco_ativo_cdi]
    blc.Bco_ativo_titulo_merc_abert = rs![Bco_ativo_titulo_merc_abert]
    blc.Bco_ativo_tvm = rs![Bco_ativo_tvm]
    blc.Bco_ativo_operac_cred = rs![Bco_ativo_operac_cred]
    blc.Bco_ativo_pdd = rs![Bco_ativo_pdd]
    blc.Bco_ativo_desp_antec = rs![Bco_ativo_desp_antec]
    blc.Bco_ativo_circ = rs![Bco_ativo_circ]
    blc.Bco_ativo_circ_tvm = rs![Bco_ativo_circ_tvm]
    blc.Bco_ativo_circ_operac_cred = rs![Bco_ativo_circ_operac_cred]
    blc.Bco_ativo_circ_pdd_op_cred = rs![Bco_ativo_circ_pdd_op_cred]
    blc.Bco_ativo_circ_op_arrend_merc = rs![Bco_ativo_circ_op_arrend_merc]
    blc.Bco_ativo_circ_pdd_op_arr_merc = rs![Bco_ativo_circ_pdd_op_arr_merc]
    blc.Bco_ativo_circ_outros_cred = rs![Bco_ativo_circ_outros_cred]
    blc.Bco_atv_n_circ = rs![Bco_atv_n_circ]
    blc.Bco_pass_depos_avista = rs![Bco_pass_depos_avista]
    blc.Bco_pass_poupanca = rs![Bco_pass_poupanca]
    blc.Bco_interfin = rs![Bco_interfin]
    blc.Bco_pass_depos_interfinan = rs![Bco_pass_depos_interfinan]
    blc.Bco_capt_merc_aber = rs![Bco_capt_merc_aber]
    blc.Bco_pass_capt_merc_abert = rs![Bco_pass_capt_merc_abert]
    blc.Bco_pass_circ_emprest_pais = rs![Bco_pass_circ_emprest_pais]
    blc.Bco_outras_contas = rs![Bco_outras_contas]
    blc.Bco_depos = rs![Bco_depos]
    blc.Bco_pass_depos = rs![Bco_pass_depos]
    blc.Bco_pass_emprest_pais = rs![Bco_pass_emprest_pais]
    blc.Bco_pass_emprest_exterior = rs![Bco_pass_emprest_exterior]
    blc.Bco_pass_repass_exterior = rs![Bco_pass_repass_exterior]
    blc.Bco_pass_outras_contas = rs![Bco_pass_outras_contas]
    blc.Bco_pass_n_circ = rs![Bco_pass_n_circ]
    blc.Bco_pass_n_circ_capit_soc = rs![Bco_pass_n_circ_capit_soc]
    blc.Bco_pass_n_circ_reserv_capt = rs![Bco_pass_n_circ_reserv_capt]
    blc.Bco_dre_rec_interm_financ = rs![Bco_dre_rec_interm_financ]
    blc.Bco_dre_tvm = rs![Bco_dre_tvm]
    blc.Bco_dre_capt_merc = rs![Bco_dre_capt_merc]
    blc.Bco_dre_outras_desp_interm = rs![Bco_dre_outras_desp_interm]
    blc.Bco_dre_desp_interm_financ = rs![Bco_dre_desp_interm_financ]
    blc.Bco_dre_res_bruto_interm = rs![Bco_dre_res_bruto_interm]
    blc.Bco_dre_const_pdd = rs![Bco_dre_const_pdd]
    blc.Bco_dre_res_interm_apos_pdd = rs![Bco_dre_res_interm_apos_pdd]
    blc.Bco_dre_rect_prest_serv = rs![Bco_dre_rect_prest_serv]
    blc.Bco_dre_custo_operac = rs![Bco_dre_custo_operac]
    blc.Bco_dre_desp_tribut = rs![Bco_dre_desp_tribut]
    blc.Bco_dre_outras_rect_desp_operac = rs![Bco_dre_outras_rect_desp_operac]
    blc.Bco_dre_res_operac = rs![Bco_dre_res_operac]
    blc.Bco_dre_equiv_patrim = rs![Bco_dre_equiv_patrim]
    blc.Bco_dre_res_apos_equiv_patrim = rs![Bco_dre_res_apos_equiv_patrim]
    blc.Bco_dre_rect_desp_n_operac = rs![Bco_dre_rect_desp_n_operac]
    blc.Bco_dre_lucro_antes_ir = rs![Bco_dre_lucro_antes_ir]
    blc.Bco_dre_impst_renda_ctrl_soc = rs![Bco_dre_impst_renda_ctrl_soc]
    blc.Bco_dre_part = rs![Bco_dre_part]
    blc.Bco_cdi = rs![Bco_cdi]
    blc.Bco_cart_camb = rs![Bco_cart_camb]
    blc.Bco_depos_interfin = rs![Bco_depos_interfin]
    blc.Bco_depos_aprazo = rs![Bco_depos_aprazo]
    blc.Bco_depos_avista = rs![Bco_depos_avista]
    blc.Bco_desp_antpdas = rs![Bco_desp_antpdas]
    blc.Bco_disps = rs![Bco_disps]
    blc.Bco_emprest_extrior = rs![Bco_emprest_extrior]
    blc.Bco_emprest_pais = rs![Bco_emprest_pais]
    blc.Bco_outros = rs![Bco_outros]
    blc.Bco_outros_cred = rs![Bco_outros_cred]
    blc.Bco_pass_cart_camb = rs![Bco_pass_cart_camb]
    blc.Bco_poupanca = rs![Bco_poupanca]
    blc.Bco_repas_extrior = rs![Bco_repas_extrior]
    blc.Bco_repas_pais = rs![Bco_repas_pais]
    blc.Bco_titulo_merc_aber = rs![Bco_titulo_merc_aber]
    blc.Bco_lfsn = rs![Bco_lfsn]
    blc.Bco_letra_cmbio = rs![Bco_letra_cmbio]
    blc.Bco_dre_operac_cred = rs![Bco_dre_operac_cred]
    blc.Bco_tvm_caract_cred_neg = rs![Bco_tvm_caract_cred_neg]
    blc.Bco_tvm = rs![Bco_tvm]
    blc.Bco_div_subord = rs![Bco_div_subord]
    blc.Bco_pdd_avais_fiancas = rs![Bco_pdd_avais_fiancas]
    blc.Bco_pdd_caract_cred = rs![Bco_pdd_caract_cred]
    blc.Bco_pdd_cart_expand = rs![Bco_pdd_cart_expand]
    blc.Bco_tvm_caract_cred = rs![Bco_tvm_caract_cred]
    blc.Bco_reneg_estoq = rs![Bco_reneg_estoq]
    blc.Bco_disp_venda_prov_p_desv = rs![Bco_disp_venda_prov_p_desv]
    blc.Bco_instr_fin_deriv_prov_p_desv = rs![Bco_instr_fin_deriv_prov_p_desv]
    blc.Bco_mtdos_vcto_prov_p_desv = rs![Bco_mtdos_vcto_prov_p_desv]
    blc.Bco_const_pdd = rs![Bco_const_pdd]
    blc.Bco_custo_operac = rs![Bco_custo_operac]
    blc.Bco_equiv_patrim = rs![Bco_equiv_patrim]
    blc.Bco_instrm_fin_deriv = rs![Bco_instrm_fin_deriv]
    blc.Bco_partic = rs![Bco_partic]
    blc.Bco_rec_prest_serv = rs![Bco_rec_prest_serv]
    blc.Bco_desp_trib = rs![Bco_desp_trib]
    blc.Bco_eficiency_ratio_ajus_cli = rs![Bco_eficiency_ratio_ajus_cli]
    blc.Bco_nim_ajust_cli = rs![Bco_nim_ajust_cli]
    blc.Emprs_disps = rs![Emprs_disps]
    blc.Bco_pass_repass_pais = rs![Bco_pass_repass_pais]
    blc.Emprs_desp_antecip = rs![Emprs_desp_antecip]
    blc.Emprs_outros_operac = rs![Emprs_outros_operac]
    blc.Emprs_financ_receb_cp = rs![Emprs_financ_receb_cp]
    blc.Emprs_rol_mensal = rs![Emprs_rol_mensal]
    blc.Emprs_ger_cxa_operac = rs![Emprs_ger_cxa_operac]
    blc.Emprs_desp_financ = rs![Emprs_desp_financ]
    blc.Emprs_rects_financs = rs![Emprs_rects_financs]
    blc.Emprs_ger_cxa_apos_result_fin = rs![Emprs_ger_cxa_apos_result_fin]
    blc.Emprs_invest_imob_diferido = rs![Emprs_invest_imob_diferido]
    blc.Emprs_invest_control_colig = rs![Emprs_invest_control_colig]
    blc.Emprs_result_exerc_fut = rs![Emprs_result_exerc_fut]
    blc.Emprs_patrim_liq = rs![Emprs_patrim_liq]
    blc.Emprs_minoritario = rs![Emprs_minoritario]
    blc.Emprs_rol_anualzdo = rs![Emprs_rol_anualzdo]
    blc.Emprs_result_nao_operac = rs![Emprs_result_nao_operac]
    blc.Emprs_ir_cs = rs![Emprs_ir_cs]
    blc.Emprs_empres_partes_relac = rs![Emprs_empres_partes_relac]
    blc.Emprs_otrs_atv_nao_operac_cp_lp = rs![Emprs_otrs_atv_nao_operac_cp_lp]
    blc.Emprs_otrs_pass_n_operac_cp_lp = rs![Emprs_otrs_pass_n_operac_cp_lp]
    blc.Emprs_var_divid_bancar_liq = rs![Emprs_var_divid_bancar_liq]
    blc.Emprs_Bco_cp = rs![Emprs_Bco_cp]
    blc.Emprs_disp_aplic_financ = rs![Emprs_disp_aplic_financ]
    blc.Emprs_Bco_curto_pl = rs![Emprs_Bco_curto_pl]
    blc.Emprs_Bco_lp = rs![Emprs_Bco_lp]
    blc.Emprs_ebit = rs![Emprs_ebit]
    blc.Emprs_aplic_financ_lp = rs![Emprs_aplic_financ_lp]
    blc.Emprs_Bco_lp_liq = rs![Emprs_Bco_lp_liq]
    blc.Emprs_Bco_liq = rs![Emprs_Bco_liq]
    blc.Emprs_Bco_liq_rol = rs![Emprs_Bco_liq_rol]
    blc.Emprs_Bco_liq_ebtida_anual = rs![Emprs_Bco_liq_ebtida_anual]
    blc.Emprs_ccl = rs![Emprs_ccl]
    blc.Emprs_var_ccl = rs![Emprs_var_ccl]
    blc.Emprs_cgp = rs![Emprs_cgp]
    blc.Emprs_var_cgp = rs![Emprs_var_cgp]
    blc.Emprs_meio_circ = rs![Emprs_meio_circ]
    blc.Emprs_necess_capit_giro = rs![Emprs_necess_capit_giro]
    blc.Emprs_serv_divida = rs![Emprs_serv_divida]
    blc.Emprs_calc_icsd = rs![Emprs_calc_icsd]
    blc.Emprs_receb = rs![Emprs_receb]
    blc.Emprs_p_m_estoq = rs![Emprs_p_m_estoq]
    blc.Emprs_p_m_pagam = rs![Emprs_p_m_pagam]
    blc.Emprs_financ_conced_cp = rs![Emprs_financ_conced_cp]
    blc.Emprs_equiv_patrim_pl = rs![Emprs_equiv_patrim_pl]
    blc.Emprs_prov_dev_duvids = rs![Emprs_prov_dev_duvids]
    blc.Emprs_estoqs = rs![Emprs_estoqs]
    blc.Emprs_adto_forn = rs![Emprs_adto_forn]
    blc.Emprs_tit_val_mobil = rs![Emprs_tit_val_mobil]
    blc.Emprs_desp_pg_antec = rs![Emprs_desp_pg_antec]
    blc.Emprs_ativo_circ = rs![Emprs_ativo_circ]
    blc.Emprs_realzvl_a_l_p = rs![Emprs_realzvl_a_l_p]
    blc.Emprs_part_control_coligs = rs![Emprs_part_control_coligs]
    blc.Emprs_outros_invest = rs![Emprs_outros_invest]
    blc.Emprs_invest = rs![Emprs_invest]
    blc.Emprs_imob_tecn_liq = rs![Emprs_imob_tecn_liq]
    blc.Emprs_ativo_intang = rs![Emprs_ativo_intang]
    blc.Emprs_ativo_perman = rs![Emprs_ativo_perman]
    blc.Emprs_ativo_tot = rs![Emprs_ativo_tot]
    blc.Emprs_forns = rs![Emprs_forns]
    blc.Emprs_obrig_soc_tribut = rs![Emprs_obrig_soc_tribut]
    blc.Emprs_adto_cli = rs![Emprs_adto_cli]
    blc.Emprs_duplic_descts = rs![Emprs_duplic_descts]
    blc.Emprs_camb = rs![Emprs_camb]
    blc.Emprs_emprest_financs = rs![Emprs_emprest_financs]
    blc.Emprs_exig_a_lp = rs![Emprs_exig_a_lp]
    blc.Emprs_res_exerc_fut = rs![Emprs_res_exerc_fut]
    blc.Emprs_capit_soc = rs![Emprs_capit_soc]
    blc.Emprs_reserv_capit_lucro = rs![Emprs_reserv_capit_lucro]
    blc.Emprs_reserv_reaval = rs![Emprs_reserv_reaval]
    blc.Emprs_partic_minor = rs![Emprs_partic_minor]
    blc.Emprs_lucro_prej_acml = rs![Emprs_lucro_prej_acml]
    blc.Emprs_patr_liq = rs![Emprs_patr_liq]
    blc.Emprs_passivo_total = rs![Emprs_passivo_total]
    blc.Emprs_rect_oper_liq = rs![Emprs_rect_oper_liq]
    blc.Emprs_custo_prod_vends = rs![Emprs_custo_prod_vends]
    blc.Emprs_lucro_bruto = rs![Emprs_lucro_bruto]
    blc.Emprs_desp_admins = rs![Emprs_desp_admins]
    blc.Emprs_desp_vndas = rs![Emprs_desp_vndas]
    blc.Emprs_outras_desp_rec_operac = rs![Emprs_outras_desp_rec_operac]
    blc.Emprs_saldo_cor_monet = rs![Emprs_saldo_cor_monet]
    blc.Emprs_lucro_antes_res_finan = rs![Emprs_lucro_antes_res_finan]
    blc.Emprs_rect_financ = rs![Emprs_rect_financ]
    blc.Emprs_desps_financs = rs![Emprs_desps_financs]
    blc.Emprs_var_cambl_liq = rs![Emprs_var_cambl_liq]
    blc.Emprs_rect_desp_nao_operac = rs![Emprs_rect_desp_nao_operac]
    blc.Emprs_lucro_antes_equic_patr = rs![Emprs_lucro_antes_equic_patr]
    blc.Emprs_equiv_patriom = rs![Emprs_equiv_patriom]
    blc.Emprs_lucro_antes_ir = rs![Emprs_lucro_antes_ir]
    blc.Emprs_imp_rnda_contrib_soc = rs![Emprs_imp_rnda_contrib_soc]
    blc.Emprs_partic = rs![Emprs_partic]
    blc.Emprs_lucro_liq = rs![Emprs_lucro_liq]
    blc.Emprs_cli = rs![Emprs_cli]
    blc.Emprs_passivo_circ = rs![Emprs_passivo_circ]
    blc.Emprs_rect_bruta = rs![Emprs_rect_bruta]
    blc.Emprs_devol_abatim = rs![Emprs_devol_abatim]
    blc.Emprs_impos_fatrds = rs![Emprs_impos_fatrds]
    blc.Emprs_deprec = rs![Emprs_deprec]
    blc.Emprs_ebitda = rs![Emprs_ebitda]
    blc.Emprs_ebitda_rol = rs![Emprs_ebitda_rol]
    blc.Emprs_var_camb_liq = rs![Emprs_var_camb_liq]
    blc.Emprs_Bco_liq_aq_terr_ebitda = rs![Emprs_Bco_liq_aq_terr_ebitda]
    blc.Emprs_Bco_liq_aq_terr_pl = rs![Emprs_Bco_liq_aq_terr_pl]
    blc.Emprs_Bco_liq_equiv_patriom = rs![Emprs_Bco_liq_equiv_patriom]
    blc.Emprs_Bco_liq_moagem = rs![Emprs_Bco_liq_moagem]
    blc.Emprs_Bco_liq_patrim_av_cred = rs![Emprs_Bco_liq_patrim_av_cred]
    blc.Emprs_Bco_liq_pl = rs![Emprs_Bco_liq_pl]
    blc.Emprs_Bco_total_liq = rs![Emprs_Bco_total_liq]
    blc.Emprs_divid_pagos = rs![Emprs_divid_pagos]
    blc.Emprs_divid_receb = rs![Emprs_divid_receb]
    blc.Emprs_ebitda_mw_capacid_inst = rs![Emprs_ebitda_mw_capacid_inst]
    blc.Emprs_ativo_difer_pl = rs![Emprs_ativo_difer_pl]
    blc.Emprs_Bco_rol = rs![Emprs_Bco_rol]
    blc.Emprs_ajust1 = rs![Emprs_ajust1]
    blc.Emprs_ajust2 = rs![Emprs_ajust2]
    blc.Emprs_ajust3 = rs![Emprs_ajust3]
    blc.Emprs_Bco_liq_aq_terr_ebtida_aj = rs![Emprs_Bco_liq_aq_terr_ebtida_aj]
    blc.Emprs_Bco_liq_ebitda_ajust = rs![Emprs_Bco_liq_ebitda_ajust]
    blc.Emprs_ebitda_aj_mw_capac_inst = rs![Emprs_ebitda_aj_mw_capac_inst]
    blc.Emprs_Bco_ajust_pl = rs![Emprs_Bco_ajust_pl]
    blc.Emprs_Bco_ajust_pl_rol = rs![Emprs_Bco_ajust_pl_rol]
    blc.Pefis_mm_patr_comprovado = rs![Pefis_mm_patr_comprovado]
    blc.Pefis_mm_liq = rs![Pefis_mm_liq]
    blc.Pefis_mm_ativ_imblz = rs![Pefis_mm_ativ_imblz]
    blc.Pefis_mm_particip_emp = rs![Pefis_mm_particip_emp]
    blc.Pefis_mm_gado = rs![Pefis_mm_gado]
    blc.Pefis_mm_outro = rs![Pefis_mm_outro]
    blc.Pefis_mm_div_bcra = rs![Pefis_mm_div_bcra]
    blc.Pefis_mm_div_avais = rs![Pefis_mm_div_avais]
    blc.Pefis_mm_patr_liq = rs![Pefis_mm_patr_liq]
    blc.Pefis_fi_patr_comprovado = rs![Pefis_fi_patr_comprovado]
    blc.Pefis_fi_liq = rs![Pefis_fi_liq]
    blc.Pefis_fi_ativ_imblz = rs![Pefis_fi_ativ_imblz]
    blc.Pefis_fi_particip_emp = rs![Pefis_fi_particip_emp]
    blc.Pefis_fi_gado = rs![Pefis_fi_gado]
    blc.Pefis_fi_outro = rs![Pefis_fi_outro]
    blc.Pefis_fi_div_bcra = rs![Pefis_fi_div_bcra]
    blc.Pefis_fi_div_avais = rs![Pefis_fi_div_avais]
    blc.Pefis_db_patr_comprovado = rs![Pefis_db_patr_comprovado]
    blc.Pefis_db_liq = rs![Pefis_db_liq]
    blc.Pefis_db_ativ_imblz = rs![Pefis_db_ativ_imblz]
    blc.Pefis_db_particip_emp = rs![Pefis_db_particip_emp]
    blc.Pefis_db_gado = rs![Pefis_db_gado]
    blc.Pefis_db_outro = rs![Pefis_db_outro]
    blc.Pefis_db_div_bcra = rs![Pefis_db_div_bcra]
    blc.Pefis_db_div_avais = rs![Pefis_db_div_avais]
    blc.Pefis_ir_aplic_fin = rs![Pefis_ir_aplic_fin]
    blc.Pefis_ir_qt_acoes_emprs = rs![Pefis_ir_qt_acoes_emprs]
    blc.Pefis_ir_imoveis = rs![Pefis_ir_imoveis]
    blc.Pefis_ir_veiculos = rs![Pefis_ir_veiculos]
    blc.Pefis_ir_emp_terceiro = rs![Pefis_ir_emp_terceiro]
    blc.Pefis_ir_outro = rs![Pefis_ir_outro]
    blc.Pefis_ir_total_bens_dirt = rs![Pefis_ir_total_bens_dirt]
    blc.Pefis_ir_div_onus = rs![Pefis_ir_div_onus]
    blc.Pefis_ir_div_avais = rs![Pefis_ir_div_avais]
    blc.Pefis_ir_patr_liq = rs![Pefis_ir_patr_liq]
    blc.Pefis_arrec = rs![Pefis_arrec]
    blc.Pefis_ardesp = rs![Pefis_ardesp]
    blc.Pefis_arresult = rs![Pefis_arresult]
    blc.Pefis_arbens_ativ_rural = rs![Pefis_arbens_ativ_rural]
    blc.Pefis_ardiv_vin_ativ_rural = rs![Pefis_ardiv_vin_ativ_rural]
    blc.Segur_disp = rs![Segur_disp]
    blc.Segur_cred_oper_previd_compl = rs![Segur_cred_oper_previd_compl]
    blc.Segur_seguradoras = rs![Segur_seguradoras]
    blc.Segur_irb = rs![Segur_irb]
    blc.Segur_desp_comerc_diferd = rs![Segur_desp_comerc_diferd]
    blc.Segur_titulo_vl_mblro = rs![Segur_titulo_vl_mblro]
    blc.Segur_desp_pagto_antcpo = rs![Segur_desp_pagto_antcpo]
    blc.Segur_outra_conta_oper = rs![Segur_outra_conta_oper]
    blc.Segur_outra_conta_nao_oper = rs![Segur_outra_conta_nao_oper]
    blc.Segur_ativ_circ = rs![Segur_ativ_circ]
    blc.Segur_aplic = rs![Segur_aplic]
    blc.Segur_titulo_cred_receb = rs![Segur_titulo_cred_receb]
    blc.Segur_realzv_lp = rs![Segur_realzv_lp]
    blc.Segur_part_ctrl_colgd = rs![Segur_part_ctrl_colgd]
    blc.Segur_outro_invtmo = rs![Segur_outro_invtmo]
    blc.Segur_invtmo = rs![Segur_invtmo]
    blc.Segur_imbro_tecn_liq = rs![Segur_imbro_tecn_liq]
    blc.Segur_ativ_dfrd = rs![Segur_ativ_dfrd]
    blc.Segur_ativ_perman = rs![Segur_ativ_perman]
    blc.Segur_ativ_total = rs![Segur_ativ_total]
    blc.Segur_deb_oper_previd = rs![Segur_deb_oper_previd]
    blc.Segur_obrig_soc_trib = rs![Segur_obrig_soc_trib]
    blc.Segur_sinis_liq = rs![Segur_sinis_liq]
    blc.Segur_emprest_fin = rs![Segur_emprest_fin]
    blc.Segur_prov_tecn = rs![Segur_prov_tecn]
    blc.Segur_depos_terc = rs![Segur_depos_terc]
    blc.Segur_ctrl_colgd = rs![Segur_ctrl_colgd]
    blc.Segur_pasv_circ = rs![Segur_pasv_circ]
    blc.Segur_outra_conta = rs![Segur_outra_conta]
    blc.Segur_exig_lp = rs![Segur_exig_lp]
    blc.Segur_res_exerc_fut = rs![Segur_res_exerc_fut]
    blc.Segur_capital_soc = rs![Segur_capital_soc]
    blc.Segur_res_capital_lcr = rs![Segur_res_capital_lcr]
    blc.Segur_res_reaval = rs![Segur_res_reaval]
    blc.Segur_particip_mntro = rs![Segur_particip_mntro]
    blc.Segur_lcr_prej_acum = rs![Segur_lcr_prej_acum]
    blc.Segur_patr_liq = rs![Segur_patr_liq]
    blc.Segur_pasv_total = rs![Segur_pasv_total]
    blc.Segur_renda_contrib = rs![Segur_renda_contrib]
    blc.Segur_contrib_rps = rs![Segur_contrib_rps]
    blc.Segur_var_prov_premios = rs![Segur_var_prov_premios]
    blc.Segur_rec_oper_liq = rs![Segur_rec_oper_liq]
    blc.Segur_desp_benef_resgt = rs![Segur_desp_benef_resgt]
    blc.Segur_var_prov_evento_nao_avis = rs![Segur_var_prov_evento_nao_avis]
    blc.Segur_lcr_bruto = rs![Segur_lcr_bruto]
    blc.Segur_desp_adm = rs![Segur_desp_adm]
    blc.Segur_desp_vda = rs![Segur_desp_vda]
    blc.Segur_outro_desp_rec_oper = rs![Segur_outro_desp_rec_oper]
    blc.Segur_saldo_correc_monet = rs![Segur_saldo_correc_monet]
    blc.Segur_lcr_antes_res_fin = rs![Segur_lcr_antes_res_fin]
    blc.Segur_rect_fin = rs![Segur_rect_fin]
    blc.Segur_desp_fin = rs![Segur_desp_fin]
    blc.Segur_rec_desp_nao_oper = rs![Segur_rec_desp_nao_oper]
    blc.Segur_lcr_antes_equiv_patrim = rs![Segur_lcr_antes_equiv_patrim]
    blc.Segur_equiv_patrim = rs![Segur_equiv_patrim]
    blc.Segur_lcr_antes_ir = rs![Segur_lcr_antes_ir]
    blc.Segur_ir_renda_contrib_soc = rs![Segur_ir_renda_contrib_soc]
    blc.Segur_particip = rs![Segur_particip]
    blc.Segur_lcr_liq = rs![Segur_lcr_liq]
    blc.Op_ativ_circ = rs![Op_ativ_circ]
    blc.Op_cx_equivl_cx = rs![Op_cx_equivl_cx]
    blc.Op_cred_a_cp = rs![Op_cred_a_cp]
    blc.Op_estoq = rs![Op_estoq]
    blc.Op_vpd_pagas_antecip = rs![Op_vpd_pagas_antecip]
    blc.Op_ativ_n_circ = rs![Op_ativ_n_circ]
    blc.Op_ativ_realzvl_lp = rs![Op_ativ_realzvl_lp]
    blc.Op_cred_a_lp = rs![Op_cred_a_lp]
    blc.Op_div_ativ_tribtr = rs![Op_div_ativ_tribtr]
    blc.Op_empre_financ_conced = rs![Op_empre_financ_conced]
    blc.Op_aj_perda_cred_lp = rs![Op_aj_perda_cred_lp]
    blc.Op_demais_cred_vlrs_lp = rs![Op_demais_cred_vlrs_lp]
    blc.Op_investimentos = rs![Op_investimentos]
    blc.Op_imobilizado = rs![Op_imobilizado]
    blc.Op_intangivel = rs![Op_intangivel]
    blc.Op_total_ativo = rs![Op_total_ativo]
    blc.Op_passivo_circulante = rs![Op_passivo_circulante]
    blc.Op_obrig_trab_prev_assist_cp = rs![Op_obrig_trab_prev_assist_cp]
    blc.Op_forn_ctas_pg_cp = rs![Op_forn_ctas_pg_cp]
    blc.Op_obrig_fiscais_cp = rs![Op_obrig_fiscais_cp]
    blc.Op_prov_cp = rs![Op_prov_cp]
    blc.Op_demais_obrig_cp = rs![Op_demais_obrig_cp]
    blc.Op_emprest_finan_cp = rs![Op_emprest_finan_cp]
    blc.Op_passivo_n_circ = rs![Op_passivo_n_circ]
    blc.Op_emprest_financ_lp = rs![Op_emprest_financ_lp]
    blc.Op_fornecedores_lp = rs![Op_fornecedores_lp]
    blc.Op_obrig_ficais_lp = rs![Op_obrig_ficais_lp]
    blc.Op_previsoes_lp = rs![Op_previsoes_lp]
    blc.Op_patrimonio_lp = rs![Op_patrimonio_lp]
    blc.Op_result_acumulados = rs![Op_result_acumulados]
    blc.Op_total_passivo = rs![Op_total_passivo]
    blc.Op_receitas_correntes = rs![Op_receitas_correntes]
    blc.Op_tributarias = rs![Op_tributarias]
    blc.Op_contribuicoes = rs![Op_contribuicoes]
    blc.Op_transf_correntes = rs![Op_transf_correntes]
    blc.Op_patrimoniais = rs![Op_patrimoniais]
    blc.Op_outras_receitas_correntes = rs![Op_outras_receitas_correntes]
    blc.Op_deducoes = rs![Op_deducoes]
    blc.Op_despesas_correntes = rs![Op_despesas_correntes]
    blc.Op_pessoal_encargos_sociais = rs![Op_pessoal_encargos_sociais]
    blc.Op_juros_encargos_dividas = rs![Op_juros_encargos_dividas]
    blc.Op_transferencias_correntes = rs![Op_transferencias_correntes]
    blc.Op_outras_despesas_correntes = rs![Op_outras_despesas_correntes]
    blc.Op_saldo_corrente = rs![Op_saldo_corrente]
    blc.Op_receitas_capital = rs![Op_receitas_capital]
    blc.Op_operacoes_credito = rs![Op_operacoes_credito]
    blc.Op_alienacao_bens = rs![Op_alienacao_bens]
    blc.Op_transferencia_capital = rs![Op_transferencia_capital]
    blc.Op_receita_capital_outras = rs![Op_receita_capital_outras]
    blc.Op_despesas_capital = rs![Op_despesas_capital]
    blc.Op_inversoes_financeiras = rs![Op_inversoes_financeiras]
    blc.Op_amortizacao_divida = rs![Op_amortizacao_divida]
    blc.Op_outras_despesas_capital = rs![Op_outras_despesas_capital]
    blc.Op_outras_receitas_despesas = rs![Op_outras_receitas_despesas]
    blc.Op_reservas_contingencias = rs![Op_reservas_contingencias]
    blc.Op_deficit_superavit = rs![Op_deficit_superavit]
    blc.Op_orcado_receitas_correntes = rs![Op_orcado_receitas_correntes]
    blc.Op_orcado_tributarias = rs![Op_orcado_tributarias]
    blc.Op_orcado_contribuicoes = rs![Op_orcado_contribuicoes]
    blc.Op_orcado_transf_correntes = rs![Op_orcado_transf_correntes]
    blc.Op_orcado_patrimoniais = rs![Op_orcado_patrimoniais]
    blc.Op_orcado_outras_rect_corren = rs![Op_orcado_outras_rect_corren]
    blc.Op_orcado_deducoes = rs![Op_orcado_deducoes]
    blc.Op_orcado_despesas_correntes = rs![Op_orcado_despesas_correntes]
    blc.Op_orcado_pess_encg_div = rs![Op_orcado_pess_encg_div]
    blc.Op_orcado_juros_encarg_div = rs![Op_orcado_juros_encarg_div]
    blc.Op_orcado_transf_corr = rs![Op_orcado_transf_corr]
    blc.Op_orcado_outras_desp_corr = rs![Op_orcado_outras_desp_corr]
    blc.Op_orcado_saldo_corr = rs![Op_orcado_saldo_corr]
    blc.Op_orcado_rect_capital = rs![Op_orcado_rect_capital]
    blc.Op_orcado_oper_credito = rs![Op_orcado_oper_credito]
    blc.Op_orcado_alienacao_bens = rs![Op_orcado_alienacao_bens]
    blc.Op_orcado_transf_capital = rs![Op_orcado_transf_capital]
    blc.Op_orcado_rect_capital_outras = rs![Op_orcado_rect_capital_outras]
    blc.Op_orcado_despesas_capital = rs![Op_orcado_despesas_capital]
    blc.Op_orcado_investimentos = rs![Op_orcado_investimentos]
    blc.Op_orcado_inversoes_fin = rs![Op_orcado_inversoes_fin]
    blc.Op_orcado_amort_divida = rs![Op_orcado_amort_divida]
    blc.Op_orcado_outras_desp_capital = rs![Op_orcado_outras_desp_capital]
    blc.Op_orcado_outras_rect_despesas = rs![Op_orcado_outras_rect_despesas]
    blc.Op_orcado_reservas_conti = rs![Op_orcado_reservas_conti]
    blc.Op_orcado_defict_superavt = rs![Op_orcado_defict_superavt]
    blc.Op_gap_receitas_correntes = rs![Op_gap_receitas_correntes]
    blc.Op_gap_tributarias = rs![Op_gap_tributarias]
    blc.Op_gap_contribuicoes = rs![Op_gap_contribuicoes]
    blc.Op_gap_transf_correntes = rs![Op_gap_transf_correntes]
    blc.Op_gap_patrimoniais = rs![Op_gap_patrimoniais]
    blc.Op_gap_outras_rect_corren = rs![Op_gap_outras_rect_corren]
    blc.Op_gap_deducoes = rs![Op_gap_deducoes]
    blc.Op_gap_despesas_correntes = rs![Op_gap_despesas_correntes]
    blc.Op_gap_pess_encg_div = rs![Op_gap_pess_encg_div]
    blc.Op_gap_juros_encarg_div = rs![Op_gap_juros_encarg_div]
    blc.Op_gap_transf_corr = rs![Op_gap_transf_corr]
    blc.Op_gap_outras_desp_corr = rs![Op_gap_outras_desp_corr]
    blc.Op_gap_saldo_corr = rs![Op_gap_saldo_corr]
    blc.Op_gap_rect_capital = rs![Op_gap_rect_capital]
    blc.Op_gap_oper_credito = rs![Op_gap_oper_credito]
    blc.Op_gap_alienacao_bens = rs![Op_gap_alienacao_bens]
    blc.Op_gap_transf_capital = rs![Op_gap_transf_capital]
    blc.Op_gap_rect_capital_outras = rs![Op_gap_rect_capital_outras]
    blc.Op_gap_despesas_capital = rs![Op_gap_despesas_capital]
    blc.Op_gap_investimentos = rs![Op_gap_investimentos]
    blc.Op_gap_inversoes_fin = rs![Op_gap_inversoes_fin]
    blc.Op_gap_amort_divida = rs![Op_gap_amort_divida]
    blc.Op_gap_outras_desp_capital = rs![Op_gap_outras_desp_capital]
    blc.Op_gap_outras_rect_despesas = rs![Op_gap_outras_rect_despesas]
    blc.Op_gap_reservas_conti = rs![Op_gap_reservas_conti]
    blc.op_gap_defict_superavt = rs![op_gap_defict_superavt]
    blc.Bco_tot_valor_custo = rs![Bco_tot_valor_custo]
    blc.Bco_tot_valor_contab = rs![Bco_tot_valor_contab]
    blc.Bco_tot_mtm = rs![Bco_tot_mtm]
    blc.Bco_tot_perc = rs![Bco_tot_perc]
    blc.Bco_tot_prov_p_desv = rs![Bco_tot_prov_p_desv]
    blc.Emprs_cp_derivat = rs![Emprs_cp_derivat]
    blc.Emprs_lp_derivat = rs![Emprs_lp_derivat]
    
    Set populaBalanco = blc

End Function


Function salvaBalanco(blc As Balanco) As String



End Function



Function somaBalanco(dt As String, cod_cli As String, cod_grupo As String, mes_fech As String, moeda As String, dt_carga As String) As String

    
        qry = qry & "insert into lb_plani.fato_balanco_aux_3(cd_cli,dt_crg,cd_grp,dt_exerc,MES_DE_FECHAMENTO,MOEDA,BCO_ATIVO_CART_CAMB,BCO_ATIVO_OUTROS_CREDS,"
        qry = qry & "BCO_ATV_N_CIRC_PART_CTRL_COLIG,BCO_ATV_N_CIRC_OUTROS_INVEST,BCO_ATV_N_CIRC_INVEST,BCO_ATV_N_CIRC_IMOB_TEC_LIQ,BCO_ATV_N_CIRC_ATV_INTANG,"
        qry = qry & "BCO_ATV_N_CIRC_ATV_PERMAN,BCO_ATV_N_CIRC_ATV_TOTAL,BCO_ATIVO_OP_ARREND_MERCATL,BCO_ATIVO_DISP,BCO_ATIVO_PDD_ARREND_MERCATL,BCO_PASS_CIRC,"
        qry = qry & "BCO_PASS_N_CIRC_PART_MINOR,BCO_PASS_N_CIRC_AJUST_VLR_MERC,BCO_PASS_N_CIRC_LCR_PREJ_ACML,BCO_PASS_N_CIRC_PATRIM_LIQ,BCO_PASS_N_CIRC_PASS_TOTAL,"
        qry = qry & "BCO_PASS_DEPOS_APRAZO,BCO_PASS_CIRC_REPASS_PAIS,BCO_DRE_LUCRO_LIQ,BCO_DRE_EMPR_CESS_REPASS,BCO_CIVEIS_CONTING_NAO_PROVS,BCO_TRABLSTAS_CONTING_NAO_PROVS,"
        qry = qry & "BCO_FISCAIS_CONTING_NAO_PROVS,BCO_TOTAL_CONTING_NAO_PROVS,BCO_TRABLSTAS_CONTING_PROVS,BCO_FISCAIS_CONTING_PROVS,BCO_TOTAL_CONTING_PROVS,"
        qry = qry & "BCO_CIVEIS_DEPOS_JUDC,BCO_TRABLSTAS_DEPOS_JUDC,BCO_FISCAIS_DEPOS_JUDC,BCO_TOTAL_DEPOS_JUDC,BCO_CIVEIS_CONTING_PROVS,BCO_P_NEGOC_VALOR_CUSTO,"
        qry = qry & "BCO_P_NEGOC_VALOR_CONTAB,BCO_P_NEGOC_MTM,BCO_DISP_VENDA_VALOR_CUSTO,BCO_DISP_VENDA_VALOR_CONTAB,BCO_DISP_VENDA_MTM,BCO_MTDOS_VCTO_VALOR_CUSTO,"
        qry = qry & "BCO_DRE_OUTRAS_REC_INTERM,BCO_MTDOS_VCTO_VALOR_CONTAB,BCO_MTDOS_VCTO_MTM,BCO_INSTR_FINANC_DERIV_VLR_CUSTO,BCO_INSTR_FIN_DERIV_VLR_CONTAB,"
        qry = qry & "BCO_INSTR_FINANC_DERIV_MTM,BCO_AA,BCO_TOTAL_CART,BCO_D_H,BCO_PDD_EXIG,BCO_PDD_CONST,BCO_VENCD,BCO_VENCD_90D,BCO_A,BCO_B,BCO_C,BCO_D,"
        qry = qry & "BCO_E,BCO_F,BCO_G,BCO_H,BCO_BXDOS_SECTZDOS,BCO_IND_BASILEIA_BR,BCO_AVAIS_FIANCAS_PRESTDOS,BCO_PASS_REPASS_PAIS,BCO_AG,BCO_FUNC,"
        qry = qry & "BCO_FNDS_ADMN,BCO_DEPOS_JUDC,BCO_BNDU,BCO_PART_CONTRLDS_CLGDS,BCO_ATIVO_INTANG,BCO_CDI_LIQDZ_DIA,BCO_TVM_VINC_PREST_GAR_NEG,BCO_TVM_BAIXA_LIQDZ,"
        qry = qry & "BCO_INSTRM_FIN_DERIV_PASS_NEG,BCO_CAPT_MERC_ABER_NEG,BCO_CAIXA_DISPNVL,BCO_CAIXA_DISPNVL_PL,BCO_CAIXA_DISPNVL_CART_CRED,BCO_OPER_CRED_ARREND_MERCTL,"
        qry = qry & "BCO_PDD_NEG,BCO_IND_BASILEIA,BCO_CGP_AJUST_ATV_CAPT_MERC_ABER,BCO_ALVCGEM_PASSVA,BCO_ALVCGEM_CRED,BCO_ALVCGEM_OPER,BCO_CAPT_GIRO_PROP,"
        qry = qry & "BCO_CAPT_GIRO_PROP_AJUST,BCO_CGP_AJUST_PL,BCO_SALDO_INICIAL,BCO_CONST,BCO_REVERSAO,BCO_BAIXAS,BCO_SALDO_FINAL,BCO_RENEG_FLUXO,BCO_RECUP,"
        qry = qry & "BCO_LCA,BCO_LCI,BCO_LF,BCO_TOTAL,BCO_PATR_LIQ,BCO_PASS_CIRC_EMPREST_EXTERIOR,BCO_PASS_CIRC_REPASS_EXTERIOR,BCO_PASS_CIRC_OUTRAS_CONTAS,"
        qry = qry & "BCO_PARTS_RELAC,BCO_REC_INTERFIN,BCO_RESULT_NAO_OPERAC,BCO_LUCRO_ANTES_IR,BCO_IR_CS,BCO_LUCRO_LIQ,BCO_NIM,BCO_EFICIENCY_RATIO,BCO_ROAE,BCO_ROAA,"
        qry = qry & "BCO_DESP_INT_FINANC,BCO_RES_BRTO_INTERM,BCO_OUTRAS_REC_DESP_OPERAC,BCO_RES_OPER,BCO_BASILEIA_TIER_I,BCO_DPGE_I,BCO_DPGE_II,BCO_CRED_TRIB,"
        qry = qry & "BCO_OUTROS_PASS,BCO_APRAZO,BCO_ATIVO_CDI,BCO_ATIVO_TITULO_MERC_ABERT,BCO_ATIVO_TVM,BCO_ATIVO_OPERAC_CRED,BCO_ATIVO_PDD,BCO_ATIVO_DESP_ANTEC,"
        qry = qry & "BCO_ATIVO_CIRC,BCO_ATIVO_CIRC_TVM,BCO_ATIVO_CIRC_OPERAC_CRED,BCO_ATIVO_CIRC_PDD_OP_CRED,BCO_ATIVO_CIRC_OP_ARREND_MERC,BCO_ATIVO_CIRC_PDD_OP_ARR_MERC,"
        qry = qry & "BCO_ATIVO_CIRC_OUTROS_CRED,BCO_ATV_N_CIRC,BCO_PASS_DEPOS_AVISTA,BCO_PASS_POUPANCA,BCO_INTERFIN,BCO_PASS_DEPOS_INTERFINAN,BCO_CAPT_MERC_ABER,"
        qry = qry & "BCO_PASS_CAPT_MERC_ABERT,BCO_PASS_CIRC_EMPREST_PAIS,BCO_OUTRAS_CONTAS,BCO_DEPOS,BCO_PASS_DEPOS,BCO_PASS_EMPREST_PAIS,BCO_PASS_EMPREST_EXTERIOR,"
        qry = qry & "BCO_PASS_REPASS_EXTERIOR,BCO_PASS_OUTRAS_CONTAS,BCO_PASS_N_CIRC,BCO_PASS_N_CIRC_CAPIT_SOC,BCO_PASS_N_CIRC_RESERV_CAPT,BCO_DRE_REC_INTERM_FINANC,"
        qry = qry & "BCO_DRE_TVM,BCO_DRE_CAPT_MERC,BCO_DRE_OUTRAS_DESP_INTERM,BCO_DRE_DESP_INTERM_FINANC,BCO_DRE_RES_BRUTO_INTERM,BCO_DRE_CONST_PDD,BCO_DRE_RES_INTERM_APOS_PDD,"
        qry = qry & "BCO_DRE_RECT_PREST_SERV,BCO_DRE_CUSTO_OPERAC,BCO_DRE_DESP_TRIBUT,BCO_DRE_OUTRAS_RECT_DESP_OPERAC,BCO_DRE_RES_OPERAC,BCO_DRE_EQUIV_PATRIM,"
        qry = qry & "BCO_DRE_RES_APOS_EQUIV_PATRIM,BCO_DRE_RECT_DESP_N_OPERAC,BCO_DRE_LUCRO_ANTES_IR,BCO_DRE_IMPST_RENDA_CTRL_SOC,BCO_DRE_PART,BCO_CDI,BCO_CART_CAMB,"
        qry = qry & "BCO_DEPOS_INTERFIN,BCO_DEPOS_APRAZO,BCO_DEPOS_AVISTA,BCO_DESP_ANTPDAS,BCO_DISPS,BCO_EMPREST_EXTRIOR,BCO_EMPREST_PAIS,BCO_OUTROS,BCO_OUTROS_CRED,"
        qry = qry & "BCO_PASS_CART_CAMB,BCO_POUPANCA,BCO_REPAS_EXTRIOR,BCO_REPAS_PAIS,BCO_TITULO_MERC_ABER,BCO_LFSN,BCO_LETRA_CMBIO,BCO_DRE_OPERAC_CRED,"
        qry = qry & "BCO_TVM_CARACT_CRED_NEG,BCO_TVM,BCO_DIV_SUBORD,BCO_PDD_AVAIS_FIANCAS,BCO_PDD_CARACT_CRED,BCO_PDD_CART_EXPAND,BCO_TVM_CARACT_CRED,BCO_RENEG_ESTOQ,"
        qry = qry & "BCO_DISP_VENDA_PROV_P_DESV,BCO_INSTR_FIN_DERIV_PROV_P_DESV,BCO_MTDOS_VCTO_PROV_P_DESV,BCO_P_NEGOC_PROV_P_DESV,BCO_PERC_DISP_VENDA,BCO_PERC_INSTR_FINANC,"
        qry = qry & "BCO_PERC_MTDOS_VCTO,BCO_PERC_P_NEGOC,BCO_TOT_VALOR_CUSTO,BCO_TOT_VALOR_CONTAB,BCO_TOT_MTM,BCO_TOT_PERC,BCO_TOT_PROV_P_DESV,BCO_CONST_PDD,"
        qry = qry & "BCO_CUSTO_OPERAC,BCO_EQUIV_PATRIM,BCO_INSTRM_FIN_DERIV,BCO_CREDORES_CRED_C_OBRIG,BCO_PARTIC,BCO_REC_PREST_SERV,BCO_DESP_TRIB,"
        qry = qry & "BCO_EFICIENCY_RATIO_AJUS_CLI,BCO_NIM_AJUST_CLI,EMPRS_DISPS,EMPRS_DESP_ANTECIP,EMPRS_OUTROS_OPERAC,EMPRS_FINANC_RECEB_CP,EMPRS_ROL_MENSAL,"
        qry = qry & "EMPRS_GER_CXA_OPERAC,EMPRS_DESP_FINANC,EMPRS_RECTS_FINANCS,EMPRS_GER_CXA_APOS_RESULT_FIN,EMPRS_INVEST_IMOB_DIFERIDO,EMPRS_INVEST_CONTROL_COLIG,"
        qry = qry & "EMPRS_RESULT_EXERC_FUT,EMPRS_PATRIM_LIQ,EMPRS_MINORITARIO,EMPRS_ROL_ANUALZDO,EMPRS_RESULT_NAO_OPERAC,EMPRS_IR_CS,EMPRS_EMPRES_PARTES_RELAC,"
        qry = qry & "EMPRS_OTRS_ATV_NAO_OPERAC_CP_LP,EMPRS_OTRS_PASS_N_OPERAC_CP_LP,EMPRS_VAR_DIVID_BANCAR_LIQ,EMPRS_BCO_CP,EMPRS_DISP_APLIC_FINANC,EMPRS_BCO_CURTO_PL,"
        qry = qry & "EMPRS_BCO_LP,EMPRS_EBIT,EMPRS_APLIC_FINANC_LP,EMPRS_BCO_LP_LIQ,EMPRS_BCO_LIQ,EMPRS_BCO_LIQ_ROL,EMPRS_BCO_LIQ_EBTIDA_ANUAL,EMPRS_CCL,EMPRS_VAR_CCL,"
        qry = qry & "EMPRS_CGP,EMPRS_VAR_CGP,EMPRS_MEIO_CIRC,EMPRS_NECESS_CAPIT_GIRO,EMPRS_SERV_DIVIDA,EMPRS_CALC_ICSD,EMPRS_RECEB,EMPRS_P_M_ESTOQ,EMPRS_P_M_PAGAM,"
        qry = qry & "EMPRS_FINANC_CONCED_CP,EMPRS_EQUIV_PATRIM_PL,EMPRS_ATIVO_CC_CONTROL_COLIG,EMPRS_PASS_CC_CONTROL_COLIG,EMPRS_ATIVO_OUTRAS_CONTAS,EMPRS_PASS_OUTRAS_CONTAS,"
        qry = qry & "EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER,EMPRS_PASS_OUTRAS_CTAS_NAO_OPER,EMPRS_ATIVO_OUTRAS_CTAS_OPERAC,EMPRS_PASS_OUTRAS_CTAS_OPERAC,EMPRS_PROV_DEV_DUVIDS,"
        qry = qry & "EMPRS_ESTOQS,EMPRS_ADTO_FORN,EMPRS_TIT_VAL_MOBIL,EMPRS_DESP_PG_ANTEC,EMPRS_ATIVO_CIRC,EMPRS_REALZVL_A_L_P,EMPRS_PART_CONTROL_COLIGS,EMPRS_OUTROS_INVEST,"
        qry = qry & "EMPRS_INVEST,EMPRS_IMOB_TECN_LIQ,EMPRS_ATIVO_INTANG,EMPRS_ATIVO_PERMAN,EMPRS_ATIVO_TOT,EMPRS_FORNS,EMPRS_OBRIG_SOC_TRIBUT,EMPRS_ADTO_CLI,"
        qry = qry & "EMPRS_DUPLIC_DESCTS,EMPRS_CAMB,EMPRS_EMPREST_FINANCS,EMPRS_EXIG_A_LP,EMPRS_RES_EXERC_FUT,EMPRS_CAPIT_SOC,EMPRS_RESERV_CAPIT_LUCRO,"
        qry = qry & "EMPRS_RESERV_REAVAL,EMPRS_PARTIC_MINOR,EMPRS_LUCRO_PREJ_ACML,EMPRS_PATR_LIQ,EMPRS_PASSIVO_TOTAL,EMPRS_RECT_OPER_LIQ,EMPRS_CUSTO_PROD_VENDS,"
        qry = qry & "EMPRS_LUCRO_BRUTO,EMPRS_DESP_ADMINS,EMPRS_DESP_VNDAS,EMPRS_OUTRAS_DESP_REC_OPERAC,EMPRS_SALDO_COR_MONET,EMPRS_LUCRO_ANTES_RES_FINAN,EMPRS_RECT_FINANC,"
        qry = qry & "EMPRS_DESPS_FINANCS,EMPRS_VAR_CAMBL_LIQ,EMPRS_RECT_DESP_NAO_OPERAC,EMPRS_LUCRO_ANTES_EQUIC_PATR,EMPRS_EQUIV_PATRIOM,EMPRS_LUCRO_ANTES_IR,"
        qry = qry & "EMPRS_IMP_RNDA_CONTRIB_SOC,EMPRS_PARTIC,EMPRS_LUCRO_LIQ,EMPRS_CLI,EMPRS_PASSIVO_CIRC,EMPRS_RECT_BRUTA,EMPRS_DEVOL_ABATIM,EMPRS_IMPOS_FATRDS,"
        qry = qry & "EMPRS_DEPREC,EMPRS_EBITDA,EMPRS_EBITDA_ROL,EMPRS_VAR_CAMB_LIQ,EMPRS_BCO_LIQ_AQ_TERR_EBITDA,EMPRS_BCO_LIQ_AQ_TERR_PL,EMPRS_BCO_LIQ_EQUIV_PATRIOM,"
        qry = qry & "EMPRS_BCO_LIQ_MOAGEM,EMPRS_BCO_LIQ_PATRIM_AV_CRED,EMPRS_BCO_LIQ_PL,EMPRS_BCO_TOTAL_LIQ,EMPRS_CP_DERIVAT,EMPRS_LP_DERIVAT,EMPRS_DIVID_PAGOS,"
        qry = qry & "EMPRS_DIVID_RECEB,EMPRS_EBITDA_MW_CAPACID_INST,EMPRS_ATIVO_DIFER_PL,EMPRS_BCO_ROL,EMPRS_AJUST1,EMPRS_AJUST2,EMPRS_AJUST3,EMPRS_BCO_LIQ_AQ_TERR_EBTIDA_AJ,"
        qry = qry & "EMPRS_BCO_LIQ_EBITDA_AJUST,EMPRS_EBITDA_AJ_MW_CAPAC_INST,EMPRS_BCO_AJUST_PL,EMPRS_BCO_AJUST_PL_ROL,PEFIS_MM_PATR_COMPROVADO,PEFIS_MM_LIQ,"
        qry = qry & "PEFIS_MM_ATIV_IMBLZ,PEFIS_MM_PARTICIP_EMP,PEFIS_MM_GADO,PEFIS_MM_OUTRO,PEFIS_MM_DIV_BCRA,PEFIS_MM_DIV_AVAIS,PEFIS_MM_PATR_LIQ,PEFIS_FI_PATR_COMPROVADO,"
        qry = qry & "PEFIS_FI_LIQ,PEFIS_FI_ATIV_IMBLZ,PEFIS_FI_PARTICIP_EMP,PEFIS_FI_GADO,PEFIS_FI_OUTRO,PEFIS_FI_DIV_BCRA,PEFIS_FI_DIV_AVAIS,PEFIS_DB_PATR_COMPROVADO,"
        qry = qry & "PEFIS_DB_LIQ,PEFIS_DB_ATIV_IMBLZ,PEFIS_DB_PARTICIP_EMP,PEFIS_DB_GADO,PEFIS_DB_OUTRO,PEFIS_DB_DIV_BCRA,PEFIS_DB_DIV_AVAIS,PEFIS_IR_APLIC_FIN,"
        qry = qry & "PEFIS_IR_QT_ACOES_EMPRS,PEFIS_IR_IMOVEIS,PEFIS_IR_VEICULOS,PEFIS_IR_EMP_TERCEIRO,PEFIS_IR_OUTRO,PEFIS_IR_TOTAL_BENS_DIRT,PEFIS_IR_DIV_ONUS,"
        qry = qry & "PEFIS_IR_DIV_AVAIS,PEFIS_IR_PATR_LIQ,PEFIS_ARREC,PEFIS_ARDESP,PEFIS_ARRESULT,PEFIS_ARBENS_ATIV_RURAL,PEFIS_ARDIV_VIN_ATIV_RURAL,SEGUR_DISP,"
        qry = qry & "SEGUR_CRED_OPER_PREVID_COMPL,SEGUR_SEGURADORAS,SEGUR_IRB,SEGUR_DESP_COMERC_DIFERD,SEGUR_TITULO_VL_MBLRO,SEGUR_DESP_PAGTO_ANTCPO,SEGUR_OUTRA_CONTA_OPER,"
        qry = qry & "SEGUR_OUTRA_CONTA_NAO_OPER,SEGUR_ATIV_CIRC,SEGUR_APLIC,SEGUR_TITULO_CRED_RECEB,SEGUR_REALZV_LP,SEGUR_PART_CTRL_COLGD,SEGUR_OUTRO_INVTMO,SEGUR_INVTMO,"
        qry = qry & "SEGUR_IMBRO_TECN_LIQ,SEGUR_ATIV_DFRD,SEGUR_ATIV_PERMAN,SEGUR_ATIV_TOTAL,SEGUR_DEB_OPER_PREVID,SEGUR_OBRIG_SOC_TRIB,SEGUR_SINIS_LIQ,SEGUR_EMPREST_FIN,"
        qry = qry & "SEGUR_PROV_TECN,SEGUR_DEPOS_TERC,SEGUR_CTRL_COLGD,SEGUR_PASV_CIRC,SEGUR_OUTRA_CONTA,SEGUR_EXIG_LP,SEGUR_RES_EXERC_FUT,SEGUR_CAPITAL_SOC,"
        qry = qry & "SEGUR_RES_CAPITAL_LCR,SEGUR_RES_REAVAL,SEGUR_PARTICIP_MNTRO,SEGUR_LCR_PREJ_ACUM,SEGUR_PATR_LIQ,SEGUR_PASV_TOTAL,SEGUR_RENDA_CONTRIB,SEGUR_CONTRIB_RPS,"
        qry = qry & "SEGUR_VAR_PROV_PREMIOS,SEGUR_REC_OPER_LIQ,SEGUR_DESP_BENEF_RESGT,SEGUR_VAR_PROV_EVENTO_NAO_AVIS,SEGUR_LCR_BRUTO,SEGUR_DESP_ADM,SEGUR_DESP_VDA,"
        qry = qry & "SEGUR_OUTRO_DESP_REC_OPER,SEGUR_SALDO_CORREC_MONET,SEGUR_LCR_ANTES_RES_FIN,SEGUR_RECT_FIN,SEGUR_DESP_FIN,SEGUR_REC_DESP_NAO_OPER,"
        qry = qry & "SEGUR_LCR_ANTES_EQUIV_PATRIM,SEGUR_EQUIV_PATRIM,SEGUR_LCR_ANTES_IR,SEGUR_IR_RENDA_CONTRIB_SOC,SEGUR_PARTICIP,SEGUR_LCR_LIQ,OP_DISPNVL,OP_CRED_A_CP,"
        qry = qry & "OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP,OP_ATIVO_INVESTIMENTOS,OP_ATIVO_CIRC_ESTOQ,OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP,OP_CRED_A_LP,OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP,"
        qry = qry & "OP_INVESTIMENTOS,OP_ATV_RLZ_ESTOQ,OP_ATV_RLZ_VPD_PAGAS_ANTECIP,OP_IMOBILIZADO,OP_INTANGIVEL,OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP,OP_EMPREST_FINAN_CP,"
        qry = qry & "OP_FORN_CTAS_PG_CP,OP_OBRIG_FISCAIS_CP,OP_OBRIG_REPART,OP_PROV_CP,OP_DEMAIS_OBRIG_CP,OP_PASS_N_CIRC_OB_TRB_PREV_AS_CP,OP_EMPREST_FINANC_LP,"
        qry = qry & "OP_FORNECEDORES_LP,OP_PREVISOES_LP,OP_DEMAIS_OBRIG_LP,OP_PATRIMONIO_LP,OP_TRIBUTARIAS,OP_CONTRIBUICOES,OP_TRANSF_CORRENTES,OP_PATRIMONIAIS,"
        qry = qry & "OP_OUTRAS_RECEITAS_CORRENTES,OP_DEDUCOES,OP_PESSOAL_ENCARGOS_SOCIAIS,OP_JUROS_ENCARGOS_DIVIDAS,OP_TRANSFERENCIAS_CORRENTES,OP_OUTRAS_DESPESAS_CORRENTES,"
        qry = qry & "OP_OPERACOES_CREDITO,OP_ALIENACAO_BENS,OP_TRANSFERENCIA_CAPITAL,OP_RECEITA_CAPITAL_OUTRAS,OP_INVERSOES_FINANCEIRAS,OP_AMORTIZACAO_DIVIDA,"
        qry = qry & "OP_OUTRAS_DESPESAS_CAPITAL,OP_OUTRAS_RECEITAS_DESPESAS,OP_RESERVAS_CONTINGENCIAS,OP_ORCADO_TRIBUTARIAS,OP_ORCADO_CONTRIBUICOES,OP_ORCADO_TRANSF_CORRENTES,"
        qry = qry & "OP_ORCADO_PATRIMONIAIS,OP_ORCADO_OUTRAS_RECT_CORREN,OP_ORCADO_DEDUCOES,OP_ORCADO_PESS_ENCG_DIV,OP_ORCADO_JUROS_ENCARG_DIV,OP_ORCADO_TRANSF_CORR,"
        qry = qry & "OP_ORCADO_OUTRAS_DESP_CORR,OP_ORCADO_OPER_CREDITO,OP_ORCADO_ALIENACAO_BENS,OP_ORCADO_TRANSF_CAPITAL,OP_ORCADO_RECT_CAPITAL_OUTRAS,OP_ORCADO_INVESTIMENTOS,"
        qry = qry & "OP_ORCADO_INVERSOES_FIN,OP_ORCADO_AMORT_DIVIDA,OP_ORCADO_OUTRAS_DESP_CAPITAL,OP_ORCADO_OUTRAS_RECT_DESPESAS,OP_ORCADO_RESERVAS_CONTI)"
        qry = qry & "select * from ("
        qry = qry & "select " & cod_cli & " AS cd_cli,'" & dt_carga & "'dt AS dt_crg,cd_grp,dt_exerc," & mes_fech & " AS MES_DE_FECHAMENTO, '" & moeda & "' AS MOEDA,sum(BCO_ATIVO_CART_CAMB) AS BCO_ATIVO_CART_CAMB,"
        qry = qry & "sum(BCO_ATIVO_OUTROS_CREDS) AS BCO_ATIVO_OUTROS_CREDS,sum(BCO_ATV_N_CIRC_PART_CTRL_COLIG) AS BCO_ATV_N_CIRC_PART_CTRL_COLIG,"
        qry = qry & "sum(BCO_ATV_N_CIRC_OUTROS_INVEST) AS BCO_ATV_N_CIRC_OUTROS_INVEST,sum(BCO_ATV_N_CIRC_INVEST) AS BCO_ATV_N_CIRC_INVEST,"
        qry = qry & "sum(BCO_ATV_N_CIRC_IMOB_TEC_LIQ) AS BCO_ATV_N_CIRC_IMOB_TEC_LIQ,sum(BCO_ATV_N_CIRC_ATV_INTANG) AS BCO_ATV_N_CIRC_ATV_INTANG,"
        qry = qry & "sum(BCO_ATV_N_CIRC_ATV_PERMAN) AS BCO_ATV_N_CIRC_ATV_PERMAN,sum(BCO_ATV_N_CIRC_ATV_TOTAL) AS BCO_ATV_N_CIRC_ATV_TOTAL,"
        qry = qry & "sum(BCO_ATIVO_OP_ARREND_MERCATL) AS BCO_ATIVO_OP_ARREND_MERCATL,sum(BCO_ATIVO_DISP) AS BCO_ATIVO_DISP,"
        qry = qry & "sum(BCO_ATIVO_PDD_ARREND_MERCATL) AS BCO_ATIVO_PDD_ARREND_MERCATL,sum(BCO_PASS_CIRC) AS BCO_PASS_CIRC,"
        qry = qry & "sum(BCO_PASS_N_CIRC_PART_MINOR) AS BCO_PASS_N_CIRC_PART_MINOR,sum(BCO_PASS_N_CIRC_AJUST_VLR_MERC) AS BCO_PASS_N_CIRC_AJUST_VLR_MERC,"
        qry = qry & "sum(BCO_PASS_N_CIRC_LCR_PREJ_ACML) AS BCO_PASS_N_CIRC_LCR_PREJ_ACML,sum(BCO_PASS_N_CIRC_PATRIM_LIQ) AS BCO_PASS_N_CIRC_PATRIM_LIQ,"
        qry = qry & "sum(BCO_PASS_N_CIRC_PASS_TOTAL) AS BCO_PASS_N_CIRC_PASS_TOTAL,sum(BCO_PASS_DEPOS_APRAZO) AS BCO_PASS_DEPOS_APRAZO,"
        qry = qry & "sum(BCO_PASS_CIRC_REPASS_PAIS) AS BCO_PASS_CIRC_REPASS_PAIS,sum(BCO_DRE_LUCRO_LIQ) AS BCO_DRE_LUCRO_LIQ,"
        qry = qry & "sum(BCO_DRE_EMPR_CESS_REPASS) AS BCO_DRE_EMPR_CESS_REPASS,sum(BCO_CIVEIS_CONTING_NAO_PROVS) AS BCO_CIVEIS_CONTING_NAO_PROVS,"
        qry = qry & "sum(BCO_TRABLSTAS_CONTING_NAO_PROVS) AS BCO_TRABLSTAS_CONTING_NAO_PROVS,sum(BCO_FISCAIS_CONTING_NAO_PROVS) AS BCO_FISCAIS_CONTING_NAO_PROVS,"
        qry = qry & "sum(BCO_TOTAL_CONTING_NAO_PROVS) AS BCO_TOTAL_CONTING_NAO_PROVS,sum(BCO_TRABLSTAS_CONTING_PROVS) AS BCO_TRABLSTAS_CONTING_PROVS,"
        qry = qry & "sum(BCO_FISCAIS_CONTING_PROVS) AS BCO_FISCAIS_CONTING_PROVS,sum(BCO_TOTAL_CONTING_PROVS) AS BCO_TOTAL_CONTING_PROVS,"
        qry = qry & "sum(BCO_CIVEIS_DEPOS_JUDC) AS BCO_CIVEIS_DEPOS_JUDC,sum(BCO_TRABLSTAS_DEPOS_JUDC) AS BCO_TRABLSTAS_DEPOS_JUDC,"
        qry = qry & "sum(BCO_FISCAIS_DEPOS_JUDC) AS BCO_FISCAIS_DEPOS_JUDC,sum(BCO_TOTAL_DEPOS_JUDC) AS BCO_TOTAL_DEPOS_JUDC,"
        qry = qry & "sum(BCO_CIVEIS_CONTING_PROVS) AS BCO_CIVEIS_CONTING_PROVS,sum(BCO_P_NEGOC_VALOR_CUSTO) AS BCO_P_NEGOC_VALOR_CUSTO,"
        qry = qry & "sum(BCO_P_NEGOC_VALOR_CONTAB) AS BCO_P_NEGOC_VALOR_CONTAB,sum(BCO_P_NEGOC_MTM) AS BCO_P_NEGOC_MTM,sum(BCO_DISP_VENDA_VALOR_CUSTO) AS BCO_DISP_VENDA_VALOR_CUSTO,"
        qry = qry & "sum(BCO_DISP_VENDA_VALOR_CONTAB) AS BCO_DISP_VENDA_VALOR_CONTAB,sum(BCO_DISP_VENDA_MTM) AS BCO_DISP_VENDA_MTM,"
        qry = qry & "sum(BCO_MTDOS_VCTO_VALOR_CUSTO) AS BCO_MTDOS_VCTO_VALOR_CUSTO,sum(BCO_DRE_OUTRAS_REC_INTERM) AS BCO_DRE_OUTRAS_REC_INTERM,"
        qry = qry & "sum(BCO_MTDOS_VCTO_VALOR_CONTAB) AS BCO_MTDOS_VCTO_VALOR_CONTAB,sum(BCO_MTDOS_VCTO_MTM) AS BCO_MTDOS_VCTO_MTM,"
        qry = qry & "sum(BCO_INSTR_FINANC_DERIV_VLR_CUSTO) AS BCO_INSTR_FINANC_DERIV_VLR_CUSTO,sum(BCO_INSTR_FIN_DERIV_VLR_CONTAB) AS BCO_INSTR_FIN_DERIV_VLR_CONTAB,"
        qry = qry & "sum(BCO_INSTR_FINANC_DERIV_MTM) AS BCO_INSTR_FINANC_DERIV_MTM,sum(BCO_AA) AS BCO_AA,sum(BCO_TOTAL_CART) AS BCO_TOTAL_CART,sum(BCO_D_H) AS BCO_D_H,"
        qry = qry & "sum(BCO_PDD_EXIG) AS BCO_PDD_EXIG,sum(BCO_PDD_CONST) AS BCO_PDD_CONST,sum(BCO_VENCD) AS BCO_VENCD,sum(BCO_VENCD_90D) AS BCO_VENCD_90D,sum(BCO_A) AS BCO_A,"
        qry = qry & "sum(BCO_B) AS BCO_B,sum(BCO_C) AS BCO_C,sum(BCO_D) AS BCO_D,sum(BCO_E) AS BCO_E,sum(BCO_F) AS BCO_F,sum(BCO_G) AS BCO_G,sum(BCO_H) AS BCO_H,"
        qry = qry & "sum(BCO_BXDOS_SECTZDOS) AS BCO_BXDOS_SECTZDOS,sum(BCO_IND_BASILEIA_BR) AS BCO_IND_BASILEIA_BR,sum(BCO_AVAIS_FIANCAS_PRESTDOS) AS BCO_AVAIS_FIANCAS_PRESTDOS,"
        qry = qry & "sum(BCO_PASS_REPASS_PAIS) AS BCO_PASS_REPASS_PAIS,sum(BCO_AG) AS BCO_AG,sum(BCO_FUNC) AS BCO_FUNC,sum(BCO_FNDS_ADMN) AS BCO_FNDS_ADMN,"
        qry = qry & "sum(BCO_DEPOS_JUDC) AS BCO_DEPOS_JUDC,sum(BCO_BNDU) AS BCO_BNDU,sum(BCO_PART_CONTRLDS_CLGDS) AS BCO_PART_CONTRLDS_CLGDS,"
        qry = qry & "sum(BCO_ATIVO_INTANG) AS BCO_ATIVO_INTANG,sum(BCO_CDI_LIQDZ_DIA) AS BCO_CDI_LIQDZ_DIA,sum(BCO_TVM_VINC_PREST_GAR_NEG) AS BCO_TVM_VINC_PREST_GAR_NEG,"
        qry = qry & "sum(BCO_TVM_BAIXA_LIQDZ) AS BCO_TVM_BAIXA_LIQDZ,sum(BCO_INSTRM_FIN_DERIV_PASS_NEG) AS BCO_INSTRM_FIN_DERIV_PASS_NEG,"
        qry = qry & "sum(BCO_CAPT_MERC_ABER_NEG) AS BCO_CAPT_MERC_ABER_NEG,sum(BCO_CAIXA_DISPNVL) AS BCO_CAIXA_DISPNVL,sum(BCO_CAIXA_DISPNVL_PL) AS BCO_CAIXA_DISPNVL_PL,"
        qry = qry & "sum(BCO_CAIXA_DISPNVL_CART_CRED) AS BCO_CAIXA_DISPNVL_CART_CRED,sum(BCO_OPER_CRED_ARREND_MERCTL) AS BCO_OPER_CRED_ARREND_MERCTL,"
        qry = qry & "sum(BCO_PDD_NEG) AS BCO_PDD_NEG,sum(BCO_IND_BASILEIA) AS BCO_IND_BASILEIA,sum(BCO_CGP_AJUST_ATV_CAPT_MERC_ABER) AS BCO_CGP_AJUST_ATV_CAPT_MERC_ABER,"
        qry = qry & "sum(BCO_ALVCGEM_PASSVA) AS BCO_ALVCGEM_PASSVA,sum(BCO_ALVCGEM_CRED) AS BCO_ALVCGEM_CRED,sum(BCO_ALVCGEM_OPER) AS BCO_ALVCGEM_OPER,"
        qry = qry & "sum(BCO_CAPT_GIRO_PROP) AS BCO_CAPT_GIRO_PROP,sum(BCO_CAPT_GIRO_PROP_AJUST) AS BCO_CAPT_GIRO_PROP_AJUST,sum(BCO_CGP_AJUST_PL) AS BCO_CGP_AJUST_PL,"
        qry = qry & "sum(BCO_SALDO_INICIAL) AS BCO_SALDO_INICIAL,sum(BCO_CONST) AS BCO_CONST,sum(BCO_REVERSAO) AS BCO_REVERSAO,sum(BCO_BAIXAS) AS BCO_BAIXAS,"
        qry = qry & "sum(BCO_SALDO_FINAL) AS BCO_SALDO_FINAL,sum(BCO_RENEG_FLUXO) AS BCO_RENEG_FLUXO,sum(BCO_RECUP) AS BCO_RECUP,sum(BCO_LCA) AS BCO_LCA,"
        qry = qry & "sum(BCO_LCI) AS BCO_LCI,sum(BCO_LF) AS BCO_LF,sum(BCO_TOTAL) AS BCO_TOTAL,sum(BCO_PATR_LIQ) AS BCO_PATR_LIQ,"
        qry = qry & "sum(BCO_PASS_CIRC_EMPREST_EXTERIOR) AS BCO_PASS_CIRC_EMPREST_EXTERIOR,sum(BCO_PASS_CIRC_REPASS_EXTERIOR) AS BCO_PASS_CIRC_REPASS_EXTERIOR,"
        qry = qry & "sum(BCO_PASS_CIRC_OUTRAS_CONTAS) AS BCO_PASS_CIRC_OUTRAS_CONTAS,sum(BCO_PARTS_RELAC) AS BCO_PARTS_RELAC,sum(BCO_REC_INTERFIN) AS BCO_REC_INTERFIN,"
        qry = qry & "sum(BCO_RESULT_NAO_OPERAC) AS BCO_RESULT_NAO_OPERAC,sum(BCO_LUCRO_ANTES_IR) AS BCO_LUCRO_ANTES_IR,sum(BCO_IR_CS) AS BCO_IR_CS,"
        qry = qry & "sum(BCO_LUCRO_LIQ) AS BCO_LUCRO_LIQ,sum(BCO_NIM) AS BCO_NIM,sum(BCO_EFICIENCY_RATIO) AS BCO_EFICIENCY_RATIO,sum(BCO_ROAE) AS BCO_ROAE,"
        qry = qry & "sum(BCO_ROAA) AS BCO_ROAA,sum(BCO_DESP_INT_FINANC) AS BCO_DESP_INT_FINANC,sum(BCO_RES_BRTO_INTERM) AS BCO_RES_BRTO_INTERM,"
        qry = qry & "sum(BCO_OUTRAS_REC_DESP_OPERAC) AS BCO_OUTRAS_REC_DESP_OPERAC,sum(BCO_RES_OPER) AS BCO_RES_OPER,sum(BCO_BASILEIA_TIER_I) AS BCO_BASILEIA_TIER_I,"
        qry = qry & "sum(BCO_DPGE_I) AS BCO_DPGE_I,sum(BCO_DPGE_II) AS BCO_DPGE_II,sum(BCO_CRED_TRIB) AS BCO_CRED_TRIB,sum(BCO_OUTROS_PASS) AS BCO_OUTROS_PASS,"
        qry = qry & "sum(BCO_APRAZO) AS BCO_APRAZO,sum(BCO_ATIVO_CDI) AS BCO_ATIVO_CDI,sum(BCO_ATIVO_TITULO_MERC_ABERT) AS BCO_ATIVO_TITULO_MERC_ABERT,"
        qry = qry & "sum(BCO_ATIVO_TVM) AS BCO_ATIVO_TVM,sum(BCO_ATIVO_OPERAC_CRED) AS BCO_ATIVO_OPERAC_CRED,sum(BCO_ATIVO_PDD) AS BCO_ATIVO_PDD,"
        qry = qry & "sum(BCO_ATIVO_DESP_ANTEC) AS BCO_ATIVO_DESP_ANTEC,sum(BCO_ATIVO_CIRC) AS BCO_ATIVO_CIRC,sum(BCO_ATIVO_CIRC_TVM) AS BCO_ATIVO_CIRC_TVM,"
        qry = qry & "sum(BCO_ATIVO_CIRC_OPERAC_CRED) AS BCO_ATIVO_CIRC_OPERAC_CRED,sum(BCO_ATIVO_CIRC_PDD_OP_CRED) AS BCO_ATIVO_CIRC_PDD_OP_CRED,"
        qry = qry & "sum(BCO_ATIVO_CIRC_OP_ARREND_MERC) AS BCO_ATIVO_CIRC_OP_ARREND_MERC,sum(BCO_ATIVO_CIRC_PDD_OP_ARR_MERC) AS BCO_ATIVO_CIRC_PDD_OP_ARR_MERC,"
        qry = qry & "sum(BCO_ATIVO_CIRC_OUTROS_CRED) AS BCO_ATIVO_CIRC_OUTROS_CRED,sum(BCO_ATV_N_CIRC) AS BCO_ATV_N_CIRC,sum(BCO_PASS_DEPOS_AVISTA) AS BCO_PASS_DEPOS_AVISTA,"
        qry = qry & "sum(BCO_PASS_POUPANCA) AS BCO_PASS_POUPANCA,sum(BCO_INTERFIN) AS BCO_INTERFIN,sum(BCO_PASS_DEPOS_INTERFINAN) AS BCO_PASS_DEPOS_INTERFINAN,"
        qry = qry & "sum(BCO_CAPT_MERC_ABER) AS BCO_CAPT_MERC_ABER,sum(BCO_PASS_CAPT_MERC_ABERT) AS BCO_PASS_CAPT_MERC_ABERT,"
        qry = qry & "sum(BCO_PASS_CIRC_EMPREST_PAIS) AS BCO_PASS_CIRC_EMPREST_PAIS,sum(BCO_OUTRAS_CONTAS) AS BCO_OUTRAS_CONTAS,sum(BCO_DEPOS) AS BCO_DEPOS,"
        qry = qry & "sum(BCO_PASS_DEPOS) AS BCO_PASS_DEPOS,sum(BCO_PASS_EMPREST_PAIS) AS BCO_PASS_EMPREST_PAIS,sum(BCO_PASS_EMPREST_EXTERIOR) AS BCO_PASS_EMPREST_EXTERIOR,"
        qry = qry & "sum(BCO_PASS_REPASS_EXTERIOR) AS BCO_PASS_REPASS_EXTERIOR,sum(BCO_PASS_OUTRAS_CONTAS) AS BCO_PASS_OUTRAS_CONTAS,sum(BCO_PASS_N_CIRC) AS BCO_PASS_N_CIRC,"
        qry = qry & "sum(BCO_PASS_N_CIRC_CAPIT_SOC) AS BCO_PASS_N_CIRC_CAPIT_SOC,sum(BCO_PASS_N_CIRC_RESERV_CAPT) AS BCO_PASS_N_CIRC_RESERV_CAPT,"
        qry = qry & "sum(BCO_DRE_REC_INTERM_FINANC) AS BCO_DRE_REC_INTERM_FINANC,sum(BCO_DRE_TVM) AS BCO_DRE_TVM,sum(BCO_DRE_CAPT_MERC) AS BCO_DRE_CAPT_MERC,"
        qry = qry & "sum(BCO_DRE_OUTRAS_DESP_INTERM) AS BCO_DRE_OUTRAS_DESP_INTERM,sum(BCO_DRE_DESP_INTERM_FINANC) AS BCO_DRE_DESP_INTERM_FINANC,"
        qry = qry & "sum(BCO_DRE_RES_BRUTO_INTERM) AS BCO_DRE_RES_BRUTO_INTERM,sum(BCO_DRE_CONST_PDD) AS BCO_DRE_CONST_PDD,"
        qry = qry & "sum(BCO_DRE_RES_INTERM_APOS_PDD) AS BCO_DRE_RES_INTERM_APOS_PDD,sum(BCO_DRE_RECT_PREST_SERV) AS BCO_DRE_RECT_PREST_SERV,"
        qry = qry & "sum(BCO_DRE_CUSTO_OPERAC) AS BCO_DRE_CUSTO_OPERAC,sum(BCO_DRE_DESP_TRIBUT) AS BCO_DRE_DESP_TRIBUT,"
        qry = qry & "sum(BCO_DRE_OUTRAS_RECT_DESP_OPERAC) AS BCO_DRE_OUTRAS_RECT_DESP_OPERAC,sum(BCO_DRE_RES_OPERAC) AS BCO_DRE_RES_OPERAC,"
        qry = qry & "sum(BCO_DRE_EQUIV_PATRIM) AS BCO_DRE_EQUIV_PATRIM,sum(BCO_DRE_RES_APOS_EQUIV_PATRIM) AS BCO_DRE_RES_APOS_EQUIV_PATRIM,"
        qry = qry & "sum(BCO_DRE_RECT_DESP_N_OPERAC) AS BCO_DRE_RECT_DESP_N_OPERAC,sum(BCO_DRE_LUCRO_ANTES_IR) AS BCO_DRE_LUCRO_ANTES_IR,"
        qry = qry & "sum(BCO_DRE_IMPST_RENDA_CTRL_SOC) AS BCO_DRE_IMPST_RENDA_CTRL_SOC,sum(BCO_DRE_PART) AS BCO_DRE_PART,sum(BCO_CDI) AS BCO_CDI,"
        qry = qry & "sum(BCO_CART_CAMB) AS BCO_CART_CAMB,sum(BCO_DEPOS_INTERFIN) AS BCO_DEPOS_INTERFIN,sum(BCO_DEPOS_APRAZO) AS BCO_DEPOS_APRAZO,"
        qry = qry & "sum(BCO_DEPOS_AVISTA) AS BCO_DEPOS_AVISTA,sum(BCO_DESP_ANTPDAS) AS BCO_DESP_ANTPDAS,sum(BCO_DISPS) AS BCO_DISPS,"
        qry = qry & "sum(BCO_EMPREST_EXTRIOR) AS BCO_EMPREST_EXTRIOR,sum(BCO_EMPREST_PAIS) AS BCO_EMPREST_PAIS,sum(BCO_OUTROS) AS BCO_OUTROS,"
        qry = qry & "sum(BCO_OUTROS_CRED) AS BCO_OUTROS_CRED,sum(BCO_PASS_CART_CAMB) AS BCO_PASS_CART_CAMB,sum(BCO_POUPANCA) AS BCO_POUPANCA,"
        qry = qry & "sum(BCO_REPAS_EXTRIOR) AS BCO_REPAS_EXTRIOR,sum(BCO_REPAS_PAIS) AS BCO_REPAS_PAIS,sum(BCO_TITULO_MERC_ABER) AS BCO_TITULO_MERC_ABER,"
        qry = qry & "sum(BCO_LFSN) AS BCO_LFSN,sum(BCO_LETRA_CMBIO) AS BCO_LETRA_CMBIO,sum(BCO_DRE_OPERAC_CRED) AS BCO_DRE_OPERAC_CRED,"
        qry = qry & "sum(BCO_TVM_CARACT_CRED_NEG) AS BCO_TVM_CARACT_CRED_NEG,sum(BCO_TVM) AS BCO_TVM,sum(BCO_DIV_SUBORD) AS BCO_DIV_SUBORD,"
        qry = qry & "sum(BCO_PDD_AVAIS_FIANCAS) AS BCO_PDD_AVAIS_FIANCAS,sum(BCO_PDD_CARACT_CRED) AS BCO_PDD_CARACT_CRED,"
        qry = qry & "sum(BCO_PDD_CART_EXPAND) AS BCO_PDD_CART_EXPAND,sum(BCO_TVM_CARACT_CRED) AS BCO_TVM_CARACT_CRED,"
        qry = qry & "sum(BCO_RENEG_ESTOQ) AS BCO_RENEG_ESTOQ,sum(BCO_DISP_VENDA_PROV_P_DESV) AS BCO_DISP_VENDA_PROV_P_DESV,"
        qry = qry & "sum(BCO_INSTR_FIN_DERIV_PROV_P_DESV) AS BCO_INSTR_FIN_DERIV_PROV_P_DESV,sum(BCO_MTDOS_VCTO_PROV_P_DESV) AS BCO_MTDOS_VCTO_PROV_P_DESV,"
        qry = qry & "sum(BCO_P_NEGOC_PROV_P_DESV) AS BCO_P_NEGOC_PROV_P_DESV,sum(BCO_PERC_DISP_VENDA) AS BCO_PERC_DISP_VENDA,sum(BCO_PERC_INSTR_FINANC) AS BCO_PERC_INSTR_FINANC,"
        qry = qry & "sum(BCO_PERC_MTDOS_VCTO) AS BCO_PERC_MTDOS_VCTO,sum(BCO_PERC_P_NEGOC) AS BCO_PERC_P_NEGOC,sum(BCO_TOT_VALOR_CUSTO) AS BCO_TOT_VALOR_CUSTO,"
        qry = qry & "sum(BCO_TOT_VALOR_CONTAB) AS BCO_TOT_VALOR_CONTAB,sum(BCO_TOT_MTM) AS BCO_TOT_MTM,sum(BCO_TOT_PERC) AS BCO_TOT_PERC,"
        qry = qry & "sum(BCO_TOT_PROV_P_DESV) AS BCO_TOT_PROV_P_DESV,sum(BCO_CONST_PDD) AS BCO_CONST_PDD,sum(BCO_CUSTO_OPERAC) AS BCO_CUSTO_OPERAC,"
        qry = qry & "sum(BCO_EQUIV_PATRIM) AS BCO_EQUIV_PATRIM,sum(BCO_INSTRM_FIN_DERIV) AS BCO_INSTRM_FIN_DERIV,sum(BCO_CREDORES_CRED_C_OBRIG) AS BCO_CREDORES_CRED_C_OBRIG,"
        qry = qry & "sum(BCO_PARTIC) AS BCO_PARTIC,sum(BCO_REC_PREST_SERV) AS BCO_REC_PREST_SERV,sum(BCO_DESP_TRIB) AS BCO_DESP_TRIB,"
        qry = qry & "sum(BCO_EFICIENCY_RATIO_AJUS_CLI) AS BCO_EFICIENCY_RATIO_AJUS_CLI,sum(BCO_NIM_AJUST_CLI) AS BCO_NIM_AJUST_CLI,"
        qry = qry & "sum(EMPRS_DISPS) AS EMPRS_DISPS,sum(EMPRS_DESP_ANTECIP) AS EMPRS_DESP_ANTECIP,sum(EMPRS_OUTROS_OPERAC) AS EMPRS_OUTROS_OPERAC,"
        qry = qry & "sum(EMPRS_FINANC_RECEB_CP) AS EMPRS_FINANC_RECEB_CP,sum(EMPRS_ROL_MENSAL) AS EMPRS_ROL_MENSAL,sum(EMPRS_GER_CXA_OPERAC) AS EMPRS_GER_CXA_OPERAC,"
        qry = qry & "sum(EMPRS_DESP_FINANC) AS EMPRS_DESP_FINANC,sum(EMPRS_RECTS_FINANCS) AS EMPRS_RECTS_FINANCS,"
        qry = qry & "sum(EMPRS_GER_CXA_APOS_RESULT_FIN) AS EMPRS_GER_CXA_APOS_RESULT_FIN,sum(EMPRS_INVEST_IMOB_DIFERIDO) AS EMPRS_INVEST_IMOB_DIFERIDO,"
        qry = qry & "sum(EMPRS_INVEST_CONTROL_COLIG) AS EMPRS_INVEST_CONTROL_COLIG,sum(EMPRS_RESULT_EXERC_FUT) AS EMPRS_RESULT_EXERC_FUT,"
        qry = qry & "sum(EMPRS_PATRIM_LIQ) AS EMPRS_PATRIM_LIQ,sum(EMPRS_MINORITARIO) AS EMPRS_MINORITARIO,sum(EMPRS_ROL_ANUALZDO) AS EMPRS_ROL_ANUALZDO,"
        qry = qry & "sum(EMPRS_RESULT_NAO_OPERAC) AS EMPRS_RESULT_NAO_OPERAC,sum(EMPRS_IR_CS) AS EMPRS_IR_CS,sum(EMPRS_EMPRES_PARTES_RELAC) AS EMPRS_EMPRES_PARTES_RELAC,"
        qry = qry & "sum(EMPRS_OTRS_ATV_NAO_OPERAC_CP_LP) AS EMPRS_OTRS_ATV_NAO_OPERAC_CP_LP,sum(EMPRS_OTRS_PASS_N_OPERAC_CP_LP) AS EMPRS_OTRS_PASS_N_OPERAC_CP_LP,"
        qry = qry & "sum(EMPRS_VAR_DIVID_BANCAR_LIQ) AS EMPRS_VAR_DIVID_BANCAR_LIQ,sum(EMPRS_BCO_CP) AS EMPRS_BCO_CP,sum(EMPRS_DISP_APLIC_FINANC) AS EMPRS_DISP_APLIC_FINANC,"
        qry = qry & "sum(EMPRS_BCO_CURTO_PL) AS EMPRS_BCO_CURTO_PL,sum(EMPRS_BCO_LP) AS EMPRS_BCO_LP,sum(EMPRS_EBIT) AS EMPRS_EBIT,"
        qry = qry & "sum(EMPRS_APLIC_FINANC_LP) AS EMPRS_APLIC_FINANC_LP,sum(EMPRS_BCO_LP_LIQ) AS EMPRS_BCO_LP_LIQ,sum(EMPRS_BCO_LIQ) AS EMPRS_BCO_LIQ,"
        qry = qry & "sum(EMPRS_BCO_LIQ_ROL) AS EMPRS_BCO_LIQ_ROL,sum(EMPRS_BCO_LIQ_EBTIDA_ANUAL) AS EMPRS_BCO_LIQ_EBTIDA_ANUAL,sum(EMPRS_CCL) AS EMPRS_CCL,"
        qry = qry & "sum(EMPRS_VAR_CCL) AS EMPRS_VAR_CCL,sum(EMPRS_CGP) AS EMPRS_CGP,sum(EMPRS_VAR_CGP) AS EMPRS_VAR_CGP,sum(EMPRS_MEIO_CIRC) AS EMPRS_MEIO_CIRC,"
        qry = qry & "sum(EMPRS_NECESS_CAPIT_GIRO) AS EMPRS_NECESS_CAPIT_GIRO,sum(EMPRS_SERV_DIVIDA) AS EMPRS_SERV_DIVIDA,sum(EMPRS_CALC_ICSD) AS EMPRS_CALC_ICSD,"
        qry = qry & "sum(EMPRS_RECEB) AS EMPRS_RECEB,sum(EMPRS_P_M_ESTOQ) AS EMPRS_P_M_ESTOQ,sum(EMPRS_P_M_PAGAM) AS EMPRS_P_M_PAGAM,"
        qry = qry & "sum(EMPRS_FINANC_CONCED_CP) AS EMPRS_FINANC_CONCED_CP,sum(EMPRS_EQUIV_PATRIM_PL) AS EMPRS_EQUIV_PATRIM_PL,"
        qry = qry & "sum(EMPRS_ATIVO_CC_CONTROL_COLIG) AS EMPRS_ATIVO_CC_CONTROL_COLIG,sum(EMPRS_PASS_CC_CONTROL_COLIG) AS EMPRS_PASS_CC_CONTROL_COLIG,"
        qry = qry & "sum(EMPRS_ATIVO_OUTRAS_CONTAS) AS EMPRS_ATIVO_OUTRAS_CONTAS,sum(EMPRS_PASS_OUTRAS_CONTAS) AS EMPRS_PASS_OUTRAS_CONTAS,"
        qry = qry & "sum(EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER) AS EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER,sum(EMPRS_PASS_OUTRAS_CTAS_NAO_OPER) AS EMPRS_PASS_OUTRAS_CTAS_NAO_OPER,"
        qry = qry & "sum(EMPRS_ATIVO_OUTRAS_CTAS_OPERAC) AS EMPRS_ATIVO_OUTRAS_CTAS_OPERAC,sum(EMPRS_PASS_OUTRAS_CTAS_OPERAC) AS EMPRS_PASS_OUTRAS_CTAS_OPERAC,"
        qry = qry & "sum(EMPRS_PROV_DEV_DUVIDS) AS EMPRS_PROV_DEV_DUVIDS,sum(EMPRS_ESTOQS) AS EMPRS_ESTOQS,sum(EMPRS_ADTO_FORN) AS EMPRS_ADTO_FORN,"
        qry = qry & "sum(EMPRS_TIT_VAL_MOBIL) AS EMPRS_TIT_VAL_MOBIL,sum(EMPRS_DESP_PG_ANTEC) AS EMPRS_DESP_PG_ANTEC,sum(EMPRS_ATIVO_CIRC) AS EMPRS_ATIVO_CIRC,"
        qry = qry & "sum(EMPRS_REALZVL_A_L_P) AS EMPRS_REALZVL_A_L_P,sum(EMPRS_PART_CONTROL_COLIGS) AS EMPRS_PART_CONTROL_COLIGS,sum(EMPRS_OUTROS_INVEST) AS EMPRS_OUTROS_INVEST,"
        qry = qry & "sum(EMPRS_INVEST) AS EMPRS_INVEST,sum(EMPRS_IMOB_TECN_LIQ) AS EMPRS_IMOB_TECN_LIQ,sum(EMPRS_ATIVO_INTANG) AS EMPRS_ATIVO_INTANG,"
        qry = qry & "sum(EMPRS_ATIVO_PERMAN) AS EMPRS_ATIVO_PERMAN,sum(EMPRS_ATIVO_TOT) AS EMPRS_ATIVO_TOT,sum(EMPRS_FORNS) AS EMPRS_FORNS,"
        qry = qry & "sum(EMPRS_OBRIG_SOC_TRIBUT) AS EMPRS_OBRIG_SOC_TRIBUT,sum(EMPRS_ADTO_CLI) AS EMPRS_ADTO_CLI,sum(EMPRS_DUPLIC_DESCTS) AS EMPRS_DUPLIC_DESCTS,"
        qry = qry & "sum(EMPRS_CAMB) AS EMPRS_CAMB,sum(EMPRS_EMPREST_FINANCS) AS EMPRS_EMPREST_FINANCS,sum(EMPRS_EXIG_A_LP) AS EMPRS_EXIG_A_LP,"
        qry = qry & "sum(EMPRS_RES_EXERC_FUT) AS EMPRS_RES_EXERC_FUT,sum(EMPRS_CAPIT_SOC) AS EMPRS_CAPIT_SOC,sum(EMPRS_RESERV_CAPIT_LUCRO) AS EMPRS_RESERV_CAPIT_LUCRO,"
        qry = qry & "sum(EMPRS_RESERV_REAVAL) AS EMPRS_RESERV_REAVAL,sum(EMPRS_PARTIC_MINOR) AS EMPRS_PARTIC_MINOR,sum(EMPRS_LUCRO_PREJ_ACML) AS EMPRS_LUCRO_PREJ_ACML,"
        qry = qry & "sum(EMPRS_PATR_LIQ) AS EMPRS_PATR_LIQ,sum(EMPRS_PASSIVO_TOTAL) AS EMPRS_PASSIVO_TOTAL,sum(EMPRS_RECT_OPER_LIQ) AS EMPRS_RECT_OPER_LIQ,"
        qry = qry & "sum(EMPRS_CUSTO_PROD_VENDS) AS EMPRS_CUSTO_PROD_VENDS,sum(EMPRS_LUCRO_BRUTO) AS EMPRS_LUCRO_BRUTO,sum(EMPRS_DESP_ADMINS) AS EMPRS_DESP_ADMINS,"
        qry = qry & "sum(EMPRS_DESP_VNDAS) AS EMPRS_DESP_VNDAS,sum(EMPRS_OUTRAS_DESP_REC_OPERAC) AS EMPRS_OUTRAS_DESP_REC_OPERAC,"
        qry = qry & "sum(EMPRS_SALDO_COR_MONET) AS EMPRS_SALDO_COR_MONET,sum(EMPRS_LUCRO_ANTES_RES_FINAN) AS EMPRS_LUCRO_ANTES_RES_FINAN,"
        qry = qry & "sum(EMPRS_RECT_FINANC) AS EMPRS_RECT_FINANC,sum(EMPRS_DESPS_FINANCS) AS EMPRS_DESPS_FINANCS,sum(EMPRS_VAR_CAMBL_LIQ) AS EMPRS_VAR_CAMBL_LIQ,"
        qry = qry & "sum(EMPRS_RECT_DESP_NAO_OPERAC) AS EMPRS_RECT_DESP_NAO_OPERAC,sum(EMPRS_LUCRO_ANTES_EQUIC_PATR) AS EMPRS_LUCRO_ANTES_EQUIC_PATR,"
        qry = qry & "sum(EMPRS_EQUIV_PATRIOM) AS EMPRS_EQUIV_PATRIOM,sum(EMPRS_LUCRO_ANTES_IR) AS EMPRS_LUCRO_ANTES_IR,"
        qry = qry & "sum(EMPRS_IMP_RNDA_CONTRIB_SOC) AS EMPRS_IMP_RNDA_CONTRIB_SOC,sum(EMPRS_PARTIC) AS EMPRS_PARTIC,sum(EMPRS_LUCRO_LIQ) AS EMPRS_LUCRO_LIQ,"
        qry = qry & "sum(EMPRS_CLI) AS EMPRS_CLI,sum(EMPRS_PASSIVO_CIRC) AS EMPRS_PASSIVO_CIRC,sum(EMPRS_RECT_BRUTA) AS EMPRS_RECT_BRUTA,"
        qry = qry & "sum(EMPRS_DEVOL_ABATIM) AS EMPRS_DEVOL_ABATIM,sum(EMPRS_IMPOS_FATRDS) AS EMPRS_IMPOS_FATRDS,sum(EMPRS_DEPREC) AS EMPRS_DEPREC,"
        qry = qry & "sum(EMPRS_EBITDA) AS EMPRS_EBITDA,sum(EMPRS_EBITDA_ROL) AS EMPRS_EBITDA_ROL,sum(EMPRS_VAR_CAMB_LIQ) AS EMPRS_VAR_CAMB_LIQ,"
        qry = qry & "sum(EMPRS_BCO_LIQ_AQ_TERR_EBITDA) AS EMPRS_BCO_LIQ_AQ_TERR_EBITDA,sum(EMPRS_BCO_LIQ_AQ_TERR_PL) AS EMPRS_BCO_LIQ_AQ_TERR_PL,"
        qry = qry & "sum(EMPRS_BCO_LIQ_EQUIV_PATRIOM) AS EMPRS_BCO_LIQ_EQUIV_PATRIOM,sum(EMPRS_BCO_LIQ_MOAGEM) AS EMPRS_BCO_LIQ_MOAGEM,"
        qry = qry & "sum(EMPRS_BCO_LIQ_PATRIM_AV_CRED) AS EMPRS_BCO_LIQ_PATRIM_AV_CRED,sum(EMPRS_BCO_LIQ_PL) AS EMPRS_BCO_LIQ_PL,"
        qry = qry & "sum(EMPRS_BCO_TOTAL_LIQ) AS EMPRS_BCO_TOTAL_LIQ,sum(EMPRS_CP_DERIVAT) AS EMPRS_CP_DERIVAT,sum(EMPRS_LP_DERIVAT) AS EMPRS_LP_DERIVAT,"
        qry = qry & "sum(EMPRS_DIVID_PAGOS) AS EMPRS_DIVID_PAGOS,sum(EMPRS_DIVID_RECEB) AS EMPRS_DIVID_RECEB,sum(EMPRS_EBITDA_MW_CAPACID_INST) AS EMPRS_EBITDA_MW_CAPACID_INST,"
        qry = qry & "sum(EMPRS_ATIVO_DIFER_PL) AS EMPRS_ATIVO_DIFER_PL,sum(EMPRS_BCO_ROL) AS EMPRS_BCO_ROL,sum(EMPRS_AJUST1) AS EMPRS_AJUST1,sum(EMPRS_AJUST2) AS EMPRS_AJUST2,"
        qry = qry & "sum(EMPRS_AJUST3) AS EMPRS_AJUST3,sum(EMPRS_BCO_LIQ_AQ_TERR_EBTIDA_AJ) AS EMPRS_BCO_LIQ_AQ_TERR_EBTIDA_AJ,"
        qry = qry & "sum(EMPRS_BCO_LIQ_EBITDA_AJUST) AS EMPRS_BCO_LIQ_EBITDA_AJUST,sum(EMPRS_EBITDA_AJ_MW_CAPAC_INST) AS EMPRS_EBITDA_AJ_MW_CAPAC_INST,"
        qry = qry & "sum(EMPRS_BCO_AJUST_PL) AS EMPRS_BCO_AJUST_PL,sum(EMPRS_BCO_AJUST_PL_ROL) AS EMPRS_BCO_AJUST_PL_ROL,"
        qry = qry & "sum(PEFIS_MM_PATR_COMPROVADO) AS PEFIS_MM_PATR_COMPROVADO,sum(PEFIS_MM_LIQ) AS PEFIS_MM_LIQ,sum(PEFIS_MM_ATIV_IMBLZ) AS PEFIS_MM_ATIV_IMBLZ,"
        qry = qry & "sum(PEFIS_MM_PARTICIP_EMP) AS PEFIS_MM_PARTICIP_EMP,sum(PEFIS_MM_GADO) AS PEFIS_MM_GADO,sum(PEFIS_MM_OUTRO) AS PEFIS_MM_OUTRO,"
        qry = qry & "sum(PEFIS_MM_DIV_BCRA) AS PEFIS_MM_DIV_BCRA,sum(PEFIS_MM_DIV_AVAIS) AS PEFIS_MM_DIV_AVAIS,sum(PEFIS_MM_PATR_LIQ) AS PEFIS_MM_PATR_LIQ,"
        qry = qry & "sum(PEFIS_FI_PATR_COMPROVADO) AS PEFIS_FI_PATR_COMPROVADO,sum(PEFIS_FI_LIQ) AS PEFIS_FI_LIQ,sum(PEFIS_FI_ATIV_IMBLZ) AS PEFIS_FI_ATIV_IMBLZ,"
        qry = qry & "sum(PEFIS_FI_PARTICIP_EMP) AS PEFIS_FI_PARTICIP_EMP,sum(PEFIS_FI_GADO) AS PEFIS_FI_GADO,sum(PEFIS_FI_OUTRO) AS PEFIS_FI_OUTRO,"
        qry = qry & "sum(PEFIS_FI_DIV_BCRA) AS PEFIS_FI_DIV_BCRA,sum(PEFIS_FI_DIV_AVAIS) AS PEFIS_FI_DIV_AVAIS,sum(PEFIS_DB_PATR_COMPROVADO) AS PEFIS_DB_PATR_COMPROVADO,"
        qry = qry & "sum(PEFIS_DB_LIQ) AS PEFIS_DB_LIQ,sum(PEFIS_DB_ATIV_IMBLZ) AS PEFIS_DB_ATIV_IMBLZ,sum(PEFIS_DB_PARTICIP_EMP) AS PEFIS_DB_PARTICIP_EMP,"
        qry = qry & "sum(PEFIS_DB_GADO) AS PEFIS_DB_GADO,sum(PEFIS_DB_OUTRO) AS PEFIS_DB_OUTRO,sum(PEFIS_DB_DIV_BCRA) AS PEFIS_DB_DIV_BCRA,"
        qry = qry & "sum(PEFIS_DB_DIV_AVAIS) AS PEFIS_DB_DIV_AVAIS,sum(PEFIS_IR_APLIC_FIN) AS PEFIS_IR_APLIC_FIN,sum(PEFIS_IR_QT_ACOES_EMPRS) AS PEFIS_IR_QT_ACOES_EMPRS,"
        qry = qry & "sum(PEFIS_IR_IMOVEIS) AS PEFIS_IR_IMOVEIS,sum(PEFIS_IR_VEICULOS) AS PEFIS_IR_VEICULOS,sum(PEFIS_IR_EMP_TERCEIRO) AS PEFIS_IR_EMP_TERCEIRO,"
        qry = qry & "sum(PEFIS_IR_OUTRO) AS PEFIS_IR_OUTRO,sum(PEFIS_IR_TOTAL_BENS_DIRT) AS PEFIS_IR_TOTAL_BENS_DIRT,sum(PEFIS_IR_DIV_ONUS) AS PEFIS_IR_DIV_ONUS,"
        qry = qry & "sum(PEFIS_IR_DIV_AVAIS) AS PEFIS_IR_DIV_AVAIS,sum(PEFIS_IR_PATR_LIQ) AS PEFIS_IR_PATR_LIQ,sum(PEFIS_ARREC) AS PEFIS_ARREC,"
        qry = qry & "sum(PEFIS_ARDESP) AS PEFIS_ARDESP,sum(PEFIS_ARRESULT) AS PEFIS_ARRESULT,sum(PEFIS_ARBENS_ATIV_RURAL) AS PEFIS_ARBENS_ATIV_RURAL,"
        qry = qry & "sum(PEFIS_ARDIV_VIN_ATIV_RURAL) AS PEFIS_ARDIV_VIN_ATIV_RURAL,sum(SEGUR_DISP) AS SEGUR_DISP,"
        qry = qry & "sum(SEGUR_CRED_OPER_PREVID_COMPL) AS SEGUR_CRED_OPER_PREVID_COMPL,sum(SEGUR_SEGURADORAS) AS SEGUR_SEGURADORAS,sum(SEGUR_IRB) AS SEGUR_IRB,"
        qry = qry & "sum(SEGUR_DESP_COMERC_DIFERD) AS SEGUR_DESP_COMERC_DIFERD,sum(SEGUR_TITULO_VL_MBLRO) AS SEGUR_TITULO_VL_MBLRO,"
        qry = qry & "sum(SEGUR_DESP_PAGTO_ANTCPO) AS SEGUR_DESP_PAGTO_ANTCPO,sum(SEGUR_OUTRA_CONTA_OPER) AS SEGUR_OUTRA_CONTA_OPER,"
        qry = qry & "sum(SEGUR_OUTRA_CONTA_NAO_OPER) AS SEGUR_OUTRA_CONTA_NAO_OPER,sum(SEGUR_ATIV_CIRC) AS SEGUR_ATIV_CIRC,sum(SEGUR_APLIC) AS SEGUR_APLIC,"
        qry = qry & "sum(SEGUR_TITULO_CRED_RECEB) AS SEGUR_TITULO_CRED_RECEB,sum(SEGUR_REALZV_LP) AS SEGUR_REALZV_LP,sum(SEGUR_PART_CTRL_COLGD) AS SEGUR_PART_CTRL_COLGD,"
        qry = qry & "sum(SEGUR_OUTRO_INVTMO) AS SEGUR_OUTRO_INVTMO,sum(SEGUR_INVTMO) AS SEGUR_INVTMO,sum(SEGUR_IMBRO_TECN_LIQ) AS SEGUR_IMBRO_TECN_LIQ,"
        qry = qry & "sum(SEGUR_ATIV_DFRD) AS SEGUR_ATIV_DFRD,sum(SEGUR_ATIV_PERMAN) AS SEGUR_ATIV_PERMAN,sum(SEGUR_ATIV_TOTAL) AS SEGUR_ATIV_TOTAL,"
        qry = qry & "sum(SEGUR_DEB_OPER_PREVID) AS SEGUR_DEB_OPER_PREVID,sum(SEGUR_OBRIG_SOC_TRIB) AS SEGUR_OBRIG_SOC_TRIB,sum(SEGUR_SINIS_LIQ) AS SEGUR_SINIS_LIQ,"
        qry = qry & "sum(SEGUR_EMPREST_FIN) AS SEGUR_EMPREST_FIN,sum(SEGUR_PROV_TECN) AS SEGUR_PROV_TECN,sum(SEGUR_DEPOS_TERC) AS SEGUR_DEPOS_TERC,"
        qry = qry & "sum(SEGUR_CTRL_COLGD) AS SEGUR_CTRL_COLGD,sum(SEGUR_PASV_CIRC) AS SEGUR_PASV_CIRC,sum(SEGUR_OUTRA_CONTA) AS SEGUR_OUTRA_CONTA,"
        qry = qry & "sum(SEGUR_EXIG_LP) AS SEGUR_EXIG_LP,sum(SEGUR_RES_EXERC_FUT) AS SEGUR_RES_EXERC_FUT,sum(SEGUR_CAPITAL_SOC) AS SEGUR_CAPITAL_SOC,"
        qry = qry & "sum(SEGUR_RES_CAPITAL_LCR) AS SEGUR_RES_CAPITAL_LCR,sum(SEGUR_RES_REAVAL) AS SEGUR_RES_REAVAL,sum(SEGUR_PARTICIP_MNTRO) AS SEGUR_PARTICIP_MNTRO,"
        qry = qry & "sum(SEGUR_LCR_PREJ_ACUM) AS SEGUR_LCR_PREJ_ACUM,sum(SEGUR_PATR_LIQ) AS SEGUR_PATR_LIQ,sum(SEGUR_PASV_TOTAL) AS SEGUR_PASV_TOTAL,"
        qry = qry & "sum(SEGUR_RENDA_CONTRIB) AS SEGUR_RENDA_CONTRIB,sum(SEGUR_CONTRIB_RPS) AS SEGUR_CONTRIB_RPS,sum(SEGUR_VAR_PROV_PREMIOS) AS SEGUR_VAR_PROV_PREMIOS,"
        qry = qry & "sum(SEGUR_REC_OPER_LIQ) AS SEGUR_REC_OPER_LIQ,sum(SEGUR_DESP_BENEF_RESGT) AS SEGUR_DESP_BENEF_RESGT,"
        qry = qry & "sum(SEGUR_VAR_PROV_EVENTO_NAO_AVIS) AS SEGUR_VAR_PROV_EVENTO_NAO_AVIS,sum(SEGUR_LCR_BRUTO) AS SEGUR_LCR_BRUTO,"
        qry = qry & "sum(SEGUR_DESP_ADM) AS SEGUR_DESP_ADM,sum(SEGUR_DESP_VDA) AS SEGUR_DESP_VDA,sum(SEGUR_OUTRO_DESP_REC_OPER) AS SEGUR_OUTRO_DESP_REC_OPER,"
        qry = qry & "sum(SEGUR_SALDO_CORREC_MONET) AS SEGUR_SALDO_CORREC_MONET,sum(SEGUR_LCR_ANTES_RES_FIN) AS SEGUR_LCR_ANTES_RES_FIN,sum(SEGUR_RECT_FIN) AS SEGUR_RECT_FIN,"
        qry = qry & "sum(SEGUR_DESP_FIN) AS SEGUR_DESP_FIN,sum(SEGUR_REC_DESP_NAO_OPER) AS SEGUR_REC_DESP_NAO_OPER,"
        qry = qry & "sum(SEGUR_LCR_ANTES_EQUIV_PATRIM) AS SEGUR_LCR_ANTES_EQUIV_PATRIM,sum(SEGUR_EQUIV_PATRIM) AS SEGUR_EQUIV_PATRIM,"
        qry = qry & "sum(SEGUR_LCR_ANTES_IR) AS SEGUR_LCR_ANTES_IR,sum(SEGUR_IR_RENDA_CONTRIB_SOC) AS SEGUR_IR_RENDA_CONTRIB_SOC,sum(SEGUR_PARTICIP) AS SEGUR_PARTICIP,"
        qry = qry & "sum(SEGUR_LCR_LIQ) AS SEGUR_LCR_LIQ,sum(OP_DISPNVL) AS OP_DISPNVL,sum(OP_CRED_A_CP) AS OP_CRED_A_CP,"
        qry = qry & "sum(OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP) AS OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP,sum(OP_ATIVO_INVESTIMENTOS) AS OP_ATIVO_INVESTIMENTOS,"
        qry = qry & "sum(OP_ATIVO_CIRC_ESTOQ) AS OP_ATIVO_CIRC_ESTOQ,sum(OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP) AS OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP,"
        qry = qry & "sum(OP_CRED_A_LP) AS OP_CRED_A_LP,sum(OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP) AS OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP,sum(OP_INVESTIMENTOS) AS OP_INVESTIMENTOS,"
        qry = qry & "sum(OP_ATV_RLZ_ESTOQ) AS OP_ATV_RLZ_ESTOQ,sum(OP_ATV_RLZ_VPD_PAGAS_ANTECIP) AS OP_ATV_RLZ_VPD_PAGAS_ANTECIP,sum(OP_IMOBILIZADO) AS OP_IMOBILIZADO,"
        qry = qry & "sum(OP_INTANGIVEL) AS OP_INTANGIVEL,sum(OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP) AS OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP,"
        qry = qry & "sum(OP_EMPREST_FINAN_CP) AS OP_EMPREST_FINAN_CP,sum(OP_FORN_CTAS_PG_CP) AS OP_FORN_CTAS_PG_CP,sum(OP_OBRIG_FISCAIS_CP) AS OP_OBRIG_FISCAIS_CP,"
        qry = qry & "sum(OP_OBRIG_REPART) AS OP_OBRIG_REPART,sum(OP_PROV_CP) AS OP_PROV_CP,sum(OP_DEMAIS_OBRIG_CP) AS OP_DEMAIS_OBRIG_CP,"
        qry = qry & "sum(OP_PASS_N_CIRC_OB_TRB_PREV_AS_CP) AS OP_PASS_N_CIRC_OB_TRB_PREV_AS_CP,sum(OP_EMPREST_FINANC_LP) AS OP_EMPREST_FINANC_LP,"
        qry = qry & "sum(OP_FORNECEDORES_LP) AS OP_FORNECEDORES_LP,sum(OP_PREVISOES_LP) AS OP_PREVISOES_LP,sum(OP_DEMAIS_OBRIG_LP) AS OP_DEMAIS_OBRIG_LP,"
        qry = qry & "sum(OP_PATRIMONIO_LP) AS OP_PATRIMONIO_LP,sum(OP_TRIBUTARIAS) AS OP_TRIBUTARIAS,sum(OP_CONTRIBUICOES) AS OP_CONTRIBUICOES,"
        qry = qry & "sum(OP_TRANSF_CORRENTES) AS OP_TRANSF_CORRENTES,sum(OP_PATRIMONIAIS) AS OP_PATRIMONIAIS,sum(OP_OUTRAS_RECEITAS_CORRENTES) AS OP_OUTRAS_RECEITAS_CORRENTES,"
        qry = qry & "sum(OP_DEDUCOES) AS OP_DEDUCOES,sum(OP_PESSOAL_ENCARGOS_SOCIAIS) AS OP_PESSOAL_ENCARGOS_SOCIAIS,sum(OP_JUROS_ENCARGOS_DIVIDAS) AS OP_JUROS_ENCARGOS_DIVIDAS,"
        qry = qry & "sum(OP_TRANSFERENCIAS_CORRENTES) AS OP_TRANSFERENCIAS_CORRENTES,sum(OP_OUTRAS_DESPESAS_CORRENTES) AS OP_OUTRAS_DESPESAS_CORRENTES,"
        qry = qry & "sum(OP_OPERACOES_CREDITO) AS OP_OPERACOES_CREDITO,sum(OP_ALIENACAO_BENS) AS OP_ALIENACAO_BENS,sum(OP_TRANSFERENCIA_CAPITAL) AS OP_TRANSFERENCIA_CAPITAL,"
        qry = qry & "sum(OP_RECEITA_CAPITAL_OUTRAS) AS OP_RECEITA_CAPITAL_OUTRAS,sum(OP_INVERSOES_FINANCEIRAS) AS OP_INVERSOES_FINANCEIRAS,"
        qry = qry & "sum(OP_AMORTIZACAO_DIVIDA) AS OP_AMORTIZACAO_DIVIDA,sum(OP_OUTRAS_DESPESAS_CAPITAL) AS OP_OUTRAS_DESPESAS_CAPITAL,"
        qry = qry & "sum(OP_OUTRAS_RECEITAS_DESPESAS) AS OP_OUTRAS_RECEITAS_DESPESAS,sum(OP_RESERVAS_CONTINGENCIAS) AS OP_RESERVAS_CONTINGENCIAS,"
        qry = qry & "sum(OP_ORCADO_TRIBUTARIAS) AS OP_ORCADO_TRIBUTARIAS,sum(OP_ORCADO_CONTRIBUICOES) AS OP_ORCADO_CONTRIBUICOES,"
        qry = qry & "sum(OP_ORCADO_TRANSF_CORRENTES) AS OP_ORCADO_TRANSF_CORRENTES,sum(OP_ORCADO_PATRIMONIAIS) AS OP_ORCADO_PATRIMONIAIS,"
        qry = qry & "sum(OP_ORCADO_OUTRAS_RECT_CORREN) AS OP_ORCADO_OUTRAS_RECT_CORREN,sum(OP_ORCADO_DEDUCOES) AS OP_ORCADO_DEDUCOES,"
        qry = qry & "sum(OP_ORCADO_PESS_ENCG_DIV) AS OP_ORCADO_PESS_ENCG_DIV,sum(OP_ORCADO_JUROS_ENCARG_DIV) AS OP_ORCADO_JUROS_ENCARG_DIV,"
        qry = qry & "sum(OP_ORCADO_TRANSF_CORR) AS OP_ORCADO_TRANSF_CORR,sum(OP_ORCADO_OUTRAS_DESP_CORR) AS OP_ORCADO_OUTRAS_DESP_CORR,"
        qry = qry & "sum(OP_ORCADO_OPER_CREDITO) AS OP_ORCADO_OPER_CREDITO,sum(OP_ORCADO_ALIENACAO_BENS) AS OP_ORCADO_ALIENACAO_BENS,"
        qry = qry & "sum(OP_ORCADO_TRANSF_CAPITAL) AS OP_ORCADO_TRANSF_CAPITAL,sum(OP_ORCADO_RECT_CAPITAL_OUTRAS) AS OP_ORCADO_RECT_CAPITAL_OUTRAS,"
        qry = qry & "sum(OP_ORCADO_INVESTIMENTOS) AS OP_ORCADO_INVESTIMENTOS,sum(OP_ORCADO_INVERSOES_FIN) AS OP_ORCADO_INVERSOES_FIN,"
        qry = qry & "sum(OP_ORCADO_AMORT_DIVIDA) AS OP_ORCADO_AMORT_DIVIDA,sum(OP_ORCADO_OUTRAS_DESP_CAPITAL) AS OP_ORCADO_OUTRAS_DESP_CAPITAL,"
        qry = qry & "sum(OP_ORCADO_OUTRAS_RECT_DESPESAS) AS OP_ORCADO_OUTRAS_RECT_DESPESAS,"
        qry = qry & "sum(OP_ORCADO_RESERVAS_CONTI) AS OP_ORCADO_RESERVAS_CONTI from lb_plani.fato_balanco_aux "
        qry = qry & "where SUBSTR(dt_exerc,7,10) = '" & dt & "') where cd_grp = '" & cod_grupo & "'"
    
        somaBalanco = qry
        
End Function


Function updateMoeda(cotacao As Double, per As String, cd_cli As String) As String



qry = "UPDATE LB_PLANI.FATO_BALANCO_AUX "
qry = qry & "SET BCO_ATIVO_CART_CAMB  = (BCO_ATIVO_CART_CAMB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_OUTROS_CREDS  = (BCO_ATIVO_OUTROS_CREDS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATV_N_CIRC_PART_CTRL_COLIG  = (BCO_ATV_N_CIRC_PART_CTRL_COLIG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATV_N_CIRC_OUTROS_INVEST  = (BCO_ATV_N_CIRC_OUTROS_INVEST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATV_N_CIRC_INVEST  = (BCO_ATV_N_CIRC_INVEST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATV_N_CIRC_IMOB_TEC_LIQ  = (BCO_ATV_N_CIRC_IMOB_TEC_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATV_N_CIRC_ATV_INTANG  = (BCO_ATV_N_CIRC_ATV_INTANG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATV_N_CIRC_ATV_PERMAN  = (BCO_ATV_N_CIRC_ATV_PERMAN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATV_N_CIRC_ATV_TOTAL  = (BCO_ATV_N_CIRC_ATV_TOTAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_OP_ARREND_MERCATL  = (BCO_ATIVO_OP_ARREND_MERCATL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_DISP  = (BCO_ATIVO_DISP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_PDD_ARREND_MERCATL  = (BCO_ATIVO_PDD_ARREND_MERCATL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_CIRC  = (BCO_PASS_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_N_CIRC_PART_MINOR  = (BCO_PASS_N_CIRC_PART_MINOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_N_CIRC_AJUST_VLR_MERC  = (BCO_PASS_N_CIRC_AJUST_VLR_MERC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_N_CIRC_LCR_PREJ_ACML  = (BCO_PASS_N_CIRC_LCR_PREJ_ACML *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_N_CIRC_PATRIM_LIQ  = (BCO_PASS_N_CIRC_PATRIM_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_N_CIRC_PASS_TOTAL  = (BCO_PASS_N_CIRC_PASS_TOTAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_DEPOS_APRAZO  = (BCO_PASS_DEPOS_APRAZO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_CIRC_REPASS_PAIS  = (BCO_PASS_CIRC_REPASS_PAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_LUCRO_LIQ  = (BCO_DRE_LUCRO_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_EMPR_CESS_REPASS  = (BCO_DRE_EMPR_CESS_REPASS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CIVEIS_CONTING_NAO_PROVS  = (BCO_CIVEIS_CONTING_NAO_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_TRABLSTAS_CONTING_NAO_PROVS  = (BCO_TRABLSTAS_CONTING_NAO_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_FISCAIS_CONTING_NAO_PROVS  = (BCO_FISCAIS_CONTING_NAO_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_TOTAL_CONTING_NAO_PROVS  = (BCO_TOTAL_CONTING_NAO_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TRABLSTAS_CONTING_PROVS  = (BCO_TRABLSTAS_CONTING_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_FISCAIS_CONTING_PROVS  = (BCO_FISCAIS_CONTING_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOTAL_CONTING_PROVS  = (BCO_TOTAL_CONTING_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CIVEIS_DEPOS_JUDC  = (BCO_CIVEIS_DEPOS_JUDC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TRABLSTAS_DEPOS_JUDC  = (BCO_TRABLSTAS_DEPOS_JUDC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_FISCAIS_DEPOS_JUDC  = (BCO_FISCAIS_DEPOS_JUDC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOTAL_DEPOS_JUDC  = (BCO_TOTAL_DEPOS_JUDC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CIVEIS_CONTING_PROVS  = (BCO_CIVEIS_CONTING_PROVS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_P_NEGOC_VALOR_CUSTO  = (BCO_P_NEGOC_VALOR_CUSTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_P_NEGOC_VALOR_CONTAB  = (BCO_P_NEGOC_VALOR_CONTAB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_P_NEGOC_MTM  = (BCO_P_NEGOC_MTM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DISP_VENDA_VALOR_CUSTO  = (BCO_DISP_VENDA_VALOR_CUSTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DISP_VENDA_VALOR_CONTAB  = (BCO_DISP_VENDA_VALOR_CONTAB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DISP_VENDA_MTM  = (BCO_DISP_VENDA_MTM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_MTDOS_VCTO_VALOR_CUSTO  = (BCO_MTDOS_VCTO_VALOR_CUSTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_OUTRAS_REC_INTERM  = (BCO_DRE_OUTRAS_REC_INTERM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_MTDOS_VCTO_VALOR_CONTAB  = (BCO_MTDOS_VCTO_VALOR_CONTAB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_MTDOS_VCTO_MTM  = (BCO_MTDOS_VCTO_MTM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_INSTR_FINANC_DERIV_VLR_CUSTO  = (BCO_INSTR_FINANC_DERIV_VLR_CUSTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_INSTR_FIN_DERIV_VLR_CONTAB  = (BCO_INSTR_FIN_DERIV_VLR_CONTAB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_INSTR_FINANC_DERIV_MTM  = (BCO_INSTR_FINANC_DERIV_MTM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_AA  = (BCO_AA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOTAL_CART  = (BCO_TOTAL_CART *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_D_H  = (BCO_D_H *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PDD_EXIG  = (BCO_PDD_EXIG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PDD_CONST  = (BCO_PDD_CONST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_VENCD  = (BCO_VENCD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_VENCD_90D  = (BCO_VENCD_90D *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_A  = (BCO_A *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_B  = (BCO_B *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_C  = (BCO_C *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_D  = (BCO_D *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_E  = (BCO_E *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_F  = (BCO_F *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_G  = (BCO_G *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_H  = (BCO_H *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_BXDOS_SECTZDOS  = (BCO_BXDOS_SECTZDOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_IND_BASILEIA_BR  = (BCO_IND_BASILEIA_BR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_AVAIS_FIANCAS_PRESTDOS  = (BCO_AVAIS_FIANCAS_PRESTDOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_REPASS_PAIS  = (BCO_PASS_REPASS_PAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_AG  = (BCO_AG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_FUNC  = (BCO_FUNC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_FNDS_ADMN  = (BCO_FNDS_ADMN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DEPOS_JUDC  = (BCO_DEPOS_JUDC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_BNDU  = (BCO_BNDU *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PART_CONTRLDS_CLGDS  = (BCO_PART_CONTRLDS_CLGDS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_INTANG  = (BCO_ATIVO_INTANG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CDI_LIQDZ_DIA  = (BCO_CDI_LIQDZ_DIA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TVM_VINC_PREST_GAR_NEG  = (BCO_TVM_VINC_PREST_GAR_NEG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_TVM_BAIXA_LIQDZ  = (BCO_TVM_BAIXA_LIQDZ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_INSTRM_FIN_DERIV_PASS_NEG  = (BCO_INSTRM_FIN_DERIV_PASS_NEG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CAPT_MERC_ABER_NEG  = (BCO_CAPT_MERC_ABER_NEG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CAIXA_DISPNVL  = (BCO_CAIXA_DISPNVL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CAIXA_DISPNVL_PL  = (BCO_CAIXA_DISPNVL_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CAIXA_DISPNVL_CART_CRED  = (BCO_CAIXA_DISPNVL_CART_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_OPER_CRED_ARREND_MERCTL  = (BCO_OPER_CRED_ARREND_MERCTL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PDD_NEG  = (BCO_PDD_NEG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_IND_BASILEIA  = (BCO_IND_BASILEIA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CGP_AJUST_ATV_CAPT_MERC_ABER  = (BCO_CGP_AJUST_ATV_CAPT_MERC_ABER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ALVCGEM_PASSVA  = (BCO_ALVCGEM_PASSVA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ALVCGEM_CRED  = (BCO_ALVCGEM_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ALVCGEM_OPER  = (BCO_ALVCGEM_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CAPT_GIRO_PROP  = (BCO_CAPT_GIRO_PROP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CAPT_GIRO_PROP_AJUST  = (BCO_CAPT_GIRO_PROP_AJUST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CGP_AJUST_PL  = (BCO_CGP_AJUST_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_SALDO_INICIAL  = (BCO_SALDO_INICIAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CONST  = (BCO_CONST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_REVERSAO  = (BCO_REVERSAO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_BAIXAS  = (BCO_BAIXAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_SALDO_FINAL  = (BCO_SALDO_FINAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_RENEG_FLUXO  = (BCO_RENEG_FLUXO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_RECUP  = (BCO_RECUP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_LCA  = (BCO_LCA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_LCI  = (BCO_LCI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_LF  = (BCO_LF *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOTAL  = (BCO_TOTAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PATR_LIQ  = (BCO_PATR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_CIRC_EMPREST_EXTERIOR  = (BCO_PASS_CIRC_EMPREST_EXTERIOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_CIRC_REPASS_EXTERIOR  = (BCO_PASS_CIRC_REPASS_EXTERIOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_CIRC_OUTRAS_CONTAS  = (BCO_PASS_CIRC_OUTRAS_CONTAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PARTS_RELAC  = (BCO_PARTS_RELAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_REC_INTERFIN  = (BCO_REC_INTERFIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_RESULT_NAO_OPERAC  = (BCO_RESULT_NAO_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_LUCRO_ANTES_IR  = (BCO_LUCRO_ANTES_IR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_IR_CS  = (BCO_IR_CS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_LUCRO_LIQ  = (BCO_LUCRO_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_NIM  = (BCO_NIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_EFICIENCY_RATIO  = (BCO_EFICIENCY_RATIO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ROAE  = (BCO_ROAE *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ROAA  = (BCO_ROAA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DESP_INT_FINANC  = (BCO_DESP_INT_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_RES_BRTO_INTERM  = (BCO_RES_BRTO_INTERM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_OUTRAS_REC_DESP_OPERAC  = (BCO_OUTRAS_REC_DESP_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_RES_OPER  = (BCO_RES_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_BASILEIA_TIER_I  = (BCO_BASILEIA_TIER_I *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DPGE_I  = (BCO_DPGE_I *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DPGE_II  = (BCO_DPGE_II *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CRED_TRIB  = (BCO_CRED_TRIB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_OUTROS_PASS  = (BCO_OUTROS_PASS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_APRAZO  = (BCO_APRAZO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_CDI  = (BCO_ATIVO_CDI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_TITULO_MERC_ABERT  = (BCO_ATIVO_TITULO_MERC_ABERT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_TVM  = (BCO_ATIVO_TVM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_OPERAC_CRED  = (BCO_ATIVO_OPERAC_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_PDD  = (BCO_ATIVO_PDD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_DESP_ANTEC  = (BCO_ATIVO_DESP_ANTEC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_CIRC  = (BCO_ATIVO_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_CIRC_TVM  = (BCO_ATIVO_CIRC_TVM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_CIRC_OPERAC_CRED  = (BCO_ATIVO_CIRC_OPERAC_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_CIRC_PDD_OP_CRED  = (BCO_ATIVO_CIRC_PDD_OP_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_CIRC_OP_ARREND_MERC  = (BCO_ATIVO_CIRC_OP_ARREND_MERC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATIVO_CIRC_PDD_OP_ARR_MERC  = (BCO_ATIVO_CIRC_PDD_OP_ARR_MERC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_ATIVO_CIRC_OUTROS_CRED  = (BCO_ATIVO_CIRC_OUTROS_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_ATV_N_CIRC  = (BCO_ATV_N_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_DEPOS_AVISTA  = (BCO_PASS_DEPOS_AVISTA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_POUPANCA  = (BCO_PASS_POUPANCA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_INTERFIN  = (BCO_INTERFIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_DEPOS_INTERFINAN  = (BCO_PASS_DEPOS_INTERFINAN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CAPT_MERC_ABER  = (BCO_CAPT_MERC_ABER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_CAPT_MERC_ABERT  = (BCO_PASS_CAPT_MERC_ABERT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_CIRC_EMPREST_PAIS  = (BCO_PASS_CIRC_EMPREST_PAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_OUTRAS_CONTAS  = (BCO_OUTRAS_CONTAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DEPOS  = (BCO_DEPOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_DEPOS  = (BCO_PASS_DEPOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_EMPREST_PAIS  = (BCO_PASS_EMPREST_PAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_EMPREST_EXTERIOR  = (BCO_PASS_EMPREST_EXTERIOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_REPASS_EXTERIOR  = (BCO_PASS_REPASS_EXTERIOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_OUTRAS_CONTAS  = (BCO_PASS_OUTRAS_CONTAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_N_CIRC  = (BCO_PASS_N_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PASS_N_CIRC_CAPIT_SOC  = (BCO_PASS_N_CIRC_CAPIT_SOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_N_CIRC_RESERV_CAPT  = (BCO_PASS_N_CIRC_RESERV_CAPT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_REC_INTERM_FINANC  = (BCO_DRE_REC_INTERM_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_TVM  = (BCO_DRE_TVM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_CAPT_MERC  = (BCO_DRE_CAPT_MERC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_OUTRAS_DESP_INTERM  = (BCO_DRE_OUTRAS_DESP_INTERM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_DESP_INTERM_FINANC  = (BCO_DRE_DESP_INTERM_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_RES_BRUTO_INTERM  = (BCO_DRE_RES_BRUTO_INTERM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_CONST_PDD  = (BCO_DRE_CONST_PDD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_RES_INTERM_APOS_PDD  = (BCO_DRE_RES_INTERM_APOS_PDD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_RECT_PREST_SERV  = (BCO_DRE_RECT_PREST_SERV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_CUSTO_OPERAC  = (BCO_DRE_CUSTO_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_DESP_TRIBUT  = (BCO_DRE_DESP_TRIBUT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_OUTRAS_RECT_DESP_OPERAC  = (BCO_DRE_OUTRAS_RECT_DESP_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_RES_OPERAC  = (BCO_DRE_RES_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_EQUIV_PATRIM  = (BCO_DRE_EQUIV_PATRIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_RES_APOS_EQUIV_PATRIM  = (BCO_DRE_RES_APOS_EQUIV_PATRIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_RECT_DESP_N_OPERAC  = (BCO_DRE_RECT_DESP_N_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_LUCRO_ANTES_IR  = (BCO_DRE_LUCRO_ANTES_IR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DRE_IMPST_RENDA_CTRL_SOC  = (BCO_DRE_IMPST_RENDA_CTRL_SOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_PART  = (BCO_DRE_PART *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CDI  = (BCO_CDI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_CART_CAMB  = (BCO_CART_CAMB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DEPOS_INTERFIN  = (BCO_DEPOS_INTERFIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DEPOS_APRAZO  = (BCO_DEPOS_APRAZO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_DEPOS_AVISTA  = (BCO_DEPOS_AVISTA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DESP_ANTPDAS  = (BCO_DESP_ANTPDAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DISPS  = (BCO_DISPS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_EMPREST_EXTRIOR  = (BCO_EMPREST_EXTRIOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_EMPREST_PAIS  = (BCO_EMPREST_PAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_OUTROS  = (BCO_OUTROS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_OUTROS_CRED  = (BCO_OUTROS_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PASS_CART_CAMB  = (BCO_PASS_CART_CAMB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_POUPANCA  = (BCO_POUPANCA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_REPAS_EXTRIOR  = (BCO_REPAS_EXTRIOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_REPAS_PAIS  = (BCO_REPAS_PAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TITULO_MERC_ABER  = (BCO_TITULO_MERC_ABER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_LFSN  = (BCO_LFSN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_LETRA_CMBIO  = (BCO_LETRA_CMBIO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DRE_OPERAC_CRED  = (BCO_DRE_OPERAC_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_TVM_CARACT_CRED_NEG  = (BCO_TVM_CARACT_CRED_NEG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TVM  = (BCO_TVM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DIV_SUBORD  = (BCO_DIV_SUBORD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PDD_AVAIS_FIANCAS  = (BCO_PDD_AVAIS_FIANCAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PDD_CARACT_CRED  = (BCO_PDD_CARACT_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PDD_CART_EXPAND  = (BCO_PDD_CART_EXPAND *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TVM_CARACT_CRED  = (BCO_TVM_CARACT_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_RENEG_ESTOQ  = (BCO_RENEG_ESTOQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DISP_VENDA_PROV_P_DESV  = (BCO_DISP_VENDA_PROV_P_DESV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_INSTR_FIN_DERIV_PROV_P_DESV  = (BCO_INSTR_FIN_DERIV_PROV_P_DESV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_MTDOS_VCTO_PROV_P_DESV  = (BCO_MTDOS_VCTO_PROV_P_DESV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_P_NEGOC_PROV_P_DESV  = (BCO_P_NEGOC_PROV_P_DESV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PERC_DISP_VENDA  = (BCO_PERC_DISP_VENDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PERC_INSTR_FINANC  = (BCO_PERC_INSTR_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PERC_MTDOS_VCTO  = (BCO_PERC_MTDOS_VCTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_PERC_P_NEGOC  = (BCO_PERC_P_NEGOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOT_VALOR_CUSTO  = (BCO_TOT_VALOR_CUSTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_TOT_VALOR_CONTAB  = (BCO_TOT_VALOR_CONTAB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOT_MTM  = (BCO_TOT_MTM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_TOT_PERC  = (BCO_TOT_PERC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_TOT_PROV_P_DESV  = (BCO_TOT_PROV_P_DESV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CONST_PDD  = (BCO_CONST_PDD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CUSTO_OPERAC  = (BCO_CUSTO_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_EQUIV_PATRIM  = (BCO_EQUIV_PATRIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_INSTRM_FIN_DERIV  = (BCO_INSTRM_FIN_DERIV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_CREDORES_CRED_C_OBRIG  = (BCO_CREDORES_CRED_C_OBRIG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_PARTIC  = (BCO_PARTIC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_REC_PREST_SERV  = (BCO_REC_PREST_SERV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_DESP_TRIB  = (BCO_DESP_TRIB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "BCO_EFICIENCY_RATIO_AJUS_CLI  = (BCO_EFICIENCY_RATIO_AJUS_CLI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),BCO_NIM_AJUST_CLI  = (BCO_NIM_AJUST_CLI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_DISPS  = (EMPRS_DISPS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DESP_ANTECIP  = (EMPRS_DESP_ANTECIP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_OUTROS_OPERAC  = (EMPRS_OUTROS_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_FINANC_RECEB_CP  = (EMPRS_FINANC_RECEB_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ROL_MENSAL  = (EMPRS_ROL_MENSAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_GER_CXA_OPERAC  = (EMPRS_GER_CXA_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_DESP_FINANC  = (EMPRS_DESP_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_RECTS_FINANCS  = (EMPRS_RECTS_FINANCS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_GER_CXA_APOS_RESULT_FIN  = (EMPRS_GER_CXA_APOS_RESULT_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_INVEST_IMOB_DIFERIDO  = (EMPRS_INVEST_IMOB_DIFERIDO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_INVEST_CONTROL_COLIG  = (EMPRS_INVEST_CONTROL_COLIG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_RESULT_EXERC_FUT  = (EMPRS_RESULT_EXERC_FUT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_PATRIM_LIQ  = (EMPRS_PATRIM_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_MINORITARIO  = (EMPRS_MINORITARIO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ROL_ANUALZDO  = (EMPRS_ROL_ANUALZDO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_RESULT_NAO_OPERAC  = (EMPRS_RESULT_NAO_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_IR_CS  = (EMPRS_IR_CS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EMPRES_PARTES_RELAC  = (EMPRS_EMPRES_PARTES_RELAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_OTRS_ATV_NAO_OPERAC_CP_LP  = (EMPRS_OTRS_ATV_NAO_OPERAC_CP_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_OTRS_PASS_N_OPERAC_CP_LP  = (EMPRS_OTRS_PASS_N_OPERAC_CP_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_VAR_DIVID_BANCAR_LIQ  = (EMPRS_VAR_DIVID_BANCAR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_CP  = (EMPRS_BCO_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DISP_APLIC_FINANC  = (EMPRS_DISP_APLIC_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_CURTO_PL  = (EMPRS_BCO_CURTO_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LP  = (EMPRS_BCO_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EBIT  = (EMPRS_EBIT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_APLIC_FINANC_LP  = (EMPRS_APLIC_FINANC_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LP_LIQ  = (EMPRS_BCO_LP_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LIQ  = (EMPRS_BCO_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_LIQ_ROL  = (EMPRS_BCO_LIQ_ROL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LIQ_EBTIDA_ANUAL  = (EMPRS_BCO_LIQ_EBTIDA_ANUAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_CCL  = (EMPRS_CCL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_VAR_CCL  = (EMPRS_VAR_CCL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_CGP  = (EMPRS_CGP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_VAR_CGP  = (EMPRS_VAR_CGP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_MEIO_CIRC  = (EMPRS_MEIO_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_NECESS_CAPIT_GIRO  = (EMPRS_NECESS_CAPIT_GIRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_SERV_DIVIDA  = (EMPRS_SERV_DIVIDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_CALC_ICSD  = (EMPRS_CALC_ICSD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_RECEB  = (EMPRS_RECEB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_P_M_ESTOQ  = (EMPRS_P_M_ESTOQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_P_M_PAGAM  = (EMPRS_P_M_PAGAM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_FINANC_CONCED_CP  = (EMPRS_FINANC_CONCED_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EQUIV_PATRIM_PL  = (EMPRS_EQUIV_PATRIM_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_ATIVO_CC_CONTROL_COLIG  = (EMPRS_ATIVO_CC_CONTROL_COLIG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PASS_CC_CONTROL_COLIG  = (EMPRS_PASS_CC_CONTROL_COLIG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_ATIVO_OUTRAS_CONTAS  = (EMPRS_ATIVO_OUTRAS_CONTAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PASS_OUTRAS_CONTAS  = (EMPRS_PASS_OUTRAS_CONTAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER  = (EMPRS_ATIVO_OUTRAS_CTAS_NAO_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PASS_OUTRAS_CTAS_NAO_OPER  = (EMPRS_PASS_OUTRAS_CTAS_NAO_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_ATIVO_OUTRAS_CTAS_OPERAC  = (EMPRS_ATIVO_OUTRAS_CTAS_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PASS_OUTRAS_CTAS_OPERAC  = (EMPRS_PASS_OUTRAS_CTAS_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_PROV_DEV_DUVIDS  = (EMPRS_PROV_DEV_DUVIDS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ESTOQS  = (EMPRS_ESTOQS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ADTO_FORN  = (EMPRS_ADTO_FORN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_TIT_VAL_MOBIL  = (EMPRS_TIT_VAL_MOBIL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DESP_PG_ANTEC  = (EMPRS_DESP_PG_ANTEC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ATIVO_CIRC  = (EMPRS_ATIVO_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_REALZVL_A_L_P  = (EMPRS_REALZVL_A_L_P *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PART_CONTROL_COLIGS  = (EMPRS_PART_CONTROL_COLIGS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_OUTROS_INVEST  = (EMPRS_OUTROS_INVEST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_INVEST  = (EMPRS_INVEST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_IMOB_TECN_LIQ  = (EMPRS_IMOB_TECN_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ATIVO_INTANG  = (EMPRS_ATIVO_INTANG *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_ATIVO_PERMAN  = (EMPRS_ATIVO_PERMAN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ATIVO_TOT  = (EMPRS_ATIVO_TOT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_FORNS  = (EMPRS_FORNS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_OBRIG_SOC_TRIBUT  = (EMPRS_OBRIG_SOC_TRIBUT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_ADTO_CLI  = (EMPRS_ADTO_CLI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DUPLIC_DESCTS  = (EMPRS_DUPLIC_DESCTS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_CAMB  = (EMPRS_CAMB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EMPREST_FINANCS  = (EMPRS_EMPREST_FINANCS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EXIG_A_LP  = (EMPRS_EXIG_A_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_RES_EXERC_FUT  = (EMPRS_RES_EXERC_FUT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_CAPIT_SOC  = (EMPRS_CAPIT_SOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_RESERV_CAPIT_LUCRO  = (EMPRS_RESERV_CAPIT_LUCRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_RESERV_REAVAL  = (EMPRS_RESERV_REAVAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PARTIC_MINOR  = (EMPRS_PARTIC_MINOR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LUCRO_PREJ_ACML  = (EMPRS_LUCRO_PREJ_ACML *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_PATR_LIQ  = (EMPRS_PATR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PASSIVO_TOTAL  = (EMPRS_PASSIVO_TOTAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_RECT_OPER_LIQ  = (EMPRS_RECT_OPER_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_CUSTO_PROD_VENDS  = (EMPRS_CUSTO_PROD_VENDS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LUCRO_BRUTO  = (EMPRS_LUCRO_BRUTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DESP_ADMINS  = (EMPRS_DESP_ADMINS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_DESP_VNDAS  = (EMPRS_DESP_VNDAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_OUTRAS_DESP_REC_OPERAC  = (EMPRS_OUTRAS_DESP_REC_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_SALDO_COR_MONET  = (EMPRS_SALDO_COR_MONET *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LUCRO_ANTES_RES_FINAN  = (EMPRS_LUCRO_ANTES_RES_FINAN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_RECT_FINANC  = (EMPRS_RECT_FINANC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DESPS_FINANCS  = (EMPRS_DESPS_FINANCS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_VAR_CAMBL_LIQ  = (EMPRS_VAR_CAMBL_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_RECT_DESP_NAO_OPERAC  = (EMPRS_RECT_DESP_NAO_OPERAC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LUCRO_ANTES_EQUIC_PATR  = (EMPRS_LUCRO_ANTES_EQUIC_PATR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_EQUIV_PATRIOM  = (EMPRS_EQUIV_PATRIOM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LUCRO_ANTES_IR  = (EMPRS_LUCRO_ANTES_IR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_IMP_RNDA_CONTRIB_SOC  = (EMPRS_IMP_RNDA_CONTRIB_SOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PARTIC  = (EMPRS_PARTIC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LUCRO_LIQ  = (EMPRS_LUCRO_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_CLI  = (EMPRS_CLI *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_PASSIVO_CIRC  = (EMPRS_PASSIVO_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_RECT_BRUTA  = (EMPRS_RECT_BRUTA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_DEVOL_ABATIM  = (EMPRS_DEVOL_ABATIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_IMPOS_FATRDS  = (EMPRS_IMPOS_FATRDS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DEPREC  = (EMPRS_DEPREC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_EBITDA  = (EMPRS_EBITDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EBITDA_ROL  = (EMPRS_EBITDA_ROL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_VAR_CAMB_LIQ  = (EMPRS_VAR_CAMB_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_LIQ_AQ_TERR_EBITDA  = (EMPRS_BCO_LIQ_AQ_TERR_EBITDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LIQ_AQ_TERR_PL  = (EMPRS_BCO_LIQ_AQ_TERR_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_LIQ_EQUIV_PATRIOM  = (EMPRS_BCO_LIQ_EQUIV_PATRIOM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LIQ_MOAGEM  = (EMPRS_BCO_LIQ_MOAGEM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_LIQ_PATRIM_AV_CRED  = (EMPRS_BCO_LIQ_PATRIM_AV_CRED *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LIQ_PL  = (EMPRS_BCO_LIQ_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_TOTAL_LIQ  = (EMPRS_BCO_TOTAL_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_CP_DERIVAT  = (EMPRS_CP_DERIVAT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_LP_DERIVAT  = (EMPRS_LP_DERIVAT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_DIVID_PAGOS  = (EMPRS_DIVID_PAGOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_DIVID_RECEB  = (EMPRS_DIVID_RECEB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EBITDA_MW_CAPACID_INST  = (EMPRS_EBITDA_MW_CAPACID_INST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_ATIVO_DIFER_PL  = (EMPRS_ATIVO_DIFER_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_ROL  = (EMPRS_BCO_ROL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_AJUST1  = (EMPRS_AJUST1 *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_AJUST2  = (EMPRS_AJUST2 *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_AJUST3  = (EMPRS_AJUST3 *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_LIQ_AQ_TERR_EBTIDA_AJ  = (EMPRS_BCO_LIQ_AQ_TERR_EBTIDA_AJ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_LIQ_EBITDA_AJUST  = (EMPRS_BCO_LIQ_EBITDA_AJUST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_EBITDA_AJ_MW_CAPAC_INST  = (EMPRS_EBITDA_AJ_MW_CAPAC_INST *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "EMPRS_BCO_AJUST_PL  = (EMPRS_BCO_AJUST_PL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),EMPRS_BCO_AJUST_PL_ROL  = (EMPRS_BCO_AJUST_PL_ROL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_MM_PATR_COMPROVADO  = (PEFIS_MM_PATR_COMPROVADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_MM_LIQ  = (PEFIS_MM_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_MM_ATIV_IMBLZ  = (PEFIS_MM_ATIV_IMBLZ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_MM_PARTICIP_EMP  = (PEFIS_MM_PARTICIP_EMP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_MM_GADO  = (PEFIS_MM_GADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_MM_OUTRO  = (PEFIS_MM_OUTRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_MM_DIV_BCRA  = (PEFIS_MM_DIV_BCRA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_MM_DIV_AVAIS  = (PEFIS_MM_DIV_AVAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_MM_PATR_LIQ  = (PEFIS_MM_PATR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_FI_PATR_COMPROVADO  = (PEFIS_FI_PATR_COMPROVADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_FI_LIQ  = (PEFIS_FI_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_FI_ATIV_IMBLZ  = (PEFIS_FI_ATIV_IMBLZ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_FI_PARTICIP_EMP  = (PEFIS_FI_PARTICIP_EMP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_FI_GADO  = (PEFIS_FI_GADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_FI_OUTRO  = (PEFIS_FI_OUTRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_FI_DIV_BCRA  = (PEFIS_FI_DIV_BCRA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_FI_DIV_AVAIS  = (PEFIS_FI_DIV_AVAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_DB_PATR_COMPROVADO  = (PEFIS_DB_PATR_COMPROVADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_DB_LIQ  = (PEFIS_DB_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_DB_ATIV_IMBLZ  = (PEFIS_DB_ATIV_IMBLZ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_DB_PARTICIP_EMP  = (PEFIS_DB_PARTICIP_EMP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_DB_GADO  = (PEFIS_DB_GADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_DB_OUTRO  = (PEFIS_DB_OUTRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_DB_DIV_BCRA  = (PEFIS_DB_DIV_BCRA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_DB_DIV_AVAIS  = (PEFIS_DB_DIV_AVAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_APLIC_FIN  = (PEFIS_IR_APLIC_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_QT_ACOES_EMPRS  = (PEFIS_IR_QT_ACOES_EMPRS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_IR_IMOVEIS  = (PEFIS_IR_IMOVEIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_VEICULOS  = (PEFIS_IR_VEICULOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_EMP_TERCEIRO  = (PEFIS_IR_EMP_TERCEIRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_IR_OUTRO  = (PEFIS_IR_OUTRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_TOTAL_BENS_DIRT  = (PEFIS_IR_TOTAL_BENS_DIRT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_DIV_ONUS  = (PEFIS_IR_DIV_ONUS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_IR_DIV_AVAIS  = (PEFIS_IR_DIV_AVAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_IR_PATR_LIQ  = (PEFIS_IR_PATR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_ARREC  = (PEFIS_ARREC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_ARDESP  = (PEFIS_ARDESP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_ARRESULT  = (PEFIS_ARRESULT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),PEFIS_ARBENS_ATIV_RURAL  = (PEFIS_ARBENS_ATIV_RURAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "PEFIS_ARDIV_VIN_ATIV_RURAL  = (PEFIS_ARDIV_VIN_ATIV_RURAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_DISP  = (SEGUR_DISP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_CRED_OPER_PREVID_COMPL  = (SEGUR_CRED_OPER_PREVID_COMPL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_SEGURADORAS  = (SEGUR_SEGURADORAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_IRB  = (SEGUR_IRB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_DESP_COMERC_DIFERD  = (SEGUR_DESP_COMERC_DIFERD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_TITULO_VL_MBLRO  = (SEGUR_TITULO_VL_MBLRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_DESP_PAGTO_ANTCPO  = (SEGUR_DESP_PAGTO_ANTCPO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_OUTRA_CONTA_OPER  = (SEGUR_OUTRA_CONTA_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_OUTRA_CONTA_NAO_OPER  = (SEGUR_OUTRA_CONTA_NAO_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_ATIV_CIRC  = (SEGUR_ATIV_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_APLIC  = (SEGUR_APLIC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_TITULO_CRED_RECEB  = (SEGUR_TITULO_CRED_RECEB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_REALZV_LP  = (SEGUR_REALZV_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PART_CTRL_COLGD  = (SEGUR_PART_CTRL_COLGD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_OUTRO_INVTMO  = (SEGUR_OUTRO_INVTMO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_INVTMO  = (SEGUR_INVTMO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_IMBRO_TECN_LIQ  = (SEGUR_IMBRO_TECN_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_ATIV_DFRD  = (SEGUR_ATIV_DFRD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_ATIV_PERMAN  = (SEGUR_ATIV_PERMAN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_ATIV_TOTAL  = (SEGUR_ATIV_TOTAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_DEB_OPER_PREVID  = (SEGUR_DEB_OPER_PREVID *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_OBRIG_SOC_TRIB  = (SEGUR_OBRIG_SOC_TRIB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_SINIS_LIQ  = (SEGUR_SINIS_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_EMPREST_FIN  = (SEGUR_EMPREST_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PROV_TECN  = (SEGUR_PROV_TECN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_DEPOS_TERC  = (SEGUR_DEPOS_TERC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_CTRL_COLGD  = (SEGUR_CTRL_COLGD *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PASV_CIRC  = (SEGUR_PASV_CIRC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_OUTRA_CONTA  = (SEGUR_OUTRA_CONTA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_EXIG_LP  = (SEGUR_EXIG_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_RES_EXERC_FUT  = (SEGUR_RES_EXERC_FUT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_CAPITAL_SOC  = (SEGUR_CAPITAL_SOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_RES_CAPITAL_LCR  = (SEGUR_RES_CAPITAL_LCR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_RES_REAVAL  = (SEGUR_RES_REAVAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PARTICIP_MNTRO  = (SEGUR_PARTICIP_MNTRO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_LCR_PREJ_ACUM  = (SEGUR_LCR_PREJ_ACUM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PATR_LIQ  = (SEGUR_PATR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PASV_TOTAL  = (SEGUR_PASV_TOTAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_RENDA_CONTRIB  = (SEGUR_RENDA_CONTRIB *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_CONTRIB_RPS  = (SEGUR_CONTRIB_RPS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_VAR_PROV_PREMIOS  = (SEGUR_VAR_PROV_PREMIOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_REC_OPER_LIQ  = (SEGUR_REC_OPER_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_DESP_BENEF_RESGT  = (SEGUR_DESP_BENEF_RESGT *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_VAR_PROV_EVENTO_NAO_AVIS  = (SEGUR_VAR_PROV_EVENTO_NAO_AVIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_LCR_BRUTO  = (SEGUR_LCR_BRUTO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_DESP_ADM  = (SEGUR_DESP_ADM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_DESP_VDA  = (SEGUR_DESP_VDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_OUTRO_DESP_REC_OPER  = (SEGUR_OUTRO_DESP_REC_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_SALDO_CORREC_MONET  = (SEGUR_SALDO_CORREC_MONET *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_LCR_ANTES_RES_FIN  = (SEGUR_LCR_ANTES_RES_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_RECT_FIN  = (SEGUR_RECT_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_DESP_FIN  = (SEGUR_DESP_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_REC_DESP_NAO_OPER  = (SEGUR_REC_DESP_NAO_OPER *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_LCR_ANTES_EQUIV_PATRIM  = (SEGUR_LCR_ANTES_EQUIV_PATRIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_EQUIV_PATRIM  = (SEGUR_EQUIV_PATRIM *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_LCR_ANTES_IR  = (SEGUR_LCR_ANTES_IR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_IR_RENDA_CONTRIB_SOC  = (SEGUR_IR_RENDA_CONTRIB_SOC *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),SEGUR_PARTICIP  = (SEGUR_PARTICIP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "SEGUR_LCR_LIQ  = (SEGUR_LCR_LIQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_DISPNVL  = (OP_DISPNVL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_CRED_A_CP  = (OP_CRED_A_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP  = (OP_ATV_CIRC_DEMAIS_CRED_VLRS_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ATIVO_INVESTIMENTOS  = (OP_ATIVO_INVESTIMENTOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ATIVO_CIRC_ESTOQ  = (OP_ATIVO_CIRC_ESTOQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP  = (OP_ATIVO_CIRC_VPD_PAGAS_ANTECIP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_CRED_A_LP  = (OP_CRED_A_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP  = (OP_ATV_RLZ_DEMAIS_CRED_VLRS_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_INVESTIMENTOS  = (OP_INVESTIMENTOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ATV_RLZ_ESTOQ  = (OP_ATV_RLZ_ESTOQ *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ATV_RLZ_VPD_PAGAS_ANTECIP  = (OP_ATV_RLZ_VPD_PAGAS_ANTECIP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_IMOBILIZADO  = (OP_IMOBILIZADO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_INTANGIVEL  = (OP_INTANGIVEL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP  = (OP_PASS_CIRC_OB_TRAB_PREV_ASS_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_EMPREST_FINAN_CP  = (OP_EMPREST_FINAN_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_FORN_CTAS_PG_CP  = (OP_FORN_CTAS_PG_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_OBRIG_FISCAIS_CP  = (OP_OBRIG_FISCAIS_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_OBRIG_REPART  = (OP_OBRIG_REPART *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_PROV_CP  = (OP_PROV_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_DEMAIS_OBRIG_CP  = (OP_DEMAIS_OBRIG_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_PASS_N_CIRC_OB_TRB_PREV_AS_CP  = (OP_PASS_N_CIRC_OB_TRB_PREV_AS_CP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_EMPREST_FINANC_LP  = (OP_EMPREST_FINANC_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_FORNECEDORES_LP  = (OP_FORNECEDORES_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_PREVISOES_LP  = (OP_PREVISOES_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_DEMAIS_OBRIG_LP  = (OP_DEMAIS_OBRIG_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_PATRIMONIO_LP  = (OP_PATRIMONIO_LP *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_TRIBUTARIAS  = (OP_TRIBUTARIAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_CONTRIBUICOES  = (OP_CONTRIBUICOES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_TRANSF_CORRENTES  = (OP_TRANSF_CORRENTES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_PATRIMONIAIS  = (OP_PATRIMONIAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_OUTRAS_RECEITAS_CORRENTES  = (OP_OUTRAS_RECEITAS_CORRENTES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_DEDUCOES  = (OP_DEDUCOES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_PESSOAL_ENCARGOS_SOCIAIS  = (OP_PESSOAL_ENCARGOS_SOCIAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_JUROS_ENCARGOS_DIVIDAS  = (OP_JUROS_ENCARGOS_DIVIDAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_TRANSFERENCIAS_CORRENTES  = (OP_TRANSFERENCIAS_CORRENTES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_OUTRAS_DESPESAS_CORRENTES  = (OP_OUTRAS_DESPESAS_CORRENTES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_OPERACOES_CREDITO  = (OP_OPERACOES_CREDITO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ALIENACAO_BENS  = (OP_ALIENACAO_BENS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_TRANSFERENCIA_CAPITAL  = (OP_TRANSFERENCIA_CAPITAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_RECEITA_CAPITAL_OUTRAS  = (OP_RECEITA_CAPITAL_OUTRAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_INVERSOES_FINANCEIRAS  = (OP_INVERSOES_FINANCEIRAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_AMORTIZACAO_DIVIDA  = (OP_AMORTIZACAO_DIVIDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_OUTRAS_DESPESAS_CAPITAL  = (OP_OUTRAS_DESPESAS_CAPITAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_OUTRAS_RECEITAS_DESPESAS  = (OP_OUTRAS_RECEITAS_DESPESAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_RESERVAS_CONTINGENCIAS  = (OP_RESERVAS_CONTINGENCIAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_TRIBUTARIAS  = (OP_ORCADO_TRIBUTARIAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_CONTRIBUICOES  = (OP_ORCADO_CONTRIBUICOES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_TRANSF_CORRENTES  = (OP_ORCADO_TRANSF_CORRENTES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_PATRIMONIAIS  = (OP_ORCADO_PATRIMONIAIS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_OUTRAS_RECT_CORREN  = (OP_ORCADO_OUTRAS_RECT_CORREN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_DEDUCOES  = (OP_ORCADO_DEDUCOES *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_PESS_ENCG_DIV  = (OP_ORCADO_PESS_ENCG_DIV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_JUROS_ENCARG_DIV  = (OP_ORCADO_JUROS_ENCARG_DIV *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_TRANSF_CORR  = (OP_ORCADO_TRANSF_CORR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_OUTRAS_DESP_CORR  = (OP_ORCADO_OUTRAS_DESP_CORR *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_OPER_CREDITO  = (OP_ORCADO_OPER_CREDITO *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_ALIENACAO_BENS  = (OP_ORCADO_ALIENACAO_BENS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_TRANSF_CAPITAL  = (OP_ORCADO_TRANSF_CAPITAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_RECT_CAPITAL_OUTRAS  = (OP_ORCADO_RECT_CAPITAL_OUTRAS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_INVESTIMENTOS  = (OP_ORCADO_INVESTIMENTOS *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_INVERSOES_FIN  = (OP_ORCADO_INVERSOES_FIN *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),"
qry = qry & "OP_ORCADO_AMORT_DIVIDA  = (OP_ORCADO_AMORT_DIVIDA *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2) ),OP_ORCADO_OUTRAS_DESP_CAPITAL  = (OP_ORCADO_OUTRAS_DESP_CAPITAL *  input(TRANWRD('" & cotacao & " ', ',' , '.'), 4.2))"
qry = qry & " where dt_exerc = '" & per & "' and cd_cli =" & cd_cli

    updateMoeda = qry
    
End Function

Function GetIPAddress()
    Const strComputer As String = "."
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress & "," & Join(IPAddress, ", ")
        End If
    Next
    GetIPAddress = Trim(Split(strIPAddress, ",")(3))
End Function



Public Function getMachineName() As String
    getMachineName = Environ$("computername")
End Function


