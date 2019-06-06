VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Front 
   Caption         =   "FGV"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9516
   OleObjectBlob   =   "Front.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Front"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lFormularioVisivel As String

Private Sub BT_Clear_Click()
'Botao de limpeza, limpar todos os campos
Dim objeto As Control

For Each objeto In Me.Controls 'faz o looping percorrendo todos os objetos do Userform1
    If TypeName(objeto) = "TextBox" Or TypeName(objeto) = "ComboBox" Or TypeName(objeto) = "ListBox" Then  ' se o tipo do objeto encontrado tiver o nome TEXTBOX
            
            Me.TextBox1_Inserirdados = "" 'limpa o campo
            Me.ListBox2.Clear
            Me.ListBox1.Clear
            Me.ComboBox2.Clear
            ComboBox2.AddItem "CNPJ/CPF"
            ComboBox2.AddItem "CRC CLIENTE"
            ComboBox2.AddItem "CRC GRUPO"
            ComboBox2.AddItem "NOME"
            ComboBox1.Clear
            ComboBox1.AddItem "EMPRESA (CNPJ)"
            ComboBox1.AddItem "PESSOA FISICA"
            ComboBox1.AddItem "SEGURADORA"
            ComboBox1.AddItem "BANCOS"
            ComboBox1.AddItem "ORGAO PUBLICO"
            
     End If
Next objeto

End Sub

Private Sub BT_Clear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

BT_Clear.BackColor = &HC0C0C0
'BT_Clear.MousePointer = fmMousePointerHourGlass
'BT_Clear.ForeColor = &HFFFFFF

End Sub

'Botao pesquisar,, pesquisar itens selecionados

Private Sub BT_PESQUISAR_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    
    Dim i As Integer
    Dim j As Integer
    Dim l As Integer
    Dim res As String
    Dim res2 As String
    

    variable = ""
    variable2 = Me.TextBox1_Inserirdados
    habilitaConsulta = False
    
    
    Select Case ComboBox2.ListIndex
    Case Is = 0
        variable = "CNPJ"
        variable2 = Right("000000000000000" & variable2, 15)
    Case Is = 1
        variable = "CD_CLI"
    Case Is = 2
        variable = "CD_GRP"
    Case Is = 3
        variable = "NM_EMP"
    End Select
    
       
    Set conn = getConnection()
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
        conn.Open
        If ComboBox2.ListIndex = 3 Then
            qry = "select CD_CLI,CD_GRP,FLG_GRP,NM_EMP,DT_EXERC from LB_PLANI.DIM_GRP_CLI where " & variable & " like  '%" & UCase(variable2) & "%'"
        ElseIf ComboBox2.ListIndex = 1 Then
            qry = "select CD_CLI,CD_GRP,FLG_GRP,NM_EMP,DT_EXERC from LB_PLANI.DIM_GRP_CLI where " & variable & " = " & variable2
        Else
            qry = "select CD_CLI,CD_GRP,FLG_GRP,NM_EMP,DT_EXERC from LB_PLANI.DIM_GRP_CLI where " & variable & " = '" & variable2 & "'"
        End If
        
        rs.Open qry, conn, adOpenStatic
                
        If rs.RecordCount > 0 Then
            With Me.ListBox2
                   .Clear
                Do
                    .ColumnCount = 5
                    .ColumnWidths = "60;60;20;230,40"
                    .AddItem
                    .list(j, 0) = rs![cd_cli]
                    .list(j, 1) = rs![CD_GRP]
                    .list(j, 2) = rs![FLG_GRP]
                    .list(j, 3) = rs![NM_EMP]
                    .list(j, 4) = rs![DT_EXERC]
                    
                    NM_EMP = rs![NM_EMP]
                    
                    j = j + 1
                    
                    rs.MoveNext
                    
                Loop Until rs.EOF
            End With
        Else
            Me.ListBox2.Clear
            MsgBox "Nenhum resultado encontrado para essa pesquisa"
        End If
         
   End If
   
UserForm_Initialize_Exit:

    On Error Resume Next
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Exit Sub
    
UserForm_Initialize_Err:

    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error!"
    Resume UserForm_Initialize_Exit
            
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

End Sub

Private Sub BT_PESQUISAR_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Deixar verde quando passar o mouse
'BT_PESQUISAR.BackColor = &HC0C0C0
'BT_PESQUISAR.MousePointer = fmMousePointerHourGlass
'BT_PESQUISAR.ForeColor = &HFFFFFF

Me.BT_PESQUISAR.BackColor = &HC0C0C0

End Sub

Private Sub BT_Planilhar_Click()
'
    Dim strPesquisa As String
    Dim count As Integer
    Dim iCtr As Long
    Dim arr1(6)
    
    LimpaAux
    Limpa_Planilha_PDD
    Limpa_Planilha_Funding
    Limpa_Planilha_Contingencias
    Limpa_Planilha_Bancos_Mil
    Limpa_Planilha_Carteira
    Limpa_Planilha_Rentabilidada
    Limpa_Planilha_TVM
    
    Dim ic As Integer
    For ic = 1 To ActiveWorkbook.Sheets.count
        ActiveWorkbook.Sheets(ic).Visible = True
    Next ic
    
'    ThisWorkbook.Activate
'    ActiveWindow.Visible = True
    
    If IsNull(Me.ListBox2.Value) Or Me.ListBox2.Value = "" Then
        MsgBox "Favor selecionar cliente"
        Exit Sub
    End If
    
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
    
    If IsNull(ComboBox1.Text) Or ComboBox1.Text = "" Then
        MsgBox "Favor selecionar um layout"
        Exit Sub
    End If
    Layout = Front.ComboBox1.Text
    
    qry = "select layout_final from lb_plani.dim_grp_cli where cd_cli = " & cd_cli

    Set conn = getConnection()
    Set rs = New ADODB.Recordset
    
    conn.Open
    rs.Open qry, conn, adOpenStatic
    
        layout_final = rs![layout_final]
 
    rs.Close
    conn.Close
    
    If layout_final <> Layout Then
        result = MsgBox("Layout diferente do layout anterior ", vbYesNo + vbExclamation)
        If result = vbNo Then
            Exit Sub
        End If
    End If

    If ComboBox1.Text = "Banco" Then
        Planilha_Bancos
        
        strPesquisa = "BANCOS"
    ElseIf ComboBox1.Text = "Empresas" Then
        Planilha_PJ_ReaisMil
        Planilha_PJ_Fluxo
        
        strPesquisa = "PJ"
    ElseIf ComboBox1.Text = "Orgãos Públicos" Then
        Planilha_OP_ReaisMil

        strPesquisa = "OP"
    ElseIf ComboBox1.Text = "Pessoas Físicas" Then
        Planilha_PF
        
        strPesquisa = "PF"
    ElseIf ComboBox1.Text = "Seguradora" Then
        Planilha_SEGURADORA_ReaisMil
        
        strPesquisa = "SEGURADORA"
    End If
 
    'Application.Visible = True
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Sheets.count
        If InStr(1, ActiveWorkbook.Sheets(i).Name, strPesquisa) = 0 Then
            ActiveWorkbook.Sheets(i).Visible = False
        Else
            ActiveWorkbook.Sheets(i).Visible = True
        End If
    Next i
    'ThisWorkbook.Sheets("Aux").Visible = True
    ThisWorkbook.Sheets("VALIDAÇÃO").Visible = True

    Unload Front

End Sub

Private Sub BT_Planilhar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    BT_Planilhar.BackColor = &HC0C0C0
    'BT_Planilhar.MousePointer = fmMousePointerHourGlass
    'BT_Planilhar.ForeColor = &HFFFFFF

End Sub

Private Sub CommandButton1_CONSULTAR_PLAN_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    CommandButton1_CONSULTAR_PLAN.BackColor = &HC0C0C0
    'CommandButton1_CONSULTAR_PLAN.MousePointer = fmMousePointerHourGlass
    'CommandButton1_CONSULTAR_PLAN.ForeColor = &HFFFFFF

End Sub

Private Sub UserForm_Activate()
    
'    Application.Visible = False
    
    ComboBox1.AddItem "Banco"
    ComboBox1.AddItem "Empresas"
    ComboBox1.AddItem "Orgãos Públicos"
    ComboBox1.AddItem "Pessoas Físicas"
    ComboBox1.AddItem "Seguradora"
'
    alimenta_combobox
'    LimpaAux
'    Limpa_Planilha_PDD
'    Limpa_Planilha_Funding
'    Limpa_Planilha_Contingencias
'    Limpa_Planilha_Bancos_Mil
'
'    Limpa_Planilha_Carteira
'    Limpa_Planilha_Rentabilidada
'    Limpa_Planilha_TVM

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    BT_PESQUISAR.BackColor = &H8000000F
    BT_PESQUISAR.ForeColor = &H80000012
    
    CommandButton1_CONSULTAR_PLAN.BackColor = &H8000000F
    CommandButton1_CONSULTAR_PLAN.ForeColor = &H80000012
    
    BT_Clear.BackColor = &H8000000F
    BT_Clear.ForeColor = &H80000012
    
    BT_Planilhar.BackColor = &H8000000F
    BT_Planilhar.ForeColor = &H80000012

End Sub


Private Sub CommandButton1_CONSULTAR_PLAN_Click()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Me.ListBox1.MultiSelect = 1
    codCliente = Me.ListBox2.Value

    Set conn = getConnection()
            
    Set rs = New ADODB.Recordset
    
    If IsNull(codCliente) Or codCliente = "" Then
        MsgBox "Favor selecionar cliente"
    Else
        conn.Open
        
        qry = "select Distinct DT_EXERC, CD_CLI, Max(DT_CRG) As DT_CRG from LB_PLANI.FATO_BALANCO where cd_cli =  " & codCliente & " Order By DT_EXERC "
        
        rs.Open qry, conn, adOpenStatic
          
        If Not IsNull(rs) And rs.RecordCount > 0 Then
            With Me.ListBox1
                .Clear
                Do
                    .ColumnCount = 1
                    .ColumnWidths = "60"
                    .AddItem
                    .list(l, 0) = rs![DT_EXERC] & " - " & rs![cd_cli]
                    
                    l = l + 1
                    
                    rs.MoveNext
                Loop Until rs.EOF
            End With
        Else
            Me.ListBox1.Clear
            MsgBox "Nenhum resultado encontrado para essa pesquisa"
        End If
         
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
            
    obRecordset2.Close
    Set obRecordset2 = Nothing
    obConnection2.Close
    Set obConnection2 = Nothing
End Sub

Private Sub UserForm_Initialize()

'ComboBox1.AddItem "EMPRESA (CNPJ)"
'ComboBox1.AddItem "PESSOA FISICA"
'ComboBox1.AddItem "SEGURADORA"
'ComboBox1.AddItem "BANCOS"
'ComboBox1.AddItem "ORGAO PUBLICO"

ComboBox2.AddItem "CNPJ/CPF"
ComboBox2.AddItem "CRC CLIENTE"
ComboBox2.AddItem "CRC GRUPO"
ComboBox2.AddItem "NOME"

End Sub





