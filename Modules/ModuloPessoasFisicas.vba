Public Sub Planilha_PF()


Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1(6)
Dim colBalanco As Collection
Dim periodos As String
Dim count As Integer
Dim countR2 As Integer
Dim countFor As Integer

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
        
        .PEFIS_FI_PATR_COMPROVADO = rs![PEFIS_FI_PATR_COMPROVADO]
        .PEFIS_FI_LIQ = rs![PEFIS_FI_LIQ]
        .PEFIS_FI_ATIV_IMBLZ = rs![PEFIS_FI_ATIV_IMBLZ]
        .PEFIS_FI_PARTICIP_EMP = rs![PEFIS_FI_PARTICIP_EMP]
        .PEFIS_FI_GADO = rs![PEFIS_FI_GADO]
        .PEFIS_FI_OUTRO = rs![PEFIS_FI_OUTRO]
        .PEFIS_FI_DIV_BCRA = rs![PEFIS_FI_DIV_BCRA]
        .PEFIS_FI_DIV_AVAIS = rs![PEFIS_FI_DIV_AVAIS]
 
        .PEFIS_IR_APLIC_FIN = rs![PEFIS_IR_APLIC_FIN]
        .PEFIS_IR_QT_ACOES_EMPRS = rs![PEFIS_IR_QT_ACOES_EMPRS]
        .PEFIS_IR_IMOVEIS = rs![PEFIS_IR_IMOVEIS]
        .PEFIS_IR_VEICULOS = rs![PEFIS_IR_VEICULOS]
        .PEFIS_IR_EMP_TERCEIRO = rs![PEFIS_IR_EMP_TERCEIRO]
        .PEFIS_IR_OUTRO = rs![PEFIS_IR_OUTRO]
        .PEFIS_IR_DIV_ONUS = rs![PEFIS_IR_DIV_ONUS]
        .PEFIS_IR_DIV_AVAIS = rs![PEFIS_IR_DIV_AVAIS]

        .PEFIS_ARREC = rs![PEFIS_ARREC]
        .PEFIS_ARDESP = rs![PEFIS_ARDESP]

        .PEFIS_ARBENS_ATIV_RURAL = rs![PEFIS_ARBENS_ATIV_RURAL]
        .PEFIS_ARDIV_VIN_ATIV_RURAL = rs![PEFIS_ARDIV_VIN_ATIV_RURAL]

     End With
     
     colBalanco.Add Item:=blc
    
     rs.MoveNext
     Loop Until rs.EOF
     
End If
rs.Close
conn.Close

countR2 = 3

'contador de iteraçoes do For Each
countFor = 1

'ALTERAÇÃO FEITA 27/05: preenchendo células, +alteração no count, +alteração na Sheet

For Each blc In colBalanco
     
    ActiveWorkbook.Worksheets("PF").Cells(14, count2).Value = blc.PEFIS_FI_PATR_COMPROVADO
    ActiveWorkbook.Worksheets("PF").Cells(15, count2).Value = blc.PEFIS_FI_LIQ
    ActiveWorkbook.Worksheets("PF").Cells(16, count2).Value = blc.PEFIS_FI_ATIV_IMBLZ
    ActiveWorkbook.Worksheets("PF").Cells(17, count2).Value = blc.PEFIS_FI_PARTICIP_EMP
    ActiveWorkbook.Worksheets("PF").Cells(18, count2).Value = blc.PEFIS_FI_GADO
    ActiveWorkbook.Worksheets("PF").Cells(19, count2).Value = blc.PEFIS_FI_OUTRO
    ActiveWorkbook.Worksheets("PF").Cells(20, count2).Value = blc.PEFIS_FI_DIV_BCRA
    ActiveWorkbook.Worksheets("PF").Cells(21, count2).Value = blc.PEFIS_FI_DIV_AVAIS

    ActiveWorkbook.Worksheets("PF").Cells(27, count2).Value = blc.PEFIS_IR_APLIC_FIN
    ActiveWorkbook.Worksheets("PF").Cells(28, count2).Value = blc.PEFIS_IR_QT_ACOES_EMPRS
    ActiveWorkbook.Worksheets("PF").Cells(29, count2).Value = blc.PEFIS_IR_IMOVEIS
    ActiveWorkbook.Worksheets("PF").Cells(30, count2).Value = blc.PEFIS_IR_VEICULOS
    ActiveWorkbook.Worksheets("PF").Cells(31, count2).Value = blc.PEFIS_IR_EMP_TERCEIRO
    ActiveWorkbook.Worksheets("PF").Cells(32, count2).Value = blc.PEFIS_IR_OUTRO

    ActiveWorkbook.Worksheets("PF").Cells(34, count2).Value = blc.PEFIS_IR_DIV_ONUS
    ActiveWorkbook.Worksheets("PF").Cells(35, count2).Value = blc.PEFIS_IR_DIV_AVAIS

    ActiveWorkbook.Worksheets("PF").Cells(39, count2).Value = blc.PEFIS_ARREC
    ActiveWorkbook.Worksheets("PF").Cells(40, count2).Value = blc.PEFIS_ARDESP

    ActiveWorkbook.Worksheets("PF").Cells(43, count2).Value = blc.PEFIS_ARBENS_ATIV_RURAL
    ActiveWorkbook.Worksheets("PF").Cells(44, count2).Value = blc.PEFIS_ARDIV_VIN_ATIV_RURAL

    countR2 = countR2 + 1 'D

    countFor = countFor + 1
    
Next

colBalanco.count




End Sub
