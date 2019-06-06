
Public Function getConnection() As ADODB.Connection
    Dim obConnection2 As ADODB.Connection
    Set obConnection2 = New ADODB.Connection
    'obConnection2.ConnectionString = "Provider=SAS.IOMProvider;Data Source=iom-bridge://localhost:8591;User ID=fabio;Password=1234"
    obConnection2.ConnectionString = "Provider=SAS.IOMProvider;Data Source=""iom://svpsas06.abcbrasil.local:8591;Bridge;SECURITYPACKAGE=Negotiate"""
    Set getConnection = obConnection2
End Function

Public Function getLoginWindows() As String
    getLoginWindows = (Environ$("Username"))
End Function

Public Sub LimpaAux()

    Dim Lin As Integer
    Dim Col As Integer

    Col = 1
    Do While Col <= 552
        Lin = 2
        Do While Lin <= 8
            Planilha2.Cells(Lin, Col) = 0
            Planilha2.Cells(Lin, Col).Interior.Color = vbWhite
            Lin = Lin + 1
        Loop
        Col = Col + 1
    Loop
    
    Lin = 2
    Do While Lin <= 8
        Planilha2.Cells(Lin, 3) = ""
        Planilha2.Cells(Lin, 7) = ""
        Planilha2.Cells(Lin, 8) = ""
        Planilha2.Cells(Lin, 9) = ""
        Planilha2.Cells(Lin, 12) = ""
        Lin = Lin + 1
    Loop

End Sub

Public Sub GravaFatoBalanco()

    Dim sas As SASExcelAddIn
    Set sas = Application.COMAddIns.Item("SAS.ExcelAddIn").Object
    
    Dim list As SASStoredProcesses
    Set list = sas.GetStoredProcesses(ThisWorkbook)
    
    For i = 1 To list.count
        Dim stp As SASStoredProcess
        Set stp = list.Item(i)
        If stp.Path = "/User Folders/JhonyS/My Folder/Planilhamento" Then
            stp.Refresh
            Exit Sub
        End If
    Next i
    
    Dim prompts As SASPrompts
    Set prompts = sas.CreateSASPromptsObject
    prompts.Add "ProjectName", "'" + cd_grupo + " - " + cd_cli + "'"
    prompts.Add "ProjectDesk", "'" + NM_EMP + " - " + CNPJ + "'"
    prompts.Add "ID_CLIENTE", cd_cli
    prompts.Add "CNPJ", CNPJ
    prompts.Add "LAYOUT", Layout
    
    Dim inputStreams As SASRanges
    Set inputStreams = New SASRanges
    
    inputStreams.Add "planilha", Planilha2.Range("A1:UF8")
    
    'Dim stp As SASStoredProcess
    Set stp = sas.InsertStoredProcess("/User Folders/JhonyS/My Folder/Planilhamento", Sheet6.Range("A20"), prompts, , inputStreams)

End Sub

Sub Usernameincell()

    Range("S3").Value = Application.UserName

End Sub
